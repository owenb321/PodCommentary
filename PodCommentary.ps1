#########
#  AddPodcastCommentary
#
#  Features:
#  - Mixes in commentary podcast audio alongside movie audio with auto-ducking
#  - Adds a title card image before and after the movie
#  - Preserves the original movie audio in a second audio track
#  - By default, copies the original video along with AAC audio for the commentary
#      and AC3 for the movie audio in an mkv container
#
#  Notes:
#  - Will use embedded subs from the video file if present, with the option to provide an external file
#
param (
    [Parameter(Mandatory=$true)]
    $movieStartTime,
    [Parameter(Mandatory=$true)]
    $inputPodcast,
    [Parameter(Mandatory=$true)]
    $inputMovie,
    $inputSubs,
    $coverImg,
    $advancedSettingsFile,
    $outputDirectory = $PSScriptRoot,
    [int]$audioStreamIndex = 0,
    $outputVideoCodec,
    $outputCommentaryCodec = "-c:a:0 aac -ac:a:0 2",
    $outputMovieAudioCodec,
    [switch]$reencodeSourceAudio,
    [switch]$useLoudnorm,
    [switch]$omitMovieAudioTrack,
    [switch]$disableBackgroundAudio,
    $ffmpegPath = "ffmpeg",
    $ffprobePath = "ffprobe"
)

function progressPercentage($current, $total) {
    if($total -le 0) { return 0 }
    elseif ($current -gt $total) { return 100 }
    else { return ($current/$total)*100 }
}

function runFfmpeg($ffmpegArguments, [timespan]$totalTime, $totalFrames, $totalBytes, $activity = "FFMPEG", $writeProgressId = 1) {
    # Hide the banner if not already specified
    if($ffmpegArguments -notcontains '-hide_banner') {
        $ffmpegArguments = @('-hide_banner') + $ffmpegArguments
    }    
    $p = New-Object System.Diagnostics.Process
    $p.StartInfo.Filename = $ffmpegPath
    $p.StartInfo.Arguments = $ffmpegArguments
    $p.StartInfo.WorkingDirectory = $workingDir
    $p.StartInfo.UseShellExecute = $false
    $p.StartInfo.RedirectStandardError = $true
    $p.StartInfo.RedirectStandardOutput = $true
    $p.StartInfo.CreateNoWindow = $true
    $p.Start() | Out-Null

    $startTime = Get-Date
    $perc = 0
    $stderrCapture = @()

    while (-not $p.HasExited) {
        if (-not [console]::IsInputRedirected -and [console]::KeyAvailable)
        {
            $q = [System.Console]::ReadKey($true)
            switch ($q.key)
            {
                "q" {
                    $p.Kill()
                    Write-Host "Encoding stopped"
                    return 0
                }
            }
        }
        if ($p.StandardError.Peek()) {
            $line = $p.StandardError.ReadLineAsync().Result
            $stderrCapture += $line
            if ($line) {
                $info = $line.Trim();
                if ($info.StartsWith("frame=") -or $info.StartsWith("size=")) {
                    if($info -match '(?:frame=\s*(\d+) fps=\s*(\d+.*?) q=\s*(-?\d+.*?) )?size=\s*(\d+.*?) time=(\d+\:\d+\:\d+\.\d+) bitrate=\s*(\d+.*?) speed=\s*(\d+.*?)x') {
                        #$speed      = $Matches[7]
                        #$bitrate    = $Matches[6]
                        $outputTime = $Matches[5]
                        $outputSize = $Matches[4]
                        #$qp         = $Matches[3]
                        $fps        = $Matches[2]
                        $curFrame   = $Matches[1]
                        #$wholeMatch = $Matches[0]

                        if($totalFrames) {
                            $perc = progressPercentage ([int]$curFrame) $totalFrames
                        } elseif ($totalTime) {
                            $perc = progressPercentage ([timespan]$outputTime) $totalTime
                        } elseif($totalBytes) {
                            $perc = progressPercentage (Invoke-Expression $outputSize) (Invoke-Expression $totalBytes)
                        }
                        $progressParams = @{
                            Id = $writeProgressId
                            Activity = $activity
                            Status = "$outputTime - $outputSize"
                        }
                        if($curFrame -and $fps) {
                            $progressParams.Status = "$fps fps - Frame: $curFrame - " + $progressParams.Status
                        }
                        # Add in percentage & time estimates to the progress bar
                        if ($totalFrames -or $totalTime -or $totalBytes) {
                            $progressParams += @{ PercentComplete = $perc }
                            $progressParams.Status = "$([Math]::round($perc, 1))% - " + $progressParams.Status
                            if($perc -gt 0) {
                                $secondsPerPercentage = [float]((Get-Date) - $startTime).TotalSeconds / $perc
                                $secondsRemaining = (100 - $perc) * $secondsPerPercentage
                                $progressParams += @{ SecondsRemaining = $secondsRemaining }
                            }
                        }
                        Write-Progress @progressParams
                    }
                }
            }
        }
    }
    $p.WaitForExit()
    # Check for error and display it
    if($p.ExitCode -ne 0) {
        Write-Host "FFMPEG exited with error:"
        Write-Host $p.StandardOutput.ReadToEnd()
        Write-Host $stderrCapture
        Write-Host $p.StandardError.ReadToEnd()
    }
    Write-Progress -Id $writeProgressId -Activity $activity -Completed
    return $p.ExitCode
}

# Generate intro/outro files
function generateCover([timespan]$span, $type, $imagePath, $width, $height, $sar, $framerate, $timebase, $videoCodec) {
    $frameLengthMs = [int](1/[double](Invoke-Expression $framerate)*1000)
    $finalPath = join-path $workingDir "$($type).mkv"
    # Generate blank frames if no image or timespan is too short
    if(-not $imagePath -or -not (test-path $imagePath) -or $span.TotalSeconds -lt 10) {
        $ffinput = "-f", "lavfi", "-r", "$framerate", "-i", "color=c=black:s=$($videoInfo.width)x$($videoInfo.height):duration=1ms"
    } else {
        $ffinput = "-f", "image2", "-r", "$framerate", "-i", "`"$imagePath`""
    }
    # Just encode the whole segment if it is short enough
    if($span.TotalSeconds -lt 10) {
        $spanMs = $span.TotalMilliseconds
        if($spanMs -lt $frameLengthMs) { $spanMs = $frameLengthMs }
        runffmpeg -activity "Generating blank frames" -workingDir $workingDir `
            -ffmpegArguments (@("-y") + $ffinput + $videoCodec + @("-vf tpad=stop_duration=$($spanMs.TotalMilliseconds)ms:stop_mode=clone,settb=$timebase", "`"$finalPath`"")) | Out-Null
        Write-Progress -Activity "Generating $type" -Completed
        return $finalPath
    }
    $scaleFilter = "scale=$($width):$($height):force_original_aspect_ratio=decrease," +
        "pad=$($width):$($height):-1:-1:color=black," +
        "setsar=sar=$($sar -replace '\:','/')," +
        "settb=$timebase"

    Write-Progress -Activity "Generating $type" -PercentComplete 0
    $segmentPath = join-path $workingDir "$($type)_cover_10s.mkv"
    runffmpeg -activity "Generating cover segment" -workingDir $workingDir `
        -ffmpegArguments (@("-y") + $ffinput + @("-vf", "$scaleFilter,tpad=stop_duration=$(10000-$frameLengthMs)ms:stop_mode=clone") + $videoCodec + @("`"$segmentPath`"")) | Out-Null

    Write-Progress -Activity "Generating $type" -PercentComplete 20
    $fadeInPath = join-path $workingDir "$($type)_cover_fade_in.mkv"
    runffmpeg -activity "Generating fade in" -workingDir $workingDir `
        -ffmpegArguments (@("-y") + $ffinput + @("-t 00:00:03", "-vf", "$scaleFilter,tpad=stop_duration=$(3000-$frameLengthMs)ms:stop_mode=clone,fade=type=in:duration=2:start_time=1") + $videoCodec + @("`"$fadeInPath`"")) | Out-Null
    $fadeInProbe = &$ffprobePath -v quiet -print_format json -show_format -show_streams "$fadeInPath" | convertfrom-json
    $fadeInMs = ([double]$fadeInProbe.format.duration)*1000

    Write-Progress -Activity "Generating $type" -PercentComplete 40
    $midPath = join-path $workingDir "$($type)_mid.mkv"
    runffmpeg -activity "Looping cover stream" -workingDir $workingDir `
        -ffmpegArguments @("-y", "-stream_loop", ([math]::Ceiling($span.TotalSeconds/10)), "-i", "`"$segmentPath`"", "-t", $span.Add('-00:00:06'), "-c copy", "`"$midPath`"") | Out-Null
    $midProbe = &$ffprobePath -v quiet -print_format json -show_format -show_streams "$midPath" | convertfrom-json
    $introMidMs = ([double]$midProbe.format.duration)*1000
    $fadeOutMs = $span.TotalMilliseconds - $fadeInMs - $introMidMs # - ($frameLengthMs*2)

    Write-Progress -Activity "Generating $type" -PercentComplete 60
    $fadeOutPath = join-path $workingDir "$($type)_fade_out.mkv"
    runffmpeg -activity "Generating fade out" -workingDir $workingDir `
        -ffmpegArguments (@("-y") + $ffinput + @("-vf", "$scaleFilter,tpad=stop_duration=$($fadeOutMs-$frameLengthMs)ms:stop_mode=clone,fade=type=out:duration=2") + $videoCodec + @("`"$fadeOutPath`"")) | Out-Null

    Write-Progress -Activity "Generating $type" -PercentComplete 80
    $concatTxtPath = Join-Path $workingDir "$($type)_concat.txt"
    @("file '$fadeInPath'", "file '$midPath'", "file '$fadeOutPath'") | set-content "$concatTxtPath"
    runffmpeg -activity "Joining all $type clips" -workingDir $workingDir `
        -ffmpegArguments (@("-y", "-f concat", "-safe 0", "-i `"$concatTxtPath`"", "-c:v copy `"$finalPath`"")) | Out-Null
    Write-Progress -Activity "Generating $type" -Completed
    
    #Clean up temp files
    Remove-Item (join-path $workingDir "$($type)_*")

    return $finalPath
}

function generateEditedSubs($cutPoints, $inputSubs) {
    $cutCount = $cutPoints.Count
    Write-Progress -Activity "Editing Subtitles"
    $startTimeMs = 0
    for($i=0; $i -lt $cutCount; $i++) {
        $cut = $cutPoints[$i]
        $segmentMs = $cut.cutPointMs - $startTimeMs
        # Subs report the duration as the timestamp of the last line of text
        #   Need to mux in some audio to retain the correct duration of this clip in at least one stream
        $silenceInput = @("-f lavfi", "-i anullsrc=channel_layout=mono:sample_rate=48000")
        $subSegmentPath = (Join-Path $workingDir "subs$i.mkv")
        $subSegmentTempPath = (Join-Path $workingDir "subs$i`_t.mkv")
        $mapOutput = @("-map 0:a", "-map 1:s", "-c:s copy", "-c:a pcm_u8", "-ar 8000", "-ac 1")
        if(runFfmpeg -ffmpegArguments (@("-y") + $silenceInput + @("-i `"$inputSubs`"", "-ss `"$($startTimeMs)ms`"", "-t `"$($segmentMs)ms`"") + $mapOutput + @("`"$subSegmentPath`"")) `
            -activity "Cutting segment $($i+1)" -workingDir $workingDir) {
                Write-Host "Unable to edit subtitle file!"
                return $null
        }
        # Check if last subtitle overruns the clip duration and correct timing if so
        $subprobe = (&$ffprobePath -v quiet -print_format json -show_format -show_streams "$subSegmentPath") | convertfrom-json
        $durationDiff = $subprobe.streams[0].duration_ts - $segmentMs
        if($durationDiff -gt 0) {
            move-item $subSegmentPath $subSegmentTempPath -force
            if(runFfmpeg -ffmpegArguments (@("-y", "-i `"$subSegmentTempPath`"", "-t `"$($segmentMs-$durationDiff)ms`"", "-c:s copy", "`"$subSegmentPath`"")) `
                -activity "Correcting sub overrun" -workingDir $workingDir) {
                    Write-Host "Unable to edit subtitle file!"
                    return $null
            }
            remove-item $subSegmentTempPath
        }
        # Apply edit timing
        if ($cut.type -eq 'pause') {
            move-item $subSegmentPath $subSegmentTempPath -force
            # Extend the clip to include the pause duration
            if(runFfmpeg -ffmpegArguments (@("-y") + $silenceInput + @("-i `"$subSegmentTempPath`"", "-t `"$($segmentMs + $cut.durationMs)ms`"") + $mapOutput + @("`"$subSegmentPath`"")) `
                -activity "Extending segment $($i+1)" -workingDir $workingDir) {
                    Write-Host "Unable to edit subtitle file!"
                    return $null
            }
            remove-item $subSegmentTempPath
            $startTimeMs = $cut.cutPointMs
        }
        elseif ($cut.type -eq 'skip') {
            # Set the next start time ahead to account for the skip
            $startTimeMs = $cut.cutPointMs + $cut.durationMs
        }
    }
    # Create last segment
    $subSegmentPath = (Join-Path $workingDir "subs$i.mkv")
    if(runFfmpeg -ffmpegArguments (@("-y") + $silenceInput + @("-i `"$inputSubs`"", "-ss `"$($startTimeMs)ms`"", "-shortest") + $mapOutput + @("`"$subSegmentPath`"")) `
        -activity "Cutting segment $($i+1)" -workingDir $workingDir) {
            Write-Host "Unable to edit subtitle file!"
            return $null
    }
    # Create concat list
    $mkvList = 0..$i | ForEach-Object { Join-Path $workingDir "subs$_.mkv" }
    $subsConcatFilePath = join-path $workingDir "subs.txt"
    $subsFinalFilePath = join-path $workingDir "edited_subs.mkv"
    set-content -Path $subsConcatFilePath -Value (($mkvList | ForEach-Object { "file '$_'" }) -join "`n")
    # join all temp files
    if(runFfmpeg -ffmpegArguments @("-y", "-f concat", "-safe 0", "-i `"$subsConcatFilePath`"", "-c:s copy", "`"$subsFinalFilePath`"") `
        -activity "Joining edited subtitles" -workingDir $workingDir) {
            Write-Host "Unable to edit subtitle file!"
            return $null
    }
    Write-Progress -Activity "Editing Subtitles" -Completed
    # Clean up temp files
    remove-item $mkvList
    remove-item $subsConcatFilePath

    return $subsFinalFilePath
}

function generateKeyframeCutConcatList ($inputMovie, $cutPoints, $framerate, $videoCodec, $movieSpan) {
    Write-Progress -Id 1 -Activity "Gathering keyframe timestamps"
    # Limit the keyframe probe to just the frames around our cut points
    $keyframeSearchRange = 60 # Seconds around cut point to search for keyframes
    $keyframeList = $cutPoints | % { 
            if($_.type -eq 'skip') { 
                "$(($_.cutPointMs/1000.0)-$keyframeSearchRange)%+$($keyframeSearchRange)"
                "$(($_.cutPointMs + $_.durationMs)/1000.0)%+$($keyframeSearchRange)"
            } else {
                "$(($_.cutPointMs/1000.0)-$keyframeSearchRange)%+$($keyframeSearchRange*2)"
            }
        } | % {
            (&$ffprobePath -v error -print_format json -read_intervals "$_" `
                -skip_frame nokey -select_streams v -show_frames `
                -show_entries frame=pkt_pts_time $inputMovie | ConvertFrom-Json).frames.pkt_pts_time
        } | % {
            [int]([double]$_*1000)
        } | select -Unique
    Write-Progress -Id 1 -Activity "Gathering keyframe timestamps" -Completed

    $concatList = @()
    $startTimeMs = 0
    for($i=0; $i -lt $cutPoints.count; $i++) {
        Write-Progress -Id 1 -Activity "Reencoding cut points" -PercentComplete (progressPercentage $i ($cutPoints.count))
        $cut = $cutPoints[$i]
        $cutPoint = $cut.cutPointMs

        if($cut.type -eq 'pause') {
            # Need to pick keyframes that don't match the cut point in order to add new frames
            $startKeyframeMs = $keyframeList | ? { $_ -gt $startTimeMs -and $_ -lt $cutPoint } | select -Last 1
            $endKeyframeMs = $keyframeList | ? { $_ -gt $cutPoint } | select -First 1

            # Add keyframe-to-keyframe clip to the concat list
            $concatList += @("file '$inputMovie'", "inpoint $startTimeMs`ms", "outpoint $startKeyframeMs`ms")

            # Offset timestamps to account for seeking the input on encode
            $offsetCutPoint = $cutPoint - $startKeyframeMs
            $offsetEndKeyframeMs = $endKeyframeMs - $startKeyframeMs

            # Pad the cut point frames
            $filter =  "[0:v]split=2[v1][v2];"
            $filter += "[v1]trim=end=$($offsetCutPoint)ms"
            $filter += ",tpad=stop_duration=$($cut.durationMs)ms:stop_mode=$($cut.padMode),trim=end=$($offsetCutPoint + $cut.durationMs)ms,setpts=PTS-STARTPTS[start];"
            $filter += "[v2]trim=start=$($offsetCutPoint)ms,trim=end=$($offsetEndKeyframeMs)ms,setpts=PTS-STARTPTS[end];"
            $filter += "[start][end]concat[edited]"

            $segmentSpan = [timespan]::new(0,0,0,0,($endKeyframeMs - $startKeyframeMs + $cut.durationMs))
        }
        elseif ($cut.type -eq 'skip') {
            $startKeyframeMs = $keyframeList | ? { $_ -gt $startTimeMs -and $_ -le $cutPoint } | select -Last 1
            $resumeFrameMs = $cutPoint + $cut.durationMs
            $endKeyframeMs = $keyframeList | ? { $_ -ge $resumeFrameMs } | select -First 1

            # Add keyframe-to-keyframe clip to the concat list
            $concatList += @("file '$inputMovie'", "inpoint $startTimeMs`ms", "outpoint $startKeyframeMs`ms")

            # Offset timestamps to account for seeking the input on encode
            $offsetCutPoint = $cutPoint - $startKeyframeMs
            $offsetResumeFrameMs = $resumeFrameMs - $startKeyframeMs
            $offsetEndKeyframeMs = $endKeyframeMs - $startKeyframeMs

            if($startKeyframeMs -eq $cutPoint) {
                if($resumeFrameMs -eq $endKeyframeMs) {
                    # No encoding needed if all involved frames are keyframes
                    $startTimeMs = $endKeyframeMs
                    continue;
                }
                # Just encode the resume point frames, the beginning of the skip is handled by the keyframe edit
                $filter = "[0:v]trim=start=$($offsetResumeFrameMs)ms,trim=end=$($offsetEndKeyframeMs)ms,setpts=PTS-STARTPTS[edited]"
            } else {
                # Encode the beginning and end of the skip's frames
                $filter =  "[0:v]split=2[v1][v2];"
                $filter += "[v1]trim=end=$($offsetCutPoint)ms,setpts=PTS-STARTPTS[start];"
                $filter += "[v2]trim=start=$($offsetResumeFrameMs)ms,trim=end=$($offsetEndKeyframeMs)ms,setpts=PTS-STARTPTS[end];"
                $filter += "[start][end]concat[edited]"
            }
            $segmentSpan = [timespan]::new(0,0,0,0,($endKeyframeMs - $startKeyframeMs - $cut.durationMs))
        }
        if(-not $startKeyframeMs -or -not $endKeyframeMs) {
            throw "Unable to find keyframes!"
        }
        $segmentPath = Join-Path $workingDir "segment_$i.mkv"
        if(runFfmpeg -ffmpegArguments @("-y", "-ss $startKeyframeMs`ms", "-i `"$inputMovie`"", "-filter_complex $filter", $videoCodec, "-map [edited]", "`"$segmentPath`"") `
            -activity "Encoding Segment $i" -totalTime $segmentSpan -workingDir $workingDir -writeProgressId 2) {
                throw "Unable to encode segment!"
        }
        # Add newly encoded clip to the concat list
        $concatList += @("file '$segmentPath'")
        # Next start point is this edit's end point
        $startTimeMs = $endKeyframeMs
    }
    Write-Progress -Id 1 -Activity "Reencoding cut points" -Completed

    # Add the last segment
    $concatList += @("file '$inputMovie'", "inpoint $startTimeMs`ms")
    return $concatList
}

function generateInputList([ref][int]$offsetMovieInputNum, [ref][int]$podInputNum, [ref][int]$movieAudioInputNum, [ref][int]$movieVideoInputNum, [ref][int]$introInputNum, [ref][int]$outroInputNum, [ref][int]$subsInputNum) {
    # Gather input sources
    $inputCount = 0
    $inputs = @()

    # Include offset movie input first to output correct chapter markers (if present)
    $inputs += "-itsoffset $($introSpan.ToString()) -i `"$inputMovie`""
    $offsetMovieInputNum.Value = $inputCount++

    $inputs += "-i `"$inputPodcast`""
    $podInputNum.Value = $inputCount++

    $inputs += "-i `"$inputMovieAudio`""
    $movieAudioInputNum.Value = $inputCount++

    if($avoidVideoTranscode) {
        # Only encode intro/outro, then sandwich in the movie
        $introPath = generateCover -span $introSpan -type intro -imagePath $coverImg `
            -width $videoInfo.width -height $videoInfo.height `
            -sar $sar -framerate $framerate `
            -timebase $timebase -videoCodec $outputVideoCodec

        $concatTxtPath = Join-Path $workingDir "intro_outro_concat.txt"
        $concatFiles = @("file '$introPath'")

        # If movie requires editing, use hybrid keyframe/reencode cutting and output the video to a temp file
        if($movieCutPoints) {
            $concatFiles += generateKeyframeCutConcatList -inputMovie $inputMovie -cutPoints $movieCutPoints -framerate $framerate `
                -videoCodec $outputVideoCodec -movieSpan $movieSpan
        } else {
            $concatFiles += @("file '$inputMovie'")
        }

        # Skip outro if not needed
        if($outroSpan.TotalMilliseconds -gt $frameMs) {
            $outroPath = generateCover -span $outroSpan -type outro -imagePath $coverImg `
                -width $videoInfo.width -height $videoInfo.height `
                -sar $sar -framerate $framerate `
                -timebase $timebase -videoCodec $outputVideoCodec 
            $concatFiles += "file '$outroPath'"
        }
        $concatFiles | set-content "$concatTxtPath"

        $inputs += "-f concat -safe 0 -i `"$concatTxtPath`""
        $movieVideoInputNum.Value = $inputCount++

    } else {
        $inputs += "-i `"$inputMovie`""
        $movieVideoInputNum.Value = $inputCount++

        # Only load one frame, the rest will be cloned with tpad
        if ($coverImg) {
            $inputs += "-f image2 -r $framerate -i `"$coverImg`""
            $introInputNum.Value = $inputCount++
            if($outroSpan.TotalSeconds -gt 0) {
                $inputs += "-f image2 -r $framerate -i `"$coverImg`""
                $outroInputNum.Value = $inputCount++
            }
        } else { # Generate blank video if no cover image is present
            $inputs += "-f lavfi -r $framerate -i color=c=black:s=$($videoInfo.width)x$($videoInfo.height):duration=1ms"
            $introInputNum.Value = $inputCount++
            if($outroSpan.TotalSeconds -gt 0) {
                $inputs += "-f lavfi -r $framerate -i color=c=black:s=$($videoInfo.width)x$($videoInfo.height):duration=1ms"
                $outroInputNum.Value = $inputCount++
            }
        }
    }

    # Handle subtitle input if found
    if ($inputSubs) {
        $inputs += "-itsoffset $($introSpan.ToString()) -i `"$inputSubs`""
        $subsInputNum.Value = $inputCount++
    }
    elseif ($embeddedSubsPresent) {
        $subsInputNum.Value = $offsetMovieInputNum.Value
    }
    return $inputs
}

# Handles edits needed to either the podcast or movie to keep commentary in sync
#   cut points are provided in an advanced settings file
function generateCutFilters($cutPoints, $linkName, $inputNum, $streamType = 'a', $streamNum = 0) {
    $cutCount = $cutPoints.Count
    $filters = "[$inputNum`:$streamType`:$streamNum]"
    if($streamType -eq 'a') {
        $filters += "aresample=async=1000,"
    }
    $filters += "asplit=$($cutCount+1)" + 
        ((0..$cutCount | ForEach-Object { "[$linkName$_]" }) -join '') + ";`n"
    $startTimeMs = 0
    for($i=0; $i -lt $cutCount; $i++) {
        $cut = $cutPoints[$i]
        $filter = "[$linkName$i]atrim=start=$($startTimeMs)ms,atrim=end=$($cut.cutPointMs)ms"
        if ($cut.type -eq 'pause') {
            if($streamType -eq 'a') {
                $filter += ",apad=whole_dur=$($cut.cutPointMs - $startTimeMs + $cut.durationMs)ms"
            }
            elseif($streamType -eq 'v') {
                $filter += ",tpad=stop_duration=$($cut.durationMs)ms:stop_mode=$($cut.padMode),trim=end=$($cut.cutPointMs + $cut.durationMs)ms"
            }
            $startTimeMs = $cut.cutPointMs
        }
        elseif ($cut.type -eq 'skip') {
            $startTimeMs = $cut.cutPointMs + $cut.durationMs
        }
        $filter += ",asetpts=PTS-STARTPTS[$linkName$i`_trimmed];`n"
        $filters += $filter
    }
    $filters += "[$linkName$cutCount]atrim=start=$($startTimeMs)ms,asetpts=PTS-STARTPTS[$linkName$cutCount`_trimmed];`n"
    $filters += ((0..$cutCount | ForEach-Object { "[$linkName$_`_trimmed]" }) -join '') + 
        "concat=n=$($cutCount+1)$(if($streamType -eq 'a'){":v=0:a=1"})[$linkName];`n"
    if($streamType -eq 'v') {
        # replace audio filters with video versions
        $filters = $filters -replace 'a(trim=|split=|setpts=)','$1'
    }
    return $filters
}

function createFilterFile() {
    $filters = ""
    if($movieCutPoints.count -eq 0) {
        if(-not $avoidVideoTranscode){ 
            $filters += "[$([int]$movieVideoInputNum):v:0]null[mv];`n"
        }
        if(-not ($disableBackgroundAudio -and $omitMovieAudioTrack)) { 
            $filters += "[$([int]$movieAudioInputNum):a:$audioStreamIndex]anull[ma];`n" 
        }
    }
    else {
        if(-not $avoidVideoTranscode){
            $filters += generateCutFilters $movieCutPoints 'mv' ([int]$movieVideoInputNum) 'v' 
        }
        if(-not ($disableBackgroundAudio -and $omitMovieAudioTrack)) { 
            $filters += generateCutFilters $movieCutPoints 'ma' ([int]$movieAudioInputNum) 'a' 
        }
    }
    if($podCutPoints.count -eq 0) {
        $filters += "[$([int]$podInputNum):a:0]anull[pod];`n"
    }
    else {
        $filters += generateCutFilters $podCutPoints 'pod' ([int]$podInputNum) 'a'
    }

    # Pad out movie audio to match podcast length
    #   and vice-versa
    # Replaced adelay=all with per-channel delays for earlier ffmpeg build support
    # [ma]adelay=$($introSpan.TotalMilliseconds):all=true,apad=whole_dur=$($totalSpan.TotalMilliseconds)ms[ma_padded];
    
    # Skip all movie audio processing if it isn't needed for background audio or the second audio track
    if(-not ($disableBackgroundAudio -and $omitMovieAudioTrack)) {
        $filters += "[ma]adelay=$((1..($audioInfo.channels) | ForEach-Object { $introSpan.TotalMilliseconds }) -join '|')[ma_delayed];`n"
        $filters += "[ma_delayed]apad=whole_dur=$($totalSpan.TotalMilliseconds)ms[ma_padded];`n"
        if($omitMovieAudioTrack) { $filters += "[ma_padded]anull[ma_to_norm];`n" }
        elseif($disableBackgroundAudio) { $filters += "[ma_padded]anull[movie_audio];`n" }
        else { $filters += "[ma_padded]asplit=2[movie_audio][ma_to_norm];`n" }
    }
    $filters += "[pod]apad=whole_dur=$($totalSpan.TotalMilliseconds)ms[podstream];`n"

    if($useLoudnorm) { $normalizeFilter = 'loudnorm' }
    else { $normalizeFilter = 'dynaudnorm' }

    # No processing beyond padding if background audio is disabled
    if(-not $disableBackgroundAudio) {
        # Duck movie audio and mix it in the background of the podcast
        $filters += @"
[ma_to_norm]pan=stereo|FL=FC+0.30*FL+0.30*FLC+0.30*BL+0.30*SL+0.60*LFE|FR=FC+0.30*FR+0.30*FRC+0.30*BR+0.30*SR+0.60*LFE,$normalizeFilter,aformat=channel_layouts=stereo[ma_normalized];
[podstream]$normalizeFilter,aformat=channel_layouts=stereo[pod_normalized];
[pod_normalized]asplit=2[pod_send][mix];
[ma_normalized][pod_send]sidechaincompress=level_sc=16:threshold=0.05:ratio=3.2:attack=10:release=750[compr];
[compr][mix]amix,aresample=48000[commentary_track];
"@
    }
    else {
        $filters += "[podstream]anull[commentary_track];"
    }

    if(-not $avoidVideoTranscode) {
        # Intro/outro cards
        #   fade in and out
        $concatStreams = @()
        if($introSpan.TotalSeconds -gt 0) {
            $concatStreams += '[intro]'
            $filters += "[$([int]$introInputNum):0]scale=$($videoInfo.width):$($videoInfo.height):force_original_aspect_ratio=decrease,pad=$($videoInfo.width):$($videoInfo.height):-1:-1:color=black,setsar=sar=$($videoInfo.sample_aspect_ratio -replace '\:','/'),tpad=stop_duration=$($introSpan.TotalMilliseconds - $frameMs)ms:stop_mode=clone[cover1];`n"
            $filters += "[cover1]fade=type=in:duration=2:start_time=1,fade=type=out:duration=2:start_time=$(if($introSpan.TotalSeconds -gt 2){$introSpan.TotalSeconds-2}else{'0'})[intro];`n"
        }
        $concatStreams += "[mv]"
        if($outroSpan.TotalSeconds -gt 0) {
            $concatStreams += '[outro]'
            $filters += "[$([int]$outroInputNum):0]scale=$($videoInfo.width):$($videoInfo.height):force_original_aspect_ratio=decrease,pad=$($videoInfo.width):$($videoInfo.height):-1:-1:color=black,setsar=sar=$($videoInfo.sample_aspect_ratio -replace '\:','/'),tpad=stop_duration=$($outroSpan.TotalMilliseconds - $frameMs)ms:stop_mode=clone[cover2];`n"
            $filters += "[cover2]fade=type=in:duration=2:start_time=1,fade=type=out:duration=2:start_time=$(if($outroSpan.TotalSeconds -gt 3){$outroSpan.TotalSeconds-3}else{'0'})[outro];`n"
        }

        # Join all clips together
        $filters += "$($concatStreams -join '')concat=n=$($concatStreams.Count):v=1[movie_padded]"
    }
    # Remove trailing semicolon if present
    $filters = $filters -replace ';$',''
    $filtersFilePath = Join-Path $workingDir "filters.txt"
    $filters | Set-Content $filtersFilePath
    return $filtersFilePath
}



###########
# Main
#

if($advancedSettingsFile) {
    if (-not (Test-Path $advancedSettingsFile)) {
        throw "Unable to open advanced settings file"
    }
    else {
        $advancedSettings = get-content -raw $advancedSettingsFile | ConvertFrom-Json
        $movieCutPoints = $advancedSettings.movieCutPoints
        $movieDurationChangeMs = 0
        $movieCutPoints | ForEach-Object {
            if($_.type -eq 'pause') {
                $movieDurationChangeMs += $_.durationMs
            }
            elseif($_.type -eq 'skip') {
                $movieDurationChangeMs -= $_.durationMs
            }
        }
        $podCutPoints = $advancedSettings.podCutPoints
        $podDurationChangeMs = 0
        $podCutPoints | ForEach-Object {
            if($_.type -eq 'pause') {
                $podDurationChangeMs += $_.durationMs
            }
            elseif($_.type -eq 'skip') {
                $podDurationChangeMs -= $_.durationMs
            }
        }
        $movieStartTime = $advancedSettings.movieStartTime
    }
}

# Check for FFMPEG
$ffmpegCmd = Get-Command $ffmpegPath -ErrorAction SilentlyContinue
if(-not $ffmpegCmd) {
    # Look in script dir
    $ffmpegCmd = Get-Command (join-path $PSScriptRoot "ffmpeg") -ErrorAction SilentlyContinue
    if(-not $ffmpegCmd) {
        throw "Unable to find ffmpeg! Please specify the path"
    }
}
$ffmpegPath = $ffmpegCmd.path
# Check for FFprobe
$ffprobeCmd = Get-Command $ffprobePath -ErrorAction SilentlyContinue
if(-not $ffprobeCmd) {
    # Look in script dir
    $ffprobeCmd = Get-Command (join-path $PSScriptRoot "ffprobe") -ErrorAction SilentlyContinue
    if(-not $ffprobeCmd) {
        throw "Unable to find ffprobe! Please specify the path"
    }
}
$ffprobePath = $ffprobeCmd.path

# Validate inputs
if(-not $inputMovie) {
    throw "No video file specified!"
} elseif (test-path $inputMovie) {
    $inputMovie = (get-item $inputMovie).FullName
    $inputMovieAudio = $inputMovie
} else {
    throw "Input movie file not found!"
}
if(-not $inputPodcast) {
    throw "No podcast file specified!"
} elseif (test-path $inputPodcast) {
    $inputPodcast = (get-item $inputPodcast).FullName
} else {
    throw "Input podcast file not found!"
}
if($inputSubs) {
    if(Test-Path $inputSubs) {
        $inputSubs = (get-item $inputSubs).FullName
    } else {
        throw "Input subtitle file not found!"
    }
}
$introSpan = [timespan]::Zero
if(-not [timespan]::TryParse($movieStartTime, [ref]$introSpan)) {
    if($movieStartTime -match '^\d+$') {
        $introSpan = [timespan]::new(0,0,0,0,$movieStartTime)
    }
    else {
        throw "Unable to parse timecode"
    }
}
# Create directory for temporary files
$workingDir = New-Item -Type directory -Path $outputDirectory -Name "$((get-item $inputMovie).BaseName)_temp" -Force

# Gather input movie details
$movieProbe = (&$ffprobePath -v quiet -print_format json -show_format -show_streams "$inputMovie") | convertfrom-json
$videoInfo = ($movieProbe.streams | Where-Object codec_type -eq 'video' | Select-Object -First 1)
$audioInfo = ($movieProbe.streams | Where-Object codec_type -eq 'audio')[$audioStreamIndex]
$framerate = $videoInfo.r_frame_rate
$frameMs = 1/(Invoke-Expression $framerate)
$timebase = $videoInfo.time_base
$sar = $videoInfo.sample_aspect_ratio
$pixfmt = $videoInfo.pix_fmt
$videoprofile = $videoInfo.profile.ToLower() -replace ' ',''
$videolevel = $videoInfo.level -replace '(\d)(\d)','$1.$2'
$codecname = $videoInfo.codec_name
$embeddedSubsPresent = ($movieProbe.streams.codec_type -like 'subtitle').Count -gt 0

# Select the default cover image based on aspect ratio (if present)
$coverPath = Join-Path $PSScriptRoot "cover.jpg"
$coverWidePath = Join-Path $PSScriptRoot "cover_wide.jpg"
if(-not $coverImg -and (test-path $coverPath) -and (test-path $coverWidePath)) {
    # Pick 16:9 or scope cover image
    $coverImg = @($coverPath,$coverWidePath)[($videoInfo.width/$videoInfo.height) -gt 2]
}

# Remux podcast to get better duration info
$podcastRemuxName = "$((get-item $inputPodcast).BaseName)_remux$((get-item $inputPodcast).Extension)"
if(runFfmpeg -ffmpegArguments @("-y", "-i `"$inputPodcast`"", "-c copy", "`"$podcastRemuxName`"") `
    -activity "Remuxing Podcast" -totalBytes (Get-Item $inputPodcast).Length) {
        throw "Unable to remux podcast!"
}
$inputPodcast = Join-Path $workingDir $podcastRemuxName

$podProbe = (&$ffprobePath -v quiet -print_format json -show_format -show_streams "$inputPodcast") | convertfrom-json

# Set stream timespans
$movieSpan = [timespan]::new(0,0,0,0,[double]($movieProbe.format.duration)*1000)
if($movieDurationChangeMs) { $movieSpan = $movieSpan + [timespan]::new(0,0,0,0,$movieDurationChangeMs) }
$podSpan = [timespan]::new(0,0,0,0,[double]($podProbe.format.duration)*1000)
if($podDurationChangeMs) { $podSpan = $podSpan + [timespan]::new(0,0,0,0,$podDurationChangeMs) }
$totalSpan = (($introSpan + $movieSpan), $podSpan) | Sort-Object -Descending | Select-Object -First 1
$outroSpan = $totalSpan - ($introSpan + $movieSpan)
$frameCount = $totalSpan.TotalSeconds * (Invoke-Expression $framerate)

# Avoid transcoding video by default, unless an output codec is specified
#   Set output codecs to match input video specs for intro/outro and cut points that may be rendered
$avoidVideoTranscode = -not ([bool]$outputVideoCodec)
if($avoidVideoTranscode) {
    if ($codecname -eq 'h264') {
        $outputVideoCodec = "-c:v libx264 -profile:v $videoprofile -level:v $videolevel -pix_fmt $pixfmt"
    }
    elseif ($codecname -eq 'hevc') {
        $outputVideoCodec = "-c:v libx265 -profile:v $videoprofile -level:v $videolevel -pix_fmt $pixfmt"
    }
    else {
        throw "Input video is not h264/h265, unable to copy the video stream. Include an outputVideoCodec parameter to reencode instead"
    }
}

# Set default movie audio codec to aac if mono/stereo, ac3 if surround
if(-not $outputMovieAudioCodec) {
    $outputMovieAudioCodec = "-c:a:1 $(@("aac","ac3")[([int]($audioInfo.channels) -gt 2)])"
}

# Transcode the original audio to pcm
#   May help with problems transcoding input files
if($reencodeSourceAudio) {
    $reencodedMovieAudio = join-path $workingDir "$((Get-Item $inputMovie).BaseName)_audio_reencode.mkv"
    $ffmpegCommand = @(
        "-y", "-i `"$inputMovie`"", "-map 0:a:$audioStreamIndex", "-c:a pcm_s16le"
        "`"$reencodedMovieAudio`""
    )
    if(runFfmpeg -ffmpegArguments $ffmpegCommand `
        -activity "Decoding Source Audio" -totalTime $movieSpan) {
            throw "Unable to decode audio!"
    }
    $inputMovieAudio = $reencodedMovieAudio
    $audioStreamIndex = 0
}

# Re-edit subtitles if movie has cut points
if($movieCutPoints) {
    if ($inputSubs) {
        $inputSubs = generateEditedSubs $movieCutPoints $inputSubs
    }
    elseif ($embeddedSubsPresent) {
        $inputSubs = generateEditedSubs $movieCutPoints $inputMovie
    }
}

$offsetMovieInputNum, $podInputNum, $movieAudioInputNum, `
    $movieVideoInputNum, $introInputNum, $outroInputNum, $subsInputNum = @(-1)*7
$inputs = generateInputList ([ref]$offsetMovieInputNum) ([ref]$podInputNum) ([ref]$movieAudioInputNum) `
    ([ref]$movieVideoInputNum) ([ref]$introInputNum) ([ref]$outroInputNum) ([ref]$subsInputNum)
$filters = createFilterFile
$ffmpegCommand = @(
    "-y",
    ($inputs -join ' '),
    "$(if($avoidVideoTranscode){"-c:v copy"}else{$outputVideoCodec})",
    $outputCommentaryCodec,
    $outputMovieAudioCodec,
    "-c:s copy",
    "-filter_complex_script `"$filters`"",
    "$(if($avoidVideoTranscode){"-map $([int]$movieVideoInputNum):v:0"}else{"-map [movie_padded]"})",
    "-map [commentary_track]",
    "$(if(-not $omitMovieAudioTrack){"-map [movie_audio]"})",
    "$(if($embeddedSubsPresent -or $inputSubs){ "-map $([int]$subsInputNum):s" })",
    "-metadata:s:a:0 title=`"Commentary`" -metadata:s:a:1 title=`"Movie Audio`"",
    "-avoid_negative_ts make_zero", "-max_interleave_delta 0"
    "`"$(Join-Path $outputDirectory ((get-item $inputMovie).BaseName + '_commentary.mkv'))`""
)
$ffmpegCommand | Set-Content (Join-Path $workingDir "ffmpegCommand.txt")
if(runFfmpeg -ffmpegArguments $ffmpegCommand -activity "Encoding video" -totalFrames $frameCount) {
    Write-Host "Unable to encode video!"
}

# Clean up temporary files
for($r=0; $r -lt 4; $r++) {
    try {
        if(-not (Test-Path $workingDir)) { break; }
        Remove-Item -Force -Recurse -Path $workingDir -ErrorAction Stop
    }
    catch {
        if($r -lt 3) {
            Write-Host "Temp folder locked, retrying in 5 seconds"
            Start-Sleep -Seconds 5
        }
        else {
            Write-Host "Unable to remove temp folder" 
        }
    }
}