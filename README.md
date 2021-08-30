# PodCommentary

## Overview
Powershell script to sync up movie commentary podcasts to the movie they're commentating.

- Mixes the podcast audio alongside the movie audio with auto-ducking
- Adds in intro and outro clips to the video to accompany the podcast bits before and after the movie.
- Handles pauses/skips in the movie that happen during the commentary record via advanced settings files.
- Subtitles from the source file will be used if found, otherwise an external subtitle file can be added.
- Outputs a video file with the new mix and the original movie audio as separate tracks
- Tested on PowerShell Core (cross-platform) and PowerShell 5.1 (Windows-only)

Movie start times for a handful of podcasts are listed in the [Podcasts section of this repo](Podcasts)

## Usage

- Download `PodCommentary.ps1` to a new directory
- ffmpeg and ffprobe are required, these must either be available on the system path or placed in the same directory as the script.
  - Windows builds for ffmpeg can be found at https://www.gyan.dev/ffmpeg/builds/.
- Optionally add files `cover.jpg` and `cover_wide.jpg` to the script directory for intro/outro clips.
  - `cover.jpg` should be 1920x1080
  - `cover_wide.jpg` should be 1920x800
  - Some covers can be found in the [Podcasts](Podcasts) directories in this repo
  - A cover image can be specified manually with the `coverImg` argument.
  - If no image is provided, black frames will be used for the intro/outro.
- Run the script
  - The `inputPodcast`, `inputMovie`, and `movieStartTime` arguments are required
  - Find the movie start timecode in the [Podcasts](Podcasts) docs (or enter your own if your podcast isn't listed)
  - From a Powershell prompt in the script directory
    - `.\PodCommentary.ps1 -inputMovie "C:\Movies\Iron Man.mkv" -inputPodcast "C:\Podcasts\Blank Check\Iron Man.mp3" -movieStartTime 00:08:02.482`
  - The script may need to be unblocked before running, see [this article](https://social.technet.microsoft.com/wiki/contents/articles/38496.unblock-downloaded-powershell-scripts.aspx) for details
- The video will start encoding and be output in the same directory as the script with the same name as the input video with a `_commentary` suffix attached.
  - The output directory can also be changed with the `outputDirectory` argument.

## Arguments

| Argument               | Description |
|------------------------|-------------|
| movieStartTime         | (Required) The time in the podcast that the movie begins.<br>Can be provided in either timecode (e.g. `00:08:21.000`) or milliseconds (e.g. `501000`) |
| inputPodcast           | (Required) Path to the podcast file |
| inputMovie             | (Required) Path to the movie file   |
| inputSubs              | Path to subtitle file (e.g. `.srt` file) |
| coverImg               | Path to an image file to use for the intro/outro clips  |
| advancedSettingsFile   | Path to a `.json` settings file that provides movie or podcast edit info to correct commentary sync issues (usually due to the movie being paused).<br>Advanced settings files can be found for select podcasts in the [Podcasts directory](Podcasts) in this repo. `movieStartTime` can be omitted if using a provided settings file. |
| outputDirectory        | Path to the directory the output file should be saved in. Defaults to the directory the script is saved in. |
| audioStreamIndex       | 0-indexed number of the audio stream to select (if the file has more than one audio track). Defaults to `0`, meaning the first audio track in the movie file. |
| outputVideoCodec       | ffmpeg video encoder settings to use for output video. Only required if the output video should be transcoded, otherwise the video will be copied directly from the input file (faster).<br>Example settings for a small HEVC transcode are `-c:v libx265 -crf 21 -preset faster -pix_fmt yuv420p10le` |
| outputCommentaryCodec  | ffmpeg audio encoder settings for the output commentary track.<br>This defaults to `-c:a:0 aac -ac:a:0 2` |
| outputMovieAudioCodec  | ffmpeg audio encoder settings for the output movie audio track.<br>This defaults to AAC if the input movie audio is stereo (`-c:a:1 aac`), and AC3 if there are >2 audio channels (`-c:a:1 ac3`). |
| reencodeSourceAudio    | (Switch) Transcodes the audio from the input movie to PCM before any audio processing is done |
| useLoudnorm            | (Switch) Uses the ffmpeg `loudnorm` filter to normalize the podcast/movie audio instead of the default `dynaudnorm`. This will use EBU R 128 normalization, which is more standard, but encoding is slower. |
| omitMovieAudioTrack    | (Switch) Skips encoding the movie audio track in the output file |
| disableBackgroundAudio | (Switch) Uses the input podcast audio directly as the output commentary track, no mixing in movie audio |
| ffmpegPath             | Path to the ffmpeg binary.<br>Only needed if ffmpeg isn't in the system path or the script directory |
| ffprobePath            | Path to the ffprobe binary.<br>Only needed if ffprobe isn't in the system path or the script directory |
