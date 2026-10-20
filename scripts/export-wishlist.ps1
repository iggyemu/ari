<#
Export the original presentation without saving any changes to it.
Requires Windows, desktop PowerPoint, ffmpeg.exe and ffprobe.exe.
See WISHLIST.md for setup, validation and deployment instructions.
#>
[CmdletBinding()]
param(
    [string]$FfmpegPath = '',
    [string]$FfprobePath = '',
    [switch]$EndingPreview,
    [switch]$SkipPowerPoint
)
$ErrorActionPreference = 'Stop'
$slideSeconds = 5
$projectRoot = Split-Path -Parent $PSScriptRoot
$sourcePath = Join-Path $projectRoot 'dependencies\wishlist 2026 ari.pptx'
$buildPath = Join-Path $projectRoot '.wishlist-build'
$assetPath = Join-Path $projectRoot 'assets\wishlist-2026'
New-Item -ItemType Directory -Path $buildPath, $assetPath -Force | Out-Null
$sourceHash = (Get-FileHash -LiteralPath $sourcePath -Algorithm SHA256).Hash

function Find-VideoTool([string]$explicitPath, [string]$name) {
    if ($explicitPath) { return (Resolve-Path -LiteralPath $explicitPath).Path }
    $command = Get-Command $name -ErrorAction SilentlyContinue
    if ($command) { return $command.Source }
    $portable = Get-ChildItem -LiteralPath $buildPath -Filter $name -Recurse -File |
        Select-Object -First 1
    if ($portable) { return $portable.FullName }
    throw "Cannot find $name. Supply its path; see WISHLIST.md."
}
$ffmpeg = Find-VideoTool $FfmpegPath 'ffmpeg.exe'
$ffprobe = Find-VideoTool $FfprobePath 'ffprobe.exe'
$rawPath = Join-Path $buildPath $(if ($EndingPreview) {'ending-raw.mp4'} else {'wishlist-raw.mp4'})
$stagingPath = Join-Path $buildPath $(if ($EndingPreview) {'ending-preview.mp4'} else {'video-ready.mp4'})
$copyPath = Join-Path $buildPath 'wishlist-export-copy.pptx'

function Add-Animation($sequence, $shape, [int]$kind, [int]$trigger, [single]$duration) {
    # PowerPoint enums: 0 = no text build, 2 = with previous, 3 = after previous.
    $effect = $sequence.AddEffect($shape, $kind, 0, $trigger)
    $effect.Timing.Duration = $duration
    $effect.Timing.TriggerDelayTime = 0
    return $effect
}

if (-not $SkipPowerPoint) {
    Copy-Item -LiteralPath $sourcePath -Destination $copyPath -Force
    if (Test-Path -LiteralPath $rawPath) { Remove-Item -LiteralPath $rawPath }
    $wasPowerPointRunning = $null -ne (Get-Process POWERPNT -ErrorAction SilentlyContinue)
    $powerPoint = $null
    $deck = $null
    $previousSecurity = $null
    try {
        $powerPoint = New-Object -ComObject PowerPoint.Application
        $previousSecurity = $powerPoint.AutomationSecurity
        $powerPoint.AutomationSecurity = 3 # Disable macros in the working copy.
        $deck = $powerPoint.Presentations.Open($copyPath, 0, 0, 0)
        if ($deck.Slides.Count -ne 35) {
            throw 'The source slide count changed. Review the closing sequence and shopping list before exporting.'
        }
        for ($i = 1; $i -le $deck.Slides.Count; $i++) {
            $slide = $deck.Slides.Item($i)
            $slide.SlideShowTransition.AdvanceOnClick = 0
            $slide.SlideShowTransition.AdvanceOnTime = -1
            $slide.SlideShowTransition.AdvanceTime = [single]$slideSeconds
        }
        $deck.SlideShowSettings.AdvanceMode = 2 # ppSlideShowUseSlideTimings
        $instruction = $deck.Slides.Item(34).Shapes.Item('Title 1').TextFrame.TextRange.Find(
            'Click on 2 times the next slide')
        if ($null -ne $instruction) { $instruction.Delete() }

        $closing = $deck.Slides.Item(35)
        $golf = $closing.Shapes.Item('Video 14')
        $wave = $closing.Shapes.Item('Video 10')
        $thanks = $closing.Shapes.Item('TextBox 15')
        $golfSeconds = $golf.MediaFormat.Length / 1000.0
        $waveSeconds = $wave.MediaFormat.Length / 1000.0
        # Retain both media objects and their authored mute/volume settings.
        $golf.AnimationSettings.PlaySettings.LoopUntilStopped = 0
        $wave.AnimationSettings.PlaySettings.LoopUntilStopped = 0
        # Remove click-to-pause / click-to-reveal actions from this export copy.
        for ($i = $closing.TimeLine.InteractiveSequences.Count; $i -ge 1; $i--) {
            $sequence = $closing.TimeLine.InteractiveSequences.Item($i)
            while ($sequence.Count -gt 0) { $sequence.Item($sequence.Count).Delete() }
        }
        $sequence = $closing.TimeLine.MainSequence
        while ($sequence.Count -gt 0) { $sequence.Item($sequence.Count).Delete() }
        # The golf clip remains the background after it finishes. Reveal the
        # originally obscured wave clip in its authored position above it.
        $wave.ZOrder(0)
        $thanks.ZOrder(0)
        [void](Add-Animation $sequence $golf 83 2 $golfSeconds)
        [void](Add-Animation $sequence $wave 1 3 0.01)
        [void](Add-Animation $sequence $wave 83 2 $waveSeconds)
        [void](Add-Animation $sequence $thanks 10 3 0.5)
        $closing.SlideShowTransition.AdvanceTime = [single]($golfSeconds + $waveSeconds + 3)
        if ($EndingPreview) {
            for ($i = 34; $i -ge 1; $i--) { $deck.Slides.Item($i).Delete() }
        } else {
            $deck.Slides.Item(1).Export((Join-Path $buildPath 'poster.png'), 'PNG', 1920, 1080)
        }
        $deck.Save()
        Write-Output 'Exporting with PowerPoint at 1080p / 30 fps...'
        $deck.CreateVideo($rawPath, $true, $slideSeconds, 1080, 30, 85)
        $timer = [Diagnostics.Stopwatch]::StartNew()
        # ppMediaTaskStatus: none=0, inProgress=1, queued=2, done=3, failed=4.
        do {
            Start-Sleep -Seconds 5
            $status = [int]$deck.CreateVideoStatus
            if ($status -eq 4) { throw 'PowerPoint video export failed.' }
            if ($timer.Elapsed.TotalMinutes -gt 60) { throw 'PowerPoint export timed out.' }
            Write-Output ("PowerPoint export: {0:n0}s elapsed; status {1}" -f $timer.Elapsed.TotalSeconds, $status)
        } until ($status -eq 3)
    } finally {
        if ($null -ne $deck) { $deck.Close() }
        if ($null -ne $powerPoint) {
            if ($null -ne $previousSecurity) { $powerPoint.AutomationSecurity = $previousSecurity }
            if (-not $wasPowerPointRunning -and $powerPoint.Presentations.Count -eq 0) { $powerPoint.Quit() }
            [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($powerPoint)
        }
        if ((Get-FileHash -LiteralPath $sourcePath -Algorithm SHA256).Hash -ne $sourceHash) {
            throw 'Source PowerPoint hash changed unexpectedly.'
        }
    }
}
if (-not (Test-Path -LiteralPath $rawPath)) { throw "Missing native video export: $rawPath" }
$probeText = & $ffprobe -v error -show_format -show_streams -of json $rawPath
if ($LASTEXITCODE -ne 0) { throw 'Unable to inspect native video.' }
$rawInfo = ($probeText -join "`n") | ConvertFrom-Json
$duration = [double]::Parse($rawInfo.format.duration, [Globalization.CultureInfo]::InvariantCulture)
$encoding = @('-y', '-hide_banner', '-loglevel', 'warning', '-i', $rawPath,
    '-map', '0:v:0', '-map', '0:a?', '-c:v', 'libx264', '-preset', 'medium',
    '-crf', '23', '-pix_fmt', 'yuv420p', '-r', '30',
    '-c:a', 'aac', '-b:a', '128k', '-movflags', '+faststart', $stagingPath)
Write-Output 'Compressing the browser video...'
& $ffmpeg @encoding
if ($LASTEXITCODE -ne 0) { throw 'FFmpeg compression failed.' }
if ((Get-Item -LiteralPath $stagingPath).Length -ge 95MB) {
    # Two-pass fallback budgets 90 MiB, including audio and container overhead.
    $videoKbps = [int][Math]::Floor((90MB * 8 / $duration / 1000) - 160)
    if ($videoKbps -lt 200) { throw 'Video duration is too long for the configured file-size budget.' }
    $passLog = Join-Path $buildPath 'encoding-pass'
    $common = @('-y', '-hide_banner', '-loglevel', 'warning', '-i', $rawPath,
        '-map', '0:v:0', '-c:v', 'libx264', '-preset', 'medium',
        '-b:v', "${videoKbps}k", '-pix_fmt', 'yuv420p', '-r', '30', '-passlogfile', $passLog)
    & $ffmpeg @common -pass 1 -an -f null NUL
    if ($LASTEXITCODE -ne 0) { throw 'First encoding pass failed.' }
    & $ffmpeg @common -pass 2 -map '0:a?' -c:a aac -b:a 128k -movflags +faststart $stagingPath
    if ($LASTEXITCODE -ne 0) { throw 'Second encoding pass failed.' }
}
$probeText = & $ffprobe -v error -show_format -show_streams -of json $stagingPath
if ($LASTEXITCODE -ne 0) { throw 'Unable to inspect compressed video.' }
$info = ($probeText -join "`n") | ConvertFrom-Json
$video = $info.streams | Where-Object codec_type -eq 'video' | Select-Object -First 1
if ($video.codec_name -ne 'h264' -or $video.pix_fmt -ne 'yuv420p' -or
    $video.width -ne 1920 -or $video.height -ne 1080 -or $video.avg_frame_rate -ne '30/1' -or
    (Get-Item -LiteralPath $stagingPath).Length -ge 95MB) {
    throw 'Compressed video does not meet the publication requirements.'
}
$info | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath (
    Join-Path $buildPath $(if ($EndingPreview) {'ending-info.json'} else {'video-info.json'})) -Encoding utf8
if (-not $EndingPreview) {
    Copy-Item -LiteralPath $stagingPath -Destination (Join-Path $assetPath 'video.mp4') -Force
    & $ffmpeg -y -hide_banner -loglevel warning -i (Join-Path $buildPath 'poster.png') `
        -frames:v 1 -q:v 3 -update 1 (Join-Path $assetPath 'poster.jpg')
    if ($LASTEXITCODE -ne 0) { throw 'Poster creation failed.' }
}
Write-Output ("Ready: {0:n1} seconds, {1:n1} MiB" -f
    $duration, ((Get-Item -LiteralPath $stagingPath).Length / 1MB))
