param(
    [string]$ArchivePath = ""
)

$ErrorActionPreference = "Stop"
$projectRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$buildRoot = Join-Path $projectRoot "_build_runtime"
$outputRoot = Join-Path $buildRoot "media-tools"
$outputBin = Join-Path $outputRoot "bin"
$expectedSha256 = "5cef7bc48e513f3cee7c881e8b2999d835541edb1846c3e89443a67f8f8556ab"
$url = "https://github.com/BtbN/FFmpeg-Builds/releases/download/autobuild-2026-10-31-12-59/ffmpeg-n8.1.3-27-g8dca1586f4-win64-lgpl-shared-8.1.zip"

if (-not $ArchivePath) {
    $ArchivePath = Join-Path $buildRoot "ffmpeg-lgpl-shared.zip"
    if (-not (Test-Path -LiteralPath $ArchivePath)) {
        New-Item -ItemType Directory -Force -Path $buildRoot | Out-Null
        Invoke-WebRequest -Uri $url -OutFile $ArchivePath
    }
}
$ArchivePath = (Resolve-Path -LiteralPath $ArchivePath).Path
$actual = (Get-FileHash -LiteralPath $ArchivePath -Algorithm SHA256).Hash.ToLowerInvariant()
if ($actual -ne $expectedSha256) {
    throw "FFmpeg archive checksum mismatch: $actual"
}

New-Item -ItemType Directory -Force -Path $outputBin | Out-Null
Add-Type -AssemblyName System.IO.Compression.FileSystem
$archive = [System.IO.Compression.ZipFile]::OpenRead($ArchivePath)
try {
    $copied = @()
    foreach ($entry in $archive.Entries) {
        $name = [System.IO.Path]::GetFileName($entry.FullName)
        if ($entry.FullName -match '/bin/' -and ($name -in @('ffmpeg.exe', 'ffprobe.exe') -or $name -like '*.dll')) {
            [System.IO.Compression.ZipFileExtensions]::ExtractToFile($entry, (Join-Path $outputBin $name), $true)
            $copied += $name
        }
        elseif ($name -eq 'LICENSE.txt' -and $entry.FullName -notmatch '/bin/') {
            [System.IO.Compression.ZipFileExtensions]::ExtractToFile($entry, (Join-Path $outputRoot 'LICENSE.txt'), $true)
        }
    }
}
finally {
    $archive.Dispose()
}

foreach ($name in @('ffmpeg.exe', 'ffprobe.exe', 'avcodec-62.dll', 'avformat-62.dll', 'avutil-60.dll')) {
    if (-not (Test-Path -LiteralPath (Join-Path $outputBin $name))) {
        throw "FFmpeg package is incomplete: $name"
    }
}
if (-not (Test-Path -LiteralPath (Join-Path $outputRoot 'LICENSE.txt'))) {
    throw "FFmpeg license text is missing."
}
& (Join-Path $outputBin 'ffmpeg.exe') -version | Select-Object -First 1
& (Join-Path $outputBin 'ffprobe.exe') -version | Select-Object -First 1
Write-Host "Prepared $($copied.Count) FFmpeg files in $outputRoot"
