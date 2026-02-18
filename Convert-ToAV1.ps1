# Video Conversion Script using Hardware GPU Encoder
# Auto-detects best available encoder (AV1 preferred, HEVC fallback for Qualcomm)
# Priority: NVIDIA AV1 > AMD AV1 > Intel AV1 > Qualcomm HEVC > Software AV1
# Converts videos > 250MB, keeps original only if new file is larger
# Skips already efficient codecs (HEVC, AV1, VP9)

$ErrorActionPreference = "Continue"

# Track current state for cleanup on interruption
$script:CurrentTempFile = $null
$script:ConversionHistory = $null
$script:InterruptedCleanly = $false

# Cleanup function for interruption (Ctrl+C, shutdown, etc.)
function Invoke-CleanupOnExit {
    if ($script:InterruptedCleanly) { return }
    $script:InterruptedCleanly = $true
    
    Write-Host "`n[INTERRUPTED] Cleaning up..."
    
    # Remove temp file if exists
    if ($script:CurrentTempFile -and (Test-Path $script:CurrentTempFile)) {
        Write-Host "  Removing incomplete temp file: $script:CurrentTempFile"
        Remove-Item $script:CurrentTempFile -Force -ErrorAction SilentlyContinue
    }
    
    # Save history before exit
    if ($script:ConversionHistory -and $script:HistoryFile) {
        Write-Host "  Saving conversion history..."
        try {
            $script:ConversionHistory | ConvertTo-Json -Depth 3 | Set-Content $script:HistoryFile -Encoding UTF8
            Write-Host "  History saved to: $script:HistoryFile"
        } catch {
            Write-Host "  Warning: Could not save history: $_"
        }
    }
    
    Write-Host "[INTERRUPTED] Cleanup complete. Run script again to continue."
}

# Register Ctrl+C handler
$null = Register-EngineEvent -SourceIdentifier PowerShell.Exiting -Action { Invoke-CleanupOnExit } -ErrorAction SilentlyContinue

# Also use trap for Ctrl+C during script execution
trap {
    Invoke-CleanupOnExit
    break
}

# Auto-detect OneDrive folder (Personal > Commercial > Generic)
function Get-OneDrivePath {
    # Try personal OneDrive first
    if ($env:OneDriveConsumer -and (Test-Path $env:OneDriveConsumer)) {
        return $env:OneDriveConsumer
    }
    # Try commercial/work OneDrive
    if ($env:OneDriveCommercial -and (Test-Path $env:OneDriveCommercial)) {
        return $env:OneDriveCommercial
    }
    # Try generic OneDrive variable
    if ($env:OneDrive -and (Test-Path $env:OneDrive)) {
        return $env:OneDrive
    }
    # Fallback: check user profile for OneDrive folders
    $userProfile = $env:USERPROFILE
    $oneDriveFolders = Get-ChildItem -Path $userProfile -Directory -Filter "OneDrive*" -ErrorAction SilentlyContinue
    if ($oneDriveFolders) {
        return $oneDriveFolders[0].FullName
    }
    return $null
}

$SourcePath = Get-OneDrivePath
if (-not $SourcePath) {
    Write-Host "ERROR: Could not detect OneDrive folder. Please set SourcePath manually."
    exit 1
}

$LogFile = Join-Path $SourcePath "AV1_Conversion_Log.txt"
$script:HistoryFile = Join-Path $SourcePath "AV1_Conversion_History.json"  # Tracks processed files and cumulative savings
$MinSizeMB = 250
$MinBitrateKbps = 1500  # Skip files already compressed below this bitrate (kbps)
$MinSavingsPercent = 10  # Skip if estimated/tested savings is below this percentage
$PrefetchCount = 3  # Number of files to prefetch ahead
$TestEncodeDuration = 15  # Seconds to test encode before committing to full conversion
$VideoExtensions = @("*.mp4", "*.mov", "*.avi", "*.wmv", "*.flv", "*.webm", "*.m4v", "*.mpg", "*.mpeg")

# Codecs that are already efficient (but may still convert if high bitrate)
$EfficientCodecs = @("hevc", "h265", "av1", "vp9")

# Estimated compression ratios by codec (used to predict savings before downloading)
# These are conservative estimates - actual savings may be higher
$EstimatedCompressionRatio = @{
    "h264" = 0.40  # H.264 to AV1 typically saves 50-70%, estimate 60%
    "mpeg4" = 0.35  # MPEG-4 to AV1 typically saves 60-80%
    "msmpeg4v3" = 0.35
    "wmv3" = 0.40
    "vc1" = 0.45
    "vp8" = 0.50
    "hevc" = 0.75  # HEVC to AV1 saves less, maybe 20-30%
    "h265" = 0.75
    "default" = 0.50  # Conservative default
}

# Load or initialize conversion history (tracks processed files and cumulative savings)
function Get-ConversionHistory {
    if (Test-Path $script:HistoryFile) {
        try {
            $content = Get-Content $script:HistoryFile -Raw
            return $content | ConvertFrom-Json -AsHashtable
        } catch {
            return @{ ProcessedFiles = @{}; TotalSavedBytes = 0 }
        }
    }
    return @{ ProcessedFiles = @{}; TotalSavedBytes = 0 }
}

function Save-ConversionHistory {
    param([hashtable]$History)
    $History | ConvertTo-Json -Depth 3 | Set-Content $script:HistoryFile -Encoding UTF8
}

# Mark a file as processed in history (by original path hash to handle renames)
# Also saves history immediately to prevent data loss on interruption
function Add-ProcessedFile {
    param([hashtable]$History, [string]$FilePath, [string]$Status, [long]$SavedBytes)
    # Use relative path from OneDrive root as key (handles drive letter changes)
    $relativePath = $FilePath.Replace($SourcePath, "").TrimStart("\")
    $History.ProcessedFiles[$relativePath] = @{
        Status = $Status
        ProcessedDate = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
        SavedBytes = $SavedBytes
    }
    if ($SavedBytes -gt 0) {
        $History.TotalSavedBytes += $SavedBytes
    }
    # Save immediately to prevent data loss on interruption
    Save-ConversionHistory -History $History
}

# Check if a file was already processed
function Test-FileProcessed {
    param([hashtable]$History, [string]$FilePath)
    $relativePath = $FilePath.Replace($SourcePath, "").TrimStart("\")
    return $History.ProcessedFiles.ContainsKey($relativePath)
}

# Estimate potential savings based on codec and file size
function Get-EstimatedSavings {
    param([string]$Codec, [long]$FileSize)
    $ratio = $EstimatedCompressionRatio[$Codec]
    if (-not $ratio) { $ratio = $EstimatedCompressionRatio["default"] }
    $estimatedNewSize = [long]($FileSize * $ratio)
    $estimatedSavings = $FileSize - $estimatedNewSize
    $estimatedPercent = [math]::Round(($estimatedSavings / $FileSize) * 100, 0)
    return @{
        EstimatedNewSize = $estimatedNewSize
        EstimatedSavings = $estimatedSavings
        EstimatedPercent = $estimatedPercent
        CompressionRatio = $ratio
    }
}

# Test encode a short segment to measure actual compression ratio
function Test-EncodeSavings {
    param(
        [string]$InputPath,
        [array]$EncoderArgs,
        [int]$Duration = 15
    )
    
    $testOutput = Join-Path $env:TEMP "av1_test_segment.mkv"
    
    # Get original file's bitrate for the segment we'll test
    $origBitrate = & ffprobe -v error -select_streams v:0 -show_entries stream=bit_rate -of default=noprint_wrappers=1:nokey=1 "$InputPath" 2>$null
    
    # Encode a short segment (middle of video for more representative sample)
    # Use -ss before -i for fast seeking
    $ffmpegArgs = @("-y", "-ss", "10", "-t", "$Duration", "-i", "`"$InputPath`"") + $EncoderArgs + @("-an", "`"$testOutput`"")
    
    $process = Start-Process -FilePath "ffmpeg" -ArgumentList $ffmpegArgs -NoNewWindow -Wait -PassThru -RedirectStandardError (Join-Path $env:TEMP "ffmpeg_test_err.txt")
    
    if ($process.ExitCode -ne 0 -or -not (Test-Path $testOutput)) {
        Remove-Item $testOutput -Force -ErrorAction SilentlyContinue
        return $null
    }
    
    # Get the test output bitrate
    $newBitrate = & ffprobe -v error -select_streams v:0 -show_entries stream=bit_rate -of default=noprint_wrappers=1:nokey=1 "$testOutput" 2>$null
    
    # If bitrate detection fails, calculate from file size
    if (-not $newBitrate -or $newBitrate -eq "N/A") {
        $testSize = (Get-Item $testOutput).Length
        $newBitrate = [math]::Round(($testSize * 8) / $Duration, 0)
    }
    
    Remove-Item $testOutput -Force -ErrorAction SilentlyContinue
    
    if (-not $origBitrate -or $origBitrate -eq "N/A" -or -not $newBitrate) {
        return $null
    }
    
    $origBitrateNum = [double]$origBitrate
    $newBitrateNum = [double]$newBitrate
    
    if ($origBitrateNum -le 0) { return $null }
    
    $ratio = $newBitrateNum / $origBitrateNum
    $savingsPercent = [math]::Round((1 - $ratio) * 100, 1)
    
    return @{
        OriginalBitrate = [math]::Round($origBitrateNum / 1000, 0)  # kbps
        NewBitrate = [math]::Round($newBitrateNum / 1000, 0)  # kbps
        CompressionRatio = [math]::Round($ratio, 3)
        SavingsPercent = $savingsPercent
    }
}

function Write-Log {
    param([string]$Message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] $Message"
    Write-Host $logEntry
    Add-Content -Path $LogFile -Value $logEntry
}

# Auto-detect best AV1 encoder by actually testing if it works
function Get-BestAV1Encoder {
    $testFile = Join-Path $env:TEMP "av1_test.mp4"
    
    # Create a tiny test video
    $null = & ffmpeg -y -f lavfi -i "color=c=black:s=128x128:d=0.1" -c:v libx264 -t 0.1 "$testFile" 2>&1
    
    # Encoder configs in priority order with optimized parameters
    # Each config is tuned for best compression/quality balance
    $encoderConfigs = @(
        @{ 
            Name = "av1_nvenc"
            Type = "NVIDIA NVENC"
            # tune=hq, CQ mode for quality, lookahead for scene analysis, spatial-aq for adaptive quantization
            Args = @("-c:v", "av1_nvenc", "-tune", "hq", "-cq", "32", "-rc-lookahead", "32", "-spatial-aq", "1", "-aq-strength", "8")
            Codec = "AV1"
        },
        @{ 
            Name = "av1_amf"
            Type = "AMD AMF"
            # balanced preset for speed/quality tradeoff, CQP rate control (higher QP = smaller files)
            # QP 38/40 provides better compression while maintaining acceptable quality
            Args = @("-c:v", "av1_amf", "-quality", "balanced", "-rc", "cqp", "-qp_i", "38", "-qp_p", "40")
            Codec = "AV1"
        },
        @{ 
            Name = "av1_qsv"
            Type = "Intel Quick Sync"
            # medium preset (4), extbrc with lookahead for better compression
            Args = @("-c:v", "av1_qsv", "-preset", "4", "-extbrc", "1", "-look_ahead_depth", "40", "-adaptive_i", "1", "-adaptive_b", "1")
            Codec = "AV1"
        },
        @{ 
            Name = "av1_vulkan"
            Type = "Vulkan"
            # CQP mode with good quality QP value
            Args = @("-c:v", "av1_vulkan", "-rc_mode", "cqp", "-qp", "30")
            Codec = "AV1"
        },
        @{ 
            Name = "av1_mf"
            Type = "MediaFoundation AV1"
            # Quality rate control mode
            Args = @("-c:v", "av1_mf", "-rate_control", "quality", "-quality", "65")
            Codec = "AV1"
        },
        @{ 
            Name = "hevc_mf"
            Type = "MediaFoundation HEVC (Qualcomm)"
            # HEVC hardware encoding for Snapdragon X Elite (no AV1 HW encode yet)
            Args = @("-c:v", "hevc_mf", "-rate_control", "quality", "-quality", "70")
            Codec = "HEVC"
        },
        @{ 
            Name = "libsvtav1"
            Type = "SVT-AV1 (Software)"
            # preset 5 = good balance, crf 30 for better compression, film-grain synthesis for quality
            Args = @("-c:v", "libsvtav1", "-preset", "5", "-crf", "30", "-svtav1-params", "tune=0:film-grain=0")
            Codec = "AV1"
        }
    )
    
    foreach ($enc in $encoderConfigs) {
        $testOutput = Join-Path $env:TEMP "av1_test_out.mkv"
        $args = @("-y", "-i", "`"$testFile`"") + $enc.Args + @("-t", "0.1", "`"$testOutput`"")
        $process = Start-Process -FilePath "ffmpeg" -ArgumentList $args -NoNewWindow -Wait -PassThru -RedirectStandardError (Join-Path $env:TEMP "ffmpeg_err.txt")
        
        if ($process.ExitCode -eq 0 -and (Test-Path $testOutput)) {
            Remove-Item $testOutput -Force -ErrorAction SilentlyContinue
            Remove-Item $testFile -Force -ErrorAction SilentlyContinue
            return $enc
        }
    }
    
    Remove-Item $testFile -Force -ErrorAction SilentlyContinue
    return $null
}

# Get the best encoder
$Encoder = Get-BestAV1Encoder
if (-not $Encoder) {
    Write-Log "ERROR: No AV1 encoder available. Please install ffmpeg with AV1 support."
    exit 1
}

function Get-VideoCodec {
    param([string]$FilePath)
    $result = & ffprobe -v error -select_streams v:0 -show_entries stream=codec_name -of default=noprint_wrappers=1:nokey=1 "$FilePath" 2>$null
    if ([string]::IsNullOrWhiteSpace($result)) {
        return $null
    }
    return $result.Trim().ToLower()
}

function Get-VideoBitrate {
    param([string]$FilePath)
    # Get video stream bitrate in kbps
    $result = & ffprobe -v error -select_streams v:0 -show_entries stream=bit_rate -of default=noprint_wrappers=1:nokey=1 "$FilePath" 2>$null
    if ([string]::IsNullOrWhiteSpace($result) -or $result -eq "N/A") {
        # Fallback: estimate from file size and duration
        $duration = & ffprobe -v error -show_entries format=duration -of default=noprint_wrappers=1:nokey=1 "$FilePath" 2>$null
        if ($duration -and $duration -ne "N/A") {
            $fileSize = (Get-Item $FilePath).Length
            $durationSec = [double]$duration
            if ($durationSec -gt 0) {
                return [math]::Round(($fileSize * 8 / 1000) / $durationSec, 0)
            }
        }
        return $null
    }
    return [math]::Round([double]$result / 1000, 0)
}

function Get-FreeDiskSpace {
    param([string]$Path)
    $drive = [System.IO.Path]::GetPathRoot($Path)
    $driveInfo = [System.IO.DriveInfo]::new($drive)
    return $driveInfo.AvailableFreeSpace
}

function Wait-ForFileDownload {
    param([string]$FilePath, [int]$MaxWaitSeconds = 300)
    # Trigger download by reading the file
    Write-Log "  Triggering OneDrive download..."
    $null = Get-Content -Path $FilePath -TotalCount 1 -ErrorAction SilentlyContinue
    
    # Wait for file to become locally available
    $waited = 0
    $checkInterval = 5
    while ($waited -lt $MaxWaitSeconds) {
        Start-Sleep -Seconds $checkInterval
        $waited += $checkInterval
        
        $file = Get-Item $FilePath -ErrorAction SilentlyContinue
        if ($file) {
            $offlineAttr = 0x1000
            $recallAttr = 0x400000
            $attrs = [int]$file.Attributes
            if ((($attrs -band $offlineAttr) -eq 0) -and (($attrs -band $recallAttr) -eq 0)) {
                Write-Log "  Download complete (waited ${waited}s)"
                return $true
            }
        }
        if ($waited % 30 -eq 0) {
            Write-Log "  Still downloading... (${waited}s)"
        }
    }
    return $false
}

# Start downloading a file in background (non-blocking)
function Start-FileDownload {
    param([string]$FilePath)
    # Trigger download by reading the file - this starts the OneDrive download
    $null = Get-Content -Path $FilePath -TotalCount 1 -ErrorAction SilentlyContinue
}

# Free up local storage by setting file to cloud-only (OneDrive Files On-Demand)
function Set-FileCloudOnly {
    param([string]$FilePath)
    try {
        # Use attrib.exe to set the file as cloud-only (unpinned)
        # attrib +U -P sets the file to "free up space" state
        $result = & attrib.exe +U -P "$FilePath" 2>&1
        return $true
    } catch {
        return $false
    }
}

# Check if a file has finished downloading (non-blocking)
function Test-FileDownloadComplete {
    param([string]$FilePath)
    $file = Get-Item $FilePath -ErrorAction SilentlyContinue
    if ($file) {
        $offlineAttr = 0x1000
        $recallAttr = 0x400000
        $attrs = [int]$file.Attributes
        return (($attrs -band $offlineAttr) -eq 0) -and (($attrs -band $recallAttr) -eq 0)
    }
    return $false
}

Write-Log "=========================================="
Write-Log "Video Conversion Script Started"
Write-Log "=========================================="
Write-Log "OneDrive: $SourcePath"
Write-Log "Encoder: $($Encoder.Type) ($($Encoder.Name))"
Write-Log "Output Codec: $($Encoder.Codec)"
Write-Log "Efficient codecs (skip if low bitrate): $($EfficientCodecs -join ', ')"
Write-Log "Min bitrate threshold: ${MinBitrateKbps} kbps (skip if already below)"

# Load conversion history
$script:ConversionHistory = Get-ConversionHistory
$previousSavedGB = [math]::Round($script:ConversionHistory.TotalSavedBytes / 1GB, 2)
$previousProcessedCount = $script:ConversionHistory.ProcessedFiles.Count
Write-Log "History: $previousProcessedCount files previously processed, ${previousSavedGB} GB saved total"

# Function to check if file is locally available (not cloud-only in OneDrive)
function Test-FileLocallyAvailable {
    param([System.IO.FileInfo]$File)
    # OneDrive cloud-only files have FILE_ATTRIBUTE_RECALL_ON_DATA_ACCESS (0x400000) or 
    # FILE_ATTRIBUTE_OFFLINE (0x1000) attributes
    $offlineAttr = 0x1000
    $recallAttr = 0x400000
    $attrs = [int]$File.Attributes
    return (($attrs -band $offlineAttr) -eq 0) -and (($attrs -band $recallAttr) -eq 0)
}

# Find all video files > 250MB (exclude .mkv files that might already be AV1)
$allVideos = Get-ChildItem -Path $SourcePath -Recurse -File -Include $VideoExtensions -ErrorAction SilentlyContinue | 
    Where-Object { $_.Length -gt ($MinSizeMB * 1MB) }

# Filter out already processed files from history (before checking local/cloud status)
$historyCount = $script:ConversionHistory.ProcessedFiles.Count
$unprocessedVideos = $allVideos | Where-Object { 
    -not (Test-FileProcessed -History $script:ConversionHistory -FilePath $_.FullName) 
}
$alreadyProcessedCount = $allVideos.Count - $unprocessedVideos.Count

# Separate local and cloud files, process local first
$localVideos = $unprocessedVideos | Where-Object { Test-FileLocallyAvailable $_ }
$cloudVideos = $unprocessedVideos | Where-Object { -not (Test-FileLocallyAvailable $_) }

Write-Log "Found $($allVideos.Count) video files larger than ${MinSizeMB}MB"
Write-Log "  - Already processed (in history): $alreadyProcessedCount"
Write-Log "  - Remaining to process: $($unprocessedVideos.Count)"
Write-Log "    - Local (ready): $($localVideos.Count)"
Write-Log "    - Cloud (will download): $($cloudVideos.Count)"

# Sort by file size descending (larger files = more potential savings, process first)
$localVideos = $localVideos | Sort-Object -Property Length -Descending
$cloudVideos = $cloudVideos | Sort-Object -Property Length -Descending

# Process local files first, then cloud files
$videos = @($localVideos) + @($cloudVideos)

$totalOriginalSize = 0
$totalNewSize = 0
$convertedCount = 0
$skippedCount = 0
$failedCount = 0

# Track prefetch state for parallel downloading (multiple files)
$prefetchedFiles = @{}  # Hash of paths being prefetched

for ($i = 0; $i -lt $videos.Count; $i++) {
    $video = $videos[$i]
    $inputPath = $video.FullName
    $inputSize = $video.Length
    $inputSizeMB = [math]::Round($inputSize / 1MB, 0)
    $inputSizeGB = [math]::Round($inputSize / 1GB, 2)
    
    # Generate output path (same name, .mkv extension)
    $outputDir = $video.DirectoryName
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($video.Name)
    $outputPath = Join-Path $outputDir "$baseName.mkv"
    
    # Skip if output file already exists
    if (Test-Path $outputPath) {
        Write-Log "SKIP: Output already exists - $outputPath"
        $skippedCount++
        continue
    }
    
    # Skip if file was already scanned/processed in a previous run (don't re-download)
    if (Test-FileProcessed -History $script:ConversionHistory -FilePath $inputPath) {
        Write-Log "SKIP (already processed in previous run): $($video.Name)"
        $skippedCount++
        continue
    }
    
    Write-Log "----------------------------------------"
    $isLocal = Test-FileLocallyAvailable $video
    Write-Log "Processing: $($video.Name) $(if(-not $isLocal){'[CLOUD]'})"
    Write-Log "Size: ${inputSizeMB} MB | Path: $inputPath"
    
    # Check codec BEFORE downloading - ffprobe only needs file headers (a few KB)
    # OneDrive will stream just enough data to read the header, not the whole file
    $codec = Get-VideoCodec -FilePath $inputPath
    if ([string]::IsNullOrWhiteSpace($codec)) {
        Write-Log "SKIP: Could not read codec (file may be corrupted or inaccessible)"
        $skippedCount++
        continue
    }
    # Check bitrate before downloading - ffprobe can read this from headers too
    $bitrate = Get-VideoBitrate -FilePath $inputPath
    $isEfficientCodec = $EfficientCodecs -contains $codec
    
    # Skip logic:
    # - If efficient codec AND low bitrate: skip (already well-optimized)
    # - If efficient codec BUT high bitrate: convert (can still benefit from AV1)
    # - If inefficient codec AND low bitrate: skip (unlikely to benefit)
    # - If inefficient codec AND high bitrate: convert
    if ($bitrate -and $bitrate -lt $MinBitrateKbps) {
        Write-Log "SKIP (low bitrate: ${bitrate} kbps < ${MinBitrateKbps} kbps threshold): $($video.Name)"
        # Record in history so we don't re-scan this file
        Add-ProcessedFile -History $script:ConversionHistory -FilePath $inputPath -Status "skipped-low-bitrate" -SavedBytes 0
        # Free up local storage if file was locally available
        if ($isLocal) {
            if (Set-FileCloudOnly -FilePath $inputPath) {
                Write-Log "  Freed local storage: set to cloud-only"
            }
        }
        $skippedCount++
        continue
    }
    
    # Estimate savings before downloading - skip if not worth the time
    $estimate = Get-EstimatedSavings -Codec $codec -FileSize $inputSize
    if ($estimate.EstimatedPercent -lt $MinSavingsPercent) {
        Write-Log "SKIP (estimated savings ${estimate.EstimatedPercent}% < ${MinSavingsPercent}% threshold): $($video.Name)"
        Add-ProcessedFile -History $script:ConversionHistory -FilePath $inputPath -Status "skipped-low-savings" -SavedBytes 0
        if ($isLocal) {
            if (Set-FileCloudOnly -FilePath $inputPath) {
                Write-Log "  Freed local storage: set to cloud-only"
            }
        }
        $skippedCount++
        continue
    }
    
    $bitrateInfo = if ($bitrate) { " @ ${bitrate} kbps" } else { "" }
    $codecNote = if ($isEfficientCodec) { " (high bitrate, re-encoding)" } else { "" }
    $estimatedSavingsMB = [math]::Round($estimate.EstimatedSavings / 1MB, 0)
    Write-Log "Codec: $codec$bitrateInfo$codecNote - Est. savings: ~${estimatedSavingsMB} MB (~$($estimate.EstimatedPercent)%)"
    Write-Log "Converting to $($Encoder.Codec)..."
    
    # Now download the full file if it's cloud-only (we've confirmed it needs conversion)
    if (-not $isLocal) {
        # Check if this file was already prefetched
        if ($prefetchedFiles.ContainsKey($inputPath) -and (Test-FileDownloadComplete -FilePath $inputPath)) {
            Write-Log "  Using prefetched file"
        } else {
            Write-Log "  Downloading file for conversion..."
            $downloaded = Wait-ForFileDownload -FilePath $inputPath -MaxWaitSeconds 600
            if (-not $downloaded) {
                Write-Log "SKIP: Could not download file within timeout"
                # Don't record in history - may succeed on retry with better network
                $skippedCount++
                continue
            }
        }
    }
    $prefetchedFiles.Remove($inputPath)  # Clear from prefetch list
    
    # Start prefetching next eligible cloud files while we process this one (multi-file prefetch)
    $prefetchCount = 0
    for ($j = $i + 1; $j -lt $videos.Count -and $prefetchCount -lt $PrefetchCount; $j++) {
        $nextVideo = $videos[$j]
        # Skip if already being prefetched or is local
        if ($prefetchedFiles.ContainsKey($nextVideo.FullName) -or (Test-FileLocallyAvailable $nextVideo)) {
            continue
        }
        # Check if output already exists (would be skipped anyway)
        $nextBaseName = [System.IO.Path]::GetFileNameWithoutExtension($nextVideo.Name)
        $nextOutputPath = Join-Path $nextVideo.DirectoryName "$nextBaseName.mkv"
        if (Test-Path $nextOutputPath) { continue }
        # Skip if already in history
        if (Test-FileProcessed -History $script:ConversionHistory -FilePath $nextVideo.FullName) { continue }
        # Quick codec/bitrate check on next file before prefetching
        $nextCodec = Get-VideoCodec -FilePath $nextVideo.FullName
        $nextBitrate = Get-VideoBitrate -FilePath $nextVideo.FullName
        # Only prefetch if it needs conversion (high bitrate)
        if ($nextCodec -and $nextBitrate -and $nextBitrate -ge $MinBitrateKbps) {
            # Also check estimated savings
            $nextEstimate = Get-EstimatedSavings -Codec $nextCodec -FileSize $nextVideo.Length
            if ($nextEstimate.EstimatedPercent -ge $MinSavingsPercent) {
                Write-Log "  Prefetching: $($nextVideo.Name)"
                Start-FileDownload -FilePath $nextVideo.FullName
                $prefetchedFiles[$nextVideo.FullName] = $true
                $prefetchCount++
            }
        }
    }
    
    # Test encode a short segment to verify actual savings before full conversion
    Write-Log "  Testing ${TestEncodeDuration}s segment to verify savings..."
    $testResult = Test-EncodeSavings -InputPath $inputPath -EncoderArgs $Encoder.Args -Duration $TestEncodeDuration
    
    if ($testResult) {
        Write-Log "  Test result: $($testResult.OriginalBitrate) kbps -> $($testResult.NewBitrate) kbps ($($testResult.SavingsPercent)% savings)"
        
        if ($testResult.SavingsPercent -lt $MinSavingsPercent) {
            Write-Log "SKIP (test encode shows only $($testResult.SavingsPercent)% savings < ${MinSavingsPercent}% threshold): $($video.Name)"
            Add-ProcessedFile -History $script:ConversionHistory -FilePath $inputPath -Status "skipped-test-low-savings" -SavedBytes 0
            # Free up local storage
            if (Set-FileCloudOnly -FilePath $inputPath) {
                Write-Log "  Freed local storage: set to cloud-only"
            }
            $skippedCount++
            continue
        }
        
        # Update estimated savings with actual test result
        $estimatedSavingsMB = [math]::Round($inputSize * (1 - $testResult.CompressionRatio) / 1MB, 0)
        Write-Log "  Proceeding with full conversion (expected savings: ~${estimatedSavingsMB} MB)"
    } else {
        Write-Log "  Test encode failed, proceeding with full conversion anyway..."
    }
    
    # Convert video
    $tempOutput = Join-Path $outputDir "$baseName.conv.temp.mkv"
    $script:CurrentTempFile = $tempOutput  # Track for cleanup on interruption
    
    # Check disk space before conversion (need at least input file size as buffer)
    $freeSpace = Get-FreeDiskSpace -Path $outputDir
    $requiredSpace = $inputSize * 1.1  # 10% buffer
    if ($freeSpace -lt $requiredSpace) {
        $freeSpaceGB = [math]::Round($freeSpace / 1GB, 2)
        $requiredGB = [math]::Round($requiredSpace / 1GB, 2)
        Write-Log "SKIP (low disk space: ${freeSpaceGB} GB free, need ~${requiredGB} GB): $($video.Name)"
        $script:CurrentTempFile = $null
        $skippedCount++
        continue
    }
    
    $startTime = Get-Date
    
    # Run ffmpeg conversion with auto-detected encoder
    # -hwaccel auto: use hardware decoding if available (faster input processing)
    # -threads 0: use all CPU cores for any software processing
    $ffmpegArgs = @("-hwaccel", "auto", "-threads", "0", "-i", "`"$inputPath`"") + $Encoder.Args + @("-c:a", "copy", "-y", "`"$tempOutput`"")
    
    $process = Start-Process -FilePath "ffmpeg" -ArgumentList $ffmpegArgs -NoNewWindow -Wait -PassThru
    
    $endTime = Get-Date
    $duration = $endTime - $startTime
    
    if ($process.ExitCode -ne 0 -or -not (Test-Path $tempOutput)) {
        Write-Log "FAILED: FFmpeg conversion failed for $($video.Name)"
        if (Test-Path $tempOutput) { Remove-Item $tempOutput -Force }
        $script:CurrentTempFile = $null
        $failedCount++
        continue
    }
    
    $newSize = (Get-Item $tempOutput).Length
    $newSizeGB = [math]::Round($newSize / 1GB, 2)
    $savings = $inputSize - $newSize
    $savingsPercent = [math]::Round(($savings / $inputSize) * 100, 1)
    
    Write-Log "Conversion completed in $([math]::Round($duration.TotalMinutes, 1)) minutes"
    Write-Log "Original: ${inputSizeGB} GB | New: ${newSizeGB} GB | Savings: ${savingsPercent}%"
    
    if ($newSize -lt $inputSize) {
        # New file is smaller - rename temp to final and delete original
        Move-Item -Path $tempOutput -Destination $outputPath -Force
        Remove-Item -Path $inputPath -Force
        Write-Log "SUCCESS: Saved $([math]::Round($savings / 1MB, 0)) MB - Original deleted"
        $totalOriginalSize += $inputSize
        $totalNewSize += $newSize
        $convertedCount++
        
        # Record successful conversion in history
        Add-ProcessedFile -History $script:ConversionHistory -FilePath $inputPath -Status "converted" -SavedBytes $savings
        
        # Free up local storage by setting converted file to cloud-only
        if (Set-FileCloudOnly -FilePath $outputPath) {
            Write-Log "  Freed local storage: $($baseName).mkv set to cloud-only"
        }
    } else {
        # New file is larger or same - keep original, delete temp
        Remove-Item -Path $tempOutput -Force
        Write-Log "KEPT ORIGINAL: $($Encoder.Codec) file was larger/same size"
        # Record in history so we don't retry this file
        Add-ProcessedFile -History $script:ConversionHistory -FilePath $inputPath -Status "kept-original" -SavedBytes 0
        # Free up local storage for original file since we're keeping it
        if (Set-FileCloudOnly -FilePath $inputPath) {
            Write-Log "  Freed local storage: original set to cloud-only"
        }
        $skippedCount++
    }
    $script:CurrentTempFile = $null  # Clear temp file tracker after successful processing
}

# Save conversion history
Save-ConversionHistory -History $script:ConversionHistory

Write-Log "=========================================="
Write-Log "Conversion Summary"
Write-Log "=========================================="
Write-Log "Total files scanned this run: $($videos.Count)"
Write-Log "Successfully converted: $convertedCount"
Write-Log "Skipped (larger output or exists): $skippedCount"
Write-Log "Failed: $failedCount"
if ($convertedCount -gt 0) {
    $totalSavingsGB = [math]::Round(($totalOriginalSize - $totalNewSize) / 1GB, 2)
    Write-Log "Space saved this run: ${totalSavingsGB} GB"
}
Write-Log "----------------------------------------"
Write-Log "CUMULATIVE TOTALS (all runs):"
Write-Log "  Total files processed: $($script:ConversionHistory.ProcessedFiles.Count)"
$cumulativeSavedGB = [math]::Round($script:ConversionHistory.TotalSavedBytes / 1GB, 2)
Write-Log "  Total space saved on OneDrive: ${cumulativeSavedGB} GB"
Write-Log "----------------------------------------"
Write-Log "Log saved to: $LogFile"
Write-Log "History saved to: $script:HistoryFile"
Write-Log "=========================================="

# Mark clean exit
$script:InterruptedCleanly = $true
