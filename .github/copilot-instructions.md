# Copilot Instructions

PowerShell script for batch converting OneDrive video files to AV1/HEVC using hardware-accelerated encoding.

## Architecture

**Single-script tool** (`Convert-ToAV1.ps1`) with these key components:

1. **Encoder auto-detection** (`Get-BestAV1Encoder`) - Tests encoders in priority order by running ffmpeg against a temp file. Priority: NVIDIA > AMD > Intel > Qualcomm HEVC > Software SVT-AV1
2. **OneDrive integration** - Detects OneDrive path via environment variables, handles cloud-only files by triggering downloads and waiting for sync
3. **Safe file handling** - Only replaces original if converted file is smaller; skips already-efficient codecs (HEVC, AV1, VP9)
4. **History tracking** (`AV1_Conversion_History.json`) - Remembers processed files to skip re-downloading; tracks cumulative space saved across runs
5. **Interruption handling** - Graceful Ctrl+C/shutdown cleanup; removes temp files and saves history on exit

## Key Conventions

- **Encoder configs**: Each encoder in `$encoderConfigs` array has `Name`, `Type`, `Args` (ffmpeg parameters), and `Codec` properties
- **File attributes**: Uses Windows file attribute flags (`0x1000`, `0x400000`) to detect OneDrive cloud-only status
- **Output format**: Always `.mkv` container regardless of input format
- **Logging**: All operations logged to `AV1_Conversion_Log.txt` in OneDrive root via `Write-Log` function
- **History**: Processed files tracked in `AV1_Conversion_History.json` with status and cumulative savings
- **Script-scope variables**: `$script:ConversionHistory`, `$script:CurrentTempFile`, `$script:HistoryFile` for cleanup handler access

## Key Functions

- `Get-ConversionHistory` / `Save-ConversionHistory` - Load/save JSON history file
- `Add-ProcessedFile` - Record file in history with status (converted, skipped-low-bitrate, kept-original)
- `Test-FileProcessed` - Check if file already in history (skip without ffprobe/download)
- `Invoke-CleanupOnExit` - Cleanup handler for Ctrl+C (removes temp file, saves history)

## Testing Changes

```powershell
# Test encoder detection only (creates temp file, tests each encoder)
$encoder = Get-BestAV1Encoder
$encoder.Type  # Shows which encoder was selected

# Test on a single file (modify $videos array to contain one item)
# Or temporarily set $MinSizeMB higher to limit scope

# Clear history to reprocess all files
Remove-Item "$env:OneDrive\AV1_Conversion_History.json" -ErrorAction SilentlyContinue
```

## FFmpeg Encoder Parameters

When modifying encoder settings, refer to ffmpeg documentation for each encoder:
- `av1_nvenc`: NVIDIA-specific (tune, cq, spatial-aq)
- `av1_amf`: AMD-specific (quality, rc, qp_i/qp_p)  
- `av1_qsv`: Intel-specific (preset 1-7, extbrc, look_ahead_depth)
- `hevc_mf`: MediaFoundation (rate_control, quality percentage)
- `libsvtav1`: Software (preset 0-13, crf, svtav1-params)
