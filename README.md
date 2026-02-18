# OneDrive Video Converter (AV1/HEVC)

Automatically convert large video files to AV1 or HEVC format to reduce OneDrive storage usage by 30-70%.

## What This Script Does

- Scans your OneDrive folder for video files larger than 250MB
- Auto-detects the best available hardware encoder
- Priority: NVIDIA AV1 > AMD AV1 > Intel AV1 > Microsoft MediaFoundation HEVC (Qualcomm HW) > Software AV1
- Converts videos to the most efficient codec available on your hardware
- Compares file sizes and only keeps the converted file if it is smaller
- Preserves original filename (changes extension to `.mkv`)
- Converts high-bitrate files regardless of codec (even HEVC/AV1 with high bitrate)
- Skips low-bitrate files (already well-compressed, unlikely to benefit)
- Checks codec/bitrate before downloading cloud files (saves bandwidth)
- Checks disk space before conversion (prevents "No space left" errors)
- Handles OneDrive cloud files (downloads only when needed)
- **Remembers processed files** - won't re-download on subsequent runs
- **Tracks cumulative space saved** across all script runs
- **Auto-frees local storage** after conversion (sets files to cloud-only)
- **Safe interruption** - Ctrl+C cleans up temp files and saves progress

## Expected Storage Savings

| Original Codec | Typical Reduction |
|----------------|-------------------|
| H.264 (most common) | 50-70% smaller |
| MPEG-4 | 60-80% smaller |
| ProRes/DNxHD | 80-90% smaller |
| HEVC/H.265 (high bitrate) | 20-40% smaller |
| HEVC/H.265 (low bitrate) | Skipped |

Example: A 1GB H.264 video typically becomes 300-500MB in AV1.

## Supported Hardware Encoders

| GPU/Processor | Encoder | Output Codec | Speed |
|---------------|---------|--------------|-------|
| NVIDIA RTX 30/40 series | NVENC | AV1 | Fastest |
| AMD RX 6000/7000/8000 series | AMF | AV1 | Fast |
| Intel Arc / 11th+ Gen iGPU | Quick Sync | AV1 | Fast |
| Microsoft MediaFoundation (uses Qualcomm HW on Snapdragon) | hevc_mf | HEVC | Fast |
| Any (Software fallback) | SVT-AV1 | AV1 | Slower but works everywhere |

**Note for Qualcomm Snapdragon X Elite users:** AV1 hardware encoding is not yet available on Snapdragon. The script automatically falls back to HEVC hardware encoding, which still provides 30-40% compression vs H.264 with fast hardware acceleration.

## Prerequisites

- Windows 10/11
- OneDrive installed and synced
- FFmpeg with AV1 encoder support
- (Recommended) GPU with AV1 hardware encoding support

## Quick Setup (Windows 11)

### Step 1: Install FFmpeg via Winget

Open PowerShell or Terminal as Administrator and run:

```powershell
winget install Gyan.FFmpeg
```

Or download manually from: https://www.gyan.dev/ffmpeg/builds/

### Step 2: Verify Installation

```powershell
ffmpeg -version
```

### Step 3: Check Available AV1 Encoders

```powershell
ffmpeg -encoders 2>$null | Select-String "av1"
```

You should see encoders like `av1_nvenc`, `av1_amf`, `av1_qsv`, or `libsvtav1`.

### Step 4: Download the Script

Save `Convert-ToAV1.ps1` to any folder (e.g., your OneDrive folder or Documents).

### Step 5: Run the Script

```powershell
# Navigate to script location
cd "C:\Users\YourUsername\OneDrive"

# Run the script
.\Convert-ToAV1.ps1
```

## Configuration Options

Edit the script to customize these settings:

```powershell
# Minimum file size to process (default: 250MB)
$MinSizeMB = 250

# Minimum bitrate threshold in kbps (skip files already below this)
# Files with lower bitrate are already well-compressed and unlikely to benefit from re-encoding
# Note: Even efficient codecs (HEVC, AV1) will be converted if above this threshold
$MinBitrateKbps = 1500

# Minimum savings percentage (skip if test encode shows savings below this)
# Prevents wasting time on files that won't compress well (e.g., DJI drone HEVC)
$MinSavingsPercent = 10

# Duration in seconds for test encode (verifies actual savings before full conversion)
$TestEncodeDuration = 15

# Number of files to prefetch ahead (download while converting current file)
$PrefetchCount = 3

# Video extensions to process
$VideoExtensions = @("*.mp4", "*.mov", "*.avi", "*.wmv", "*.flv", "*.webm", "*.m4v", "*.mpg", "*.mpeg")

# Efficient codecs (may still convert if high bitrate)
$EfficientCodecs = @("hevc", "h265", "av1", "vp9")
```

## Encoder Quality Settings

The script uses optimized settings for each encoder:

### AMD AMF (AV1)
```
-quality balanced -rc cqp -qp_i 38 -qp_p 40
```
- Balanced preset for good speed/quality tradeoff
- CQP (Constant QP) for consistent quality
- Higher QP values (38/40) for better compression

### NVIDIA NVENC (AV1)
```
-tune hq -cq 32 -rc-lookahead 32 -spatial-aq 1 -aq-strength 8
```
- CQ mode for quality-based encoding
- Spatial AQ allocates bits to complex areas

### Intel QSV (AV1)
```
-preset 4 -extbrc 1 -look_ahead_depth 40 -adaptive_i 1 -adaptive_b 1
```
- Lookahead for better frame analysis
- Adaptive keyframe placement

### Microsoft MediaFoundation HEVC (uses Qualcomm HW on Snapdragon)
```
-rate_control quality -quality 70
```
- Quality-focused rate control for Snapdragon X Elite
- Hardware accelerated HEVC encoding

### SVT-AV1 (Software)
```
-preset 5 -crf 30
```
- Preset 5 balances speed and quality
- CRF 30 targets good compression

## Log File

The script creates `AV1_Conversion_Log.txt` in your OneDrive folder with:
- Processing status for each file
- Original and new file sizes
- Compression percentage
- Total space saved

## History Tracking

The script creates `AV1_Conversion_History.json` in your OneDrive folder to:
- Remember which files have already been scanned/processed
- Skip re-downloading cloud files that were already converted or skipped
- Track cumulative space saved across all script runs
- Persist history between script launches

This means running the script again won't re-download files that were already processed, saving bandwidth and time.

## Troubleshooting

### FFmpeg not found
```powershell
# Reinstall FFmpeg
winget install Gyan.FFmpeg

# Or add to PATH manually
$env:Path += ";C:\ffmpeg\bin"
```

### No AV1 encoder available
Your GPU may not support AV1 hardware encoding. The script will fall back to SVT-AV1 software encoder (slower but works on any PC).

### Error opening input file
This usually means the file is cloud-only in OneDrive. The script automatically waits for downloads, but you can also:
```powershell
# Force download by right-clicking file in Explorer > "Always keep on this device"
```

### Encoding is slow
- Hardware encoding: Should be 50-100x realtime
- Software encoding (SVT-AV1): Expect 1-5x realtime depending on CPU

### Script will not run (Execution Policy)
```powershell
# Allow running local scripts
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

## File Handling

| Scenario | Action |
|----------|--------|
| Output file is smaller | Delete original, keep converted, free local storage |
| Output file is larger/same | Delete converted, keep original |
| Output file already exists | Skip (will not re-encode) |
| File already in history | Skip (won't re-download cloud files) |
| Cloud-only file | Check codec/bitrate first, download only if conversion needed |
| Efficient codec + high bitrate | Test encode first, convert only if savings verified |
| Any codec + low bitrate (< 1500 kbps) | Skip (already well-compressed) |
| Test encode shows < 10% savings | Skip (e.g., DJI drone HEVC that won't compress well) |
| Insufficient disk space | Skip (prevents conversion errors) |

## Running Periodically

### Option 1: Manual
Run the script whenever you have new videos to convert.

### Option 2: Task Scheduler
1. Open Task Scheduler
2. Create Basic Task
3. Set trigger (e.g., weekly)
4. Action: Start a program
   - Program: `powershell.exe`
   - Arguments: `-ExecutionPolicy Bypass -File "C:\Users\YourUsername\OneDrive\Convert-ToAV1.ps1"`

## Performance Tips

1. Process local files first - The script automatically prioritizes locally cached files
2. Run overnight - Large video collections may take hours
3. Check available disk space - Temporary files need space during conversion
4. Close other GPU-intensive apps - Hardware encoders share GPU resources

## Safety Features

- Never deletes original if conversion fails
- Never deletes original if output is larger
- **Test encodes 15s segment first** to verify actual savings before full conversion
- Skips files where test shows < 10% savings (e.g., DJI drone HEVC)
- Skips already converted files
- Skips low-bitrate files that won't benefit from re-encoding
- Skips files already processed in previous runs (no re-download)
- Checks disk space before starting conversion
- Checks codec/bitrate before downloading cloud files (saves time and bandwidth)
- Auto-frees local storage after conversion (OneDrive Files On-Demand)
- **Graceful interruption handling** (Ctrl+C or shutdown):
  - Removes incomplete temp files automatically
  - Saves conversion history before exit
  - Progress is preserved - just run again to continue
- Detailed logging for review
- Handles OneDrive sync gracefully

## License

Free to use and modify.

## Credits

- FFmpeg: https://ffmpeg.org/
- SVT-AV1: https://gitlab.com/AOMediaCodec/SVT-AV1
- AMD AMF: https://github.com/GPUOpen-LibrariesAndSDKs/AMF
