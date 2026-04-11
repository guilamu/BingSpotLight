#Requires -RunAsAdministrator

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$BaseDir = Split-Path -Path $PSCommandPath -Parent
$SourceDir = Join-Path $BaseDir "source"
$RenderedDir = Join-Path $BaseDir "rendered"
$LogDir = Join-Path $BaseDir "logs"
$LogPath = Join-Path $LogDir "BingSpotlight.log"
$ConfigPath = Join-Path $BaseDir "config.json"
$SourceImagePath = Join-Path $SourceDir "bing_source.jpg"
$QrCodeImagePath = Join-Path $SourceDir "qrcode.png"
$CurrentDate = Get-Date -Format "yyyy-MM-dd"
$RenderedImagePath = Join-Path $RenderedDir ("lockscreen_{0}.jpg" -f $CurrentDate)
$OriginalImagePath = Join-Path $RenderedDir ("original_{0}.jpg" -f $CurrentDate)
$RegPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\PersonalizationCSP"

function Initialize-Folders {
    foreach ($path in @($BaseDir, $SourceDir, $RenderedDir, $LogDir)) {
        if (-not (Test-Path -LiteralPath $path)) {
            New-Item -ItemType Directory -Path $path -Force | Out-Null
        }
    }
}

function Write-Log {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,

        [ValidateSet("INFO", "WARN", "ERROR")]
        [string]$Level = "INFO"
    )

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $line = "[{0}] [{1}] {2}" -f $timestamp, $Level, $Message
    Add-Content -Path $LogPath -Value $line
}

function Get-ScreenResolution {
    try {
        $monitor = Get-CimInstance -ClassName Win32_VideoController |
            Where-Object { $_.CurrentHorizontalResolution -gt 0 } |
            Select-Object -First 1

        if ($monitor) {
            return [pscustomobject]@{
                Width  = [int]$monitor.CurrentHorizontalResolution
                Height = [int]$monitor.CurrentVerticalResolution
            }
        }
    }
    catch {
        Write-Log -Level "WARN" -Message ("Screen resolution detection failed: {0}" -f $_.Exception.Message)
    }
    return $null
}

function Get-DefaultConfig {
    return [ordered]@{
        Market = "fr-FR"
        Language = "fr"
        RetentionDays = 14
        RetryCount = 5
        RetryDelaySeconds = 15
    }
}

function Read-Config {
    $configData = Get-DefaultConfig

    if (Test-Path -LiteralPath $ConfigPath) {
        $rawConfig = Get-Content -LiteralPath $ConfigPath -Raw | ConvertFrom-Json

        foreach ($property in $rawConfig.PSObject.Properties) {
            $configData[$property.Name] = $property.Value
        }
    }

    try {
        $configData.RetentionDays = [int]$configData.RetentionDays
    }
    catch {
        throw "Invalid RetentionDays value in config.json."
    }

    try {
        $configData.RetryCount = [int]$configData.RetryCount
    }
    catch {
        throw "Invalid RetryCount value in config.json."
    }

    try {
        $configData.RetryDelaySeconds = [int]$configData.RetryDelaySeconds
    }
    catch {
        throw "Invalid RetryDelaySeconds value in config.json."
    }

    if ([string]::IsNullOrWhiteSpace([string]$configData.Market)) {
        $configData.Market = "fr-FR"
    }

    if ([string]::IsNullOrWhiteSpace([string]$configData.Language)) {
        $configData.Language = $configData.Market.Split('-')[0].ToLower()
    }

    if ($configData.RetentionDays -lt 1) {
        throw "RetentionDays must be greater than or equal to 1."
    }

    if ($configData.RetryCount -lt 1) {
        throw "RetryCount must be greater than or equal to 1."
    }

    if ($configData.RetryDelaySeconds -lt 1) {
        throw "RetryDelaySeconds must be greater than or equal to 1."
    }

    return [pscustomobject]$configData
}

function Invoke-WithRetry {
    param(
        [Parameter(Mandatory = $true)]
        [scriptblock]$ScriptBlock,

        [int]$MaxAttempts,
        [int]$DelaySeconds,
        [string]$OperationName = "Operation"
    )

    for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
        try {
            return & $ScriptBlock
        }
        catch {
            if ($attempt -eq $MaxAttempts) {
                throw
            }

            Write-Log -Level "WARN" -Message (
                "{0} failed on attempt {1}/{2}: {3}" -f $OperationName, $attempt, $MaxAttempts, $_.Exception.Message
            )
            Start-Sleep -Seconds $DelaySeconds
        }
    }
}

function Get-BingMetadata {
    param(
        [Parameter(Mandatory = $true)]
        [pscustomobject]$Config
    )

    $apiUrl = "https://www.bing.com/HPImageArchive.aspx?format=js&idx=0&n=1&mkt=$($Config.Market)"
    $response = Invoke-WithRetry -OperationName "Bing metadata request" -MaxAttempts $Config.RetryCount -DelaySeconds $Config.RetryDelaySeconds -ScriptBlock {
        Invoke-RestMethod -Uri $apiUrl -UseBasicParsing
    }

    if (-not $response.images -or $response.images.Count -lt 1) {
        throw "No image metadata returned by Bing API."
    }

    $image = $response.images[0]
    $title = if ([string]::IsNullOrWhiteSpace($image.title)) { "Bing" } else { $image.title.Trim() }
    $copyright = if ([string]::IsNullOrWhiteSpace($image.copyright)) { "" } else { $image.copyright.Trim() }

    $rawUrl = $image.url -replace '1920x1080', 'UHD'

    return [pscustomobject]@{
        ImageUrl = "https://www.bing.com$rawUrl"
        Title = $title
        Copyright = $copyright
    }
}

function Get-ImageDescription {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Copyright
    )

    # Strip the trailing copyright parenthetical, e.g. " (© Agency/Author)"
    # Matches the last " (" before a copyright symbol or the end of the string
    $stripped = $Copyright -replace '\s*\([^)]*[\u00a9\(C\)©].*?\)\s*$', ''
    $stripped = $stripped.Trim().TrimEnd(',')

    if ([string]::IsNullOrWhiteSpace($stripped)) {
        return $Copyright
    }

    return $stripped
}

function Get-GoogleSearchUrl {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Description,

        [Parameter(Mandatory = $true)]
        [string]$Language
    )

    $lang = $Language.ToLower()
    if ([string]::IsNullOrWhiteSpace($lang) -or $lang.Length -ne 2) {
        $lang = "en"
    }

    $encoded = [System.Uri]::EscapeDataString($Description)
    return "https://www.google.com/search?q={0}&hl={1}" -f $encoded, $lang
}

function Download-QrCodeImage {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Data,

        [Parameter(Mandatory = $true)]
        [string]$DestinationPath,

        [int]$SizePx = 150
    )

    $encoded = [System.Uri]::EscapeDataString($Data)
    $apiUrl = "https://api.qrserver.com/v1/create-qr-code/?size={0}x{0}&format=png&margin=4&data={1}" -f $SizePx, $encoded
    Invoke-WebRequest -Uri $apiUrl -OutFile $DestinationPath -UseBasicParsing
}

function Download-SourceImage {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Url,

        [Parameter(Mandatory = $true)]
        [string]$DestinationPath,

        [Parameter(Mandatory = $true)]
        [pscustomobject]$Config
    )

    Invoke-WithRetry -OperationName "Bing image download" -MaxAttempts $Config.RetryCount -DelaySeconds $Config.RetryDelaySeconds -ScriptBlock {
        Invoke-WebRequest -Uri $Url -OutFile $DestinationPath -UseBasicParsing
    } | Out-Null
}

function New-JpegEncoderParameters {
    param(
        [long]$Quality = 95
    )

    $encoderParams = [System.Drawing.Imaging.EncoderParameters]::new(1)
    $encoderParams.Param[0] = [System.Drawing.Imaging.EncoderParameter]::new(
        [System.Drawing.Imaging.Encoder]::Quality,
        $Quality
    )

    return $encoderParams
}

function Render-LockScreenImage {
    param(
        [Parameter(Mandatory = $true)]
        [string]$InputPath,

        [Parameter(Mandatory = $true)]
        [string]$OutputPath,

        [Parameter(Mandatory = $true)]
        [string]$OriginalPath,

        [Parameter(Mandatory = $true)]
        [string]$Title,

        [Parameter(Mandatory = $true)]
        [string]$Subtitle,

        [string]$QrCodePath,

        [int]$TargetWidth = 0,
        [int]$TargetHeight = 0
    )

    Add-Type -AssemblyName System.Drawing

    $sourceBitmap = $null
    $bitmap = $null
    $graphics = $null
    $bgBrush = $null
    $subtitleBrush = $null
    $fontTitle = $null
    $fontSubtitle = $null
    $encoderParams = $null
    $qrBitmap = $null

    try {
        $sourceBitmap = [System.Drawing.Bitmap]::FromFile($InputPath)

        if ($TargetWidth -gt 0 -and $TargetHeight -gt 0 -and
            ($sourceBitmap.Width -ne $TargetWidth -or $sourceBitmap.Height -ne $TargetHeight)) {
            $bitmap = [System.Drawing.Bitmap]::new($TargetWidth, $TargetHeight)
            $resizeGraphics = [System.Drawing.Graphics]::FromImage($bitmap)
            $resizeGraphics.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
            $resizeGraphics.CompositingQuality = [System.Drawing.Drawing2D.CompositingQuality]::HighQuality
            $resizeGraphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
            $resizeGraphics.DrawImage($sourceBitmap, 0, 0, $TargetWidth, $TargetHeight)
            $resizeGraphics.Dispose()
            $sourceBitmap.Dispose()
            $sourceBitmap = $null
        }
        else {
            $bitmap = $sourceBitmap
            $sourceBitmap = $null
        }

        $jpegEncoder = [System.Drawing.Imaging.ImageCodecInfo]::GetImageEncoders() |
            Where-Object { $_.MimeType -eq "image/jpeg" } |
            Select-Object -First 1

        if (-not $jpegEncoder) {
            throw "JPEG encoder not found."
        }

        $encoderParams = New-JpegEncoderParameters -Quality 95

        # --- Save the clean original image (no banner) ---
        $bitmap.Save($OriginalPath, $jpegEncoder, $encoderParams)

        # --- Draw the banner on top of the image for the lock screen ---
        $width = $bitmap.Width
        $height = $bitmap.Height

        $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
        $graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
        $graphics.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::AntiAliasGridFit

        $paddingX = [Math]::Max([int]($width * 0.014), 24)
        $bannerHeight = [Math]::Max([int]($height * 0.115), 120)
        $titleFontSize = [Math]::Max([int]($height * 0.024), 20)
        $subtitleFontSize = [Math]::Max([int]($height * 0.013), 11)
        $bannerTop = $height - $bannerHeight
        $titleTop = $bannerTop + [Math]::Max([int]($bannerHeight * 0.14), 12)
        $subtitleTop = $bannerTop + [Math]::Max([int]($bannerHeight * 0.58), 58)

        $bgBrush = [System.Drawing.SolidBrush]::new([System.Drawing.Color]::FromArgb(170, 0, 0, 0))
        $subtitleBrush = [System.Drawing.SolidBrush]::new([System.Drawing.Color]::FromArgb(220, 220, 220, 220))
        $fontTitle = [System.Drawing.Font]::new("Segoe UI", $titleFontSize, [System.Drawing.FontStyle]::Bold)
        $fontSubtitle = [System.Drawing.Font]::new("Segoe UI", $subtitleFontSize, [System.Drawing.FontStyle]::Regular)

        $graphics.FillRectangle($bgBrush, 0, $bannerTop, $width, $bannerHeight)
        $graphics.DrawString($Title, $fontTitle, [System.Drawing.Brushes]::White, [System.Drawing.PointF]::new($paddingX, $titleTop))
        $graphics.DrawString($Subtitle, $fontSubtitle, $subtitleBrush, [System.Drawing.PointF]::new($paddingX, $subtitleTop))

        if ($QrCodePath -and (Test-Path -LiteralPath $QrCodePath)) {
            $qrBitmap = [System.Drawing.Bitmap]::FromFile($QrCodePath)
            $qrSize = [Math]::Max([int]($bannerHeight * 0.72), 64)
            $qrX = $width - $paddingX - $qrSize
            $qrY = $bannerTop + [int](($bannerHeight - $qrSize) / 2)
            $graphics.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
            $graphics.DrawImage($qrBitmap, $qrX, $qrY, $qrSize, $qrSize)
        }

        $bitmap.Save($OutputPath, $jpegEncoder, $encoderParams)
    }
    finally {
        if ($qrBitmap) { $qrBitmap.Dispose() }
        if ($encoderParams) { $encoderParams.Dispose() }
        if ($fontSubtitle) { $fontSubtitle.Dispose() }
        if ($fontTitle) { $fontTitle.Dispose() }
        if ($subtitleBrush) { $subtitleBrush.Dispose() }
        if ($bgBrush) { $bgBrush.Dispose() }
        if ($graphics) { $graphics.Dispose() }
        if ($bitmap) { $bitmap.Dispose() }
        if ($sourceBitmap) { $sourceBitmap.Dispose() }
    }
}

function Set-LockScreenImage {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ImagePath
    )

    if (-not (Test-Path -LiteralPath $RegPath)) {
        New-Item -Path $RegPath -Force | Out-Null
    }

    Set-ItemProperty -Path $RegPath -Name "LockScreenImagePath" -Value $ImagePath -Force
    Set-ItemProperty -Path $RegPath -Name "LockScreenImageUrl" -Value $ImagePath -Force
    Set-ItemProperty -Path $RegPath -Name "LockScreenImageStatus" -Value 1 -Force
}

function Get-LatestRenderedImage {
    if (-not (Test-Path -LiteralPath $RenderedDir)) {
        return $null
    }

    return Get-ChildItem -Path $RenderedDir -Filter "lockscreen_*.jpg" -File |
        Sort-Object -Property LastWriteTime -Descending |
        Select-Object -First 1
}



function Restore-LatestRenderedImage {
    $latestImage = Get-LatestRenderedImage

    if (-not $latestImage) {
        return $false
    }

    Set-LockScreenImage -ImagePath $latestImage.FullName
    Write-Log -Level "WARN" -Message ("Bing unavailable. Restored latest rendered image: {0}" -f $latestImage.FullName)
    Write-Host ("Bing unavailable. Reusing existing image: {0}" -f $latestImage.Name)
    return $true
}

function Remove-OldRenderedImages {
    param(
        [int]$DaysToKeep
    )

    if (-not (Test-Path -LiteralPath $RenderedDir)) {
        return
    }

    $limit = (Get-Date).AddDays(-$DaysToKeep)

    Get-ChildItem -Path $RenderedDir -Filter "lockscreen_*.jpg" -File |
        Where-Object { $_.LastWriteTime -lt $limit } |
        Remove-Item -Force -ErrorAction SilentlyContinue

    Get-ChildItem -Path $RenderedDir -Filter "original_*.jpg" -File |
        Where-Object { $_.LastWriteTime -lt $limit } |
        Remove-Item -Force -ErrorAction SilentlyContinue
}

try {
    Initialize-Folders
    $config = Read-Config

    Write-Log -Message ("Execution started. RetentionDays={0}" -f $config.RetentionDays)

    try {
        $metadata = Get-BingMetadata -Config $config
        Write-Log -Message ("Metadata retrieved: {0}" -f $metadata.Title)

        Download-SourceImage -Url $metadata.ImageUrl -DestinationPath $SourceImagePath -Config $config
        Write-Log -Message "Source image downloaded."

        $qrPath = $null
        try {
            $imageDescription = Get-ImageDescription -Copyright $metadata.Copyright
            $searchUrl = Get-GoogleSearchUrl -Description $imageDescription -Language $config.Language
            Download-QrCodeImage -Data $searchUrl -DestinationPath $QrCodeImagePath -SizePx 200
            $qrPath = $QrCodeImagePath
            Write-Log -Message ("QR code generated for: {0}" -f $searchUrl)
        }
        catch {
            Write-Log -Level "WARN" -Message ("QR code generation skipped: {0}" -f $_.Exception.Message)
        }

        $screenRes = Get-ScreenResolution
        $renderParams = @{
            InputPath    = $SourceImagePath
            OutputPath   = $RenderedImagePath
            OriginalPath = $OriginalImagePath
            Title        = $metadata.Title
            Subtitle     = $metadata.Copyright
            QrCodePath   = $qrPath
        }
        if ($screenRes) {
            $renderParams.TargetWidth  = $screenRes.Width
            $renderParams.TargetHeight = $screenRes.Height
            Write-Log -Message ("Screen resolution detected: {0}x{1}" -f $screenRes.Width, $screenRes.Height)
        }
        else {
            Write-Log -Level "WARN" -Message "Screen resolution not detected, using source image size."
        }
        Render-LockScreenImage @renderParams

        Write-Log -Message ("Original image saved (clean): {0}" -f $OriginalImagePath)
        Write-Log -Message ("Lock screen image created: {0}" -f $RenderedImagePath)

        Set-LockScreenImage -ImagePath $RenderedImagePath
        Write-Log -Message "Registry updated successfully."

        Write-Host ("Lock screen updated: {0}" -f $metadata.Title)
    }
    catch {
        Write-Log -Level "ERROR" -Message ("Online refresh failed: {0}" -f $_.Exception.Message)

        if (-not (Restore-LatestRenderedImage)) {
            throw
        }
    }

    Remove-OldRenderedImages -DaysToKeep $config.RetentionDays
    Write-Log -Message "Cleanup finished."
}
catch {
    try {
        Initialize-Folders
        Write-Log -Level "ERROR" -Message $_.Exception.Message
    }
    catch {
    }

    Write-Error $_.Exception.Message
    exit 1
}