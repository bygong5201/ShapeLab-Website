param(
  [string]$SourcePath = "additive-manufacturing-hero/assets/additive-blueprint-bg.png",
  [string]$OutputPath = "additive-manufacturing-hero/assets/additive-blueprint-bg-extended.jpg",
  [int]$Width = 1920,
  [int]$Height = 12000
)

$ErrorActionPreference = "Stop"

Add-Type -AssemblyName System.Drawing

function New-LinearGradientBrush($Rect, $Color1, $Color2, $Mode) {
  return New-Object System.Drawing.Drawing2D.LinearGradientBrush($Rect, $Color1, $Color2, $Mode)
}

$root = Split-Path -Parent $PSScriptRoot
$sourceFullPath = Join-Path $root $SourcePath
$outputFullPath = Join-Path $root $OutputPath

if (-not (Test-Path -LiteralPath $sourceFullPath)) {
  throw "Missing source image: $sourceFullPath"
}

$source = [System.Drawing.Image]::FromFile($sourceFullPath)
$canvas = New-Object System.Drawing.Bitmap($Width, $Height, [System.Drawing.Imaging.PixelFormat]::Format24bppRgb)
$graphics = [System.Drawing.Graphics]::FromImage($canvas)

try {
  $graphics.CompositingQuality = [System.Drawing.Drawing2D.CompositingQuality]::HighQuality
  $graphics.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
  $graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::HighQuality
  $graphics.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
  $graphics.Clear([System.Drawing.Color]::FromArgb(247, 251, 255))

  $scale = [Math]::Max($Width / $source.Width, 1080 / $source.Height)
  $tileWidth = [int][Math]::Ceiling($source.Width * $scale)
  $tileHeight = [int][Math]::Ceiling($source.Height * $scale)
  $overlap = 180
  $step = $tileHeight - $overlap

  for ($y = -60; $y -lt $Height; $y += $step) {
    $tileIndex = [Math]::Floor(($y + 60) / $step)
    $state = $graphics.Save()
    $offset = (($tileIndex % 4) - 1.5) * 120
    $drawX = [int][Math]::Floor(($Width - $tileWidth) / 2 + $offset)
    $drawY = [int]$y

    if ($tileIndex % 2 -eq 1) {
      $graphics.TranslateTransform($Width, 0)
      $graphics.ScaleTransform(-1, 1)
      $drawX = $Width - $drawX - $tileWidth
    }

    $dest = New-Object System.Drawing.Rectangle($drawX, $drawY, $tileWidth, $tileHeight)
    $graphics.DrawImage($source, $dest)
    $graphics.Restore($state)

    if ($y -gt 0) {
      $fadeRect = New-Object System.Drawing.Rectangle(0, $y, $Width, $overlap)
      $fadeBrush = New-LinearGradientBrush $fadeRect ([System.Drawing.Color]::FromArgb(90, 247, 251, 255)) ([System.Drawing.Color]::FromArgb(0, 247, 251, 255)) ([System.Drawing.Drawing2D.LinearGradientMode]::Vertical)
      $graphics.FillRectangle($fadeBrush, $fadeRect)
      $fadeBrush.Dispose()
    }
  }

  $blue = [System.Drawing.Color]::FromArgb(34, 35, 105, 170)
  $linePen = New-Object System.Drawing.Pen($blue, 1)
  try {
    for ($x = 0; $x -lt $Width; $x += 64) {
      $graphics.DrawLine($linePen, $x, 0, $x, $Height)
    }
    for ($y = 0; $y -lt $Height; $y += 64) {
      $graphics.DrawLine($linePen, 0, $y, $Width, $y)
    }
  }
  finally {
    $linePen.Dispose()
  }

  $veilRect = New-Object System.Drawing.Rectangle(0, 0, $Width, $Height)
  $veil = New-LinearGradientBrush $veilRect ([System.Drawing.Color]::FromArgb(54, 248, 251, 255)) ([System.Drawing.Color]::FromArgb(10, 248, 251, 255)) ([System.Drawing.Drawing2D.LinearGradientMode]::Horizontal)
  $graphics.FillRectangle($veil, $veilRect)
  $veil.Dispose()

  $extension = [System.IO.Path]::GetExtension($outputFullPath).ToLowerInvariant()
  if ($extension -eq ".jpg" -or $extension -eq ".jpeg") {
    $codec = [System.Drawing.Imaging.ImageCodecInfo]::GetImageEncoders() | Where-Object { $_.MimeType -eq "image/jpeg" }
    $encoderParams = New-Object System.Drawing.Imaging.EncoderParameters(1)
    $encoderParams.Param[0] = New-Object System.Drawing.Imaging.EncoderParameter([System.Drawing.Imaging.Encoder]::Quality, 88L)
    $canvas.Save($outputFullPath, $codec, $encoderParams)
    $encoderParams.Dispose()
  }
  else {
    $canvas.Save($outputFullPath, [System.Drawing.Imaging.ImageFormat]::Png)
  }
}
finally {
  $graphics.Dispose()
  $canvas.Dispose()
  $source.Dispose()
}

Write-Host "Created $outputFullPath ($Width x $Height)"
