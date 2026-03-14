Add-Type -AssemblyName System.Drawing

function New-ContextMenuIcon($text, [System.Drawing.Color]$bgColor, $outputPath) {
    $sizes = @(16, 20, 24, 32, 48, 64, 256)
    $entries = [System.Collections.ArrayList]::new()

    foreach ($sz in $sizes) {
        [int]$s = $sz
        $bmp = [System.Drawing.Bitmap]::new($s, $s)
        $bmp.SetResolution(96, 96)
        $g = [System.Drawing.Graphics]::FromImage($bmp)
        $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::HighQuality
        $g.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::AntiAliasGridFit
        $g.Clear([System.Drawing.Color]::Transparent)

        # Filled rounded rect
        $brush = [System.Drawing.SolidBrush]::new($bgColor)
        [int]$r = [math]::Max(2, [math]::Floor($s * 0.22))
        [int]$d = $r * 2
        [int]$w = $s - 1
        $gp = [System.Drawing.Drawing2D.GraphicsPath]::new()
        $gp.AddArc(0, 0, $d, $d, 180, 90)
        $gp.AddArc($w - $d, 0, $d, $d, 270, 90)
        $gp.AddArc($w - $d, $w - $d, $d, $d, 0, 90)
        $gp.AddArc(0, $w - $d, $d, $d, 90, 90)
        $gp.CloseFigure()
        $g.FillPath($brush, $gp)

        # Center text
        [int]$fontSize = [math]::Max(7, [math]::Floor($s * 0.55))
        $font = [System.Drawing.Font]::new("Segoe UI", $fontSize, [System.Drawing.FontStyle]::Bold, [System.Drawing.GraphicsUnit]::Pixel)
        $tb = [System.Drawing.SolidBrush]::new([System.Drawing.Color]::White)
        $strFmt = [System.Drawing.StringFormat]::new()
        $strFmt.Alignment = [System.Drawing.StringAlignment]::Center
        $strFmt.LineAlignment = [System.Drawing.StringAlignment]::Center
        $tr = [System.Drawing.RectangleF]::new(0, 0, $s, $s)
        $g.DrawString($text, $font, $tb, $tr, $strFmt)

        $g.Dispose(); $font.Dispose(); $brush.Dispose(); $tb.Dispose(); $gp.Dispose()

        $ms = [System.IO.MemoryStream]::new()
        $bmp.Save($ms, [System.Drawing.Imaging.ImageFormat]::Png)
        [byte[]]$pngBytes = $ms.ToArray()
        $ms.Dispose(); $bmp.Dispose()

        [void]$entries.Add(@{ Size = $s; Data = $pngBytes })
    }

    # Write ICO
    $icoStream = [System.IO.File]::Create($outputPath)
    $bw = [System.IO.BinaryWriter]::new($icoStream)
    $bw.Write([uint16]0)
    $bw.Write([uint16]1)
    $bw.Write([uint16]$entries.Count)

    [int]$offset = 6 + 16 * $entries.Count
    foreach ($e in $entries) {
        $bw.Write([byte]$(if ($e.Size -ge 256) { 0 } else { $e.Size }))
        $bw.Write([byte]$(if ($e.Size -ge 256) { 0 } else { $e.Size }))
        $bw.Write([byte]0); $bw.Write([byte]0)
        $bw.Write([uint16]1); $bw.Write([uint16]32)
        $bw.Write([uint32]$e.Data.Length)
        $bw.Write([uint32]$offset)
        $offset += $e.Data.Length
    }
    foreach ($e in $entries) { $bw.Write($e.Data) }
    $bw.Close(); $icoStream.Close()
    Write-Host "Created: $outputPath"
}

function New-LogoPng([int]$size, [System.Drawing.Color]$bgColor, $text, $outputPath) {
    $bmp = [System.Drawing.Bitmap]::new($size, $size)
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::HighQuality
    $g.Clear($bgColor)
    [int]$fontSize = [math]::Floor($size * 0.45)
    $font = [System.Drawing.Font]::new("Segoe UI", $fontSize, [System.Drawing.FontStyle]::Bold, [System.Drawing.GraphicsUnit]::Pixel)
    $brush = [System.Drawing.SolidBrush]::new([System.Drawing.Color]::White)
    $sf = [System.Drawing.StringFormat]::new()
    $sf.Alignment = [System.Drawing.StringAlignment]::Center
    $sf.LineAlignment = [System.Drawing.StringAlignment]::Center
    $g.DrawString($text, $font, $brush, [System.Drawing.RectangleF]::new(0,0,$size,$size), $sf)
    $g.Dispose()
    $bmp.Save($outputPath, [System.Drawing.Imaging.ImageFormat]::Png)
    $bmp.Dispose()
    Write-Host "Created: $outputPath"
}

$distDir = Join-Path $PSScriptRoot "dist"
if (-not (Test-Path $distDir)) { New-Item -ItemType Directory -Path $distDir | Out-Null }

$claudeColor = [System.Drawing.Color]::FromArgb(255, 217, 119, 6)
New-ContextMenuIcon "C" $claudeColor (Join-Path $distDir "claude.ico")

$codexColor = [System.Drawing.Color]::FromArgb(255, 16, 163, 127)
New-ContextMenuIcon "X" $codexColor (Join-Path $distDir "codex.ico")

$assetsDir = Join-Path $distDir "assets"
if (-not (Test-Path $assetsDir)) { New-Item -ItemType Directory -Path $assetsDir | Out-Null }

$blendColor = [System.Drawing.Color]::FromArgb(255, 80, 80, 80)
New-LogoPng 44  $blendColor "CLI" (Join-Path $assetsDir "logo44x44.png")
New-LogoPng 150 $blendColor "CLI" (Join-Path $assetsDir "logo150x150.png")
