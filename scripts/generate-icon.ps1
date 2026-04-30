$ErrorActionPreference = "Stop"

Add-Type -AssemblyName System.Drawing

$projectRoot = Resolve-Path (Join-Path $PSScriptRoot "..")
$assetRoot = Join-Path $projectRoot "app\assets"
$iconPath = Join-Path $assetRoot "pdf-editor-icon.ico"
$source32Base64 = @"
iVBORw0KGgoAAAANSUhEUgAAACAAAAAgCAYAAABzenr0AAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAARNSURBVFhHxVf/U1RVFN+f6l/oj9jdt0Fo7ttddhLYVgzbpjVUHEoKJEoRMgMl5YuWEUboMiAWJRaRopmLOI6wghLgl7CaSKaxNPui2KDbWjr5w6c5B95r37vv7Sw2DWfmM3PvPeedc+6958t9FkuSFAwGH5Akt89qd4askhyxSfKE1S7HCDyW5AjxSIZk9d/fN6WleR+yOlzNVskZtdqdSAok63A107d6fUmT3+9/0CbJNTZJvi0YSBL0LekgXXr9CYk8t9nlYb3C+wXpSvo0rCluu83uvKpX8l9BOm2pHpvenoamd56ccWf6Irz5QRde3d4k8MxAuk1Pgu98Fsfe2PU5xu6A8VxppcA3A9kwjAkKFr1wImxt26c6MPjbTSwvWifImIFsaYzz0c8y2qUUD0o2bsXSVS9hWWEpTl+PwpOxRJAzAtnSXAXnuYGgAq/vKeQWrMGKojJkB/PhSHELMlVNbXwt/kAesnKWkRFBRgOHq5mNU9UyKzIZ2UvxyfAY+n+aRMfgGbT3DeHQhYvIfaGU+b4ly7El1I7D33yPruELOHXtFnZ0fobwt5cwdCPG45Q0r6CXITmjXDFnyqsg8PAjXhy/9AvWVtfD7nBpeIuDz6Kt9yQiV2+gtrUDgRWFvOOGjw+htLqeZea7fNh36ize/uigoFsB2bZwbTdg5q0ux+GvJjRr5BSlXv+VSZTV7RB2V1BehT3HBtT5Ao8fpyejwgbiELLMNBE9g1Prw/5hdb7o6XyEx3/AroO9mCdnCfIE12NPYODXKXVOgTp6664gp0KSI5R+EwLD7oQn80mM3LyDtAUZfMSkuKBskyCnx7nY3+qOKXC7z48LMgrItmW6pYpMwnvHB1G3uwMnf/6dlen5FJTtJ4Y0ayQre7N53Dl0HsUVdcJ3/0KOJXSA7pAUmhUZIwd6vvuR07B8WyNnRuJ0lP8wvQLCG+93crT3XLyMzMW5At8IfZevIT0rwA54fQGBH4/pKzAJwkfdj3MAzXNmYnVFDadczjOrBDk9Rqb+ROr8hcK6ISgIzdJwZcl6dH4xpplTIFY0NJumFTlNBSh+zUx2BiHTQtR0IIxzt++hpmWvukaZsXdgFN1fjhvGRV5xOfNpTO2arm5k6i92Xi9L4EJkVIrJazJOne5s7J5a+90ZOVwDKB2pSFFZfrGyTo36ynda8frOPTwuXL9Z7ZZtRyOCcbUUmzUjpd/vPHBUXdtQH+Jyq8ypOe3u6edOSM5Q4Xp3fxgtR05wtaQqeCZ6F/kvbxAdUJpRonbsSE3XzCktqeG4F+Zo1inV1tU28AmUbNqGoteq+bRYh0HnFNoxUbIPko2NrdwZ4xVTulG9oCallzeC8CAhSvZJRrulY/909Gs8/0oV1m55C31XrmPN5u2CrBFMn2REs3mUFlfUoiXch13dvQisLBL4Rkj4KFVoTp/lCs3pj4lCc/prFk9z9nOqp//r9/wfQtP3O42Xv4kAAAAASUVORK5CYII=
"@

function New-Color([int] $alpha, [string] $hex) {
    $value = $hex.TrimStart("#")
    return [System.Drawing.Color]::FromArgb(
        $alpha,
        [Convert]::ToInt32($value.Substring(0, 2), 16),
        [Convert]::ToInt32($value.Substring(2, 2), 16),
        [Convert]::ToInt32($value.Substring(4, 2), 16)
    )
}

function Draw-RotatedEllipse(
    [System.Drawing.Graphics] $graphics,
    [float] $centerX,
    [float] $centerY,
    [float] $radiusX,
    [float] $radiusY,
    [float] $angle,
    [System.Drawing.Pen] $pen
) {
    $path = New-Object System.Drawing.Drawing2D.GraphicsPath
    $matrix = New-Object System.Drawing.Drawing2D.Matrix
    try {
        $path.AddEllipse(-$radiusX, -$radiusY, 2 * $radiusX, 2 * $radiusY)
        $matrix.Rotate($angle)
        $matrix.Translate($centerX, $centerY, [System.Drawing.Drawing2D.MatrixOrder]::Append)
        $path.Transform($matrix)
        $graphics.DrawPath($pen, $path)
    }
    finally {
        $matrix.Dispose()
        $path.Dispose()
    }
}

function Fill-Dot(
    [System.Drawing.Graphics] $graphics,
    [float] $centerX,
    [float] $centerY,
    [float] $radius,
    [System.Drawing.Brush] $brush
) {
    $graphics.FillEllipse($brush, $centerX - $radius, $centerY - $radius, 2 * $radius, 2 * $radius)
}

function New-IconPng([int] $size) {
    $bitmap = New-Object System.Drawing.Bitmap $size, $size, ([System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
    $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
    try {
        $graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
        $graphics.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
        $graphics.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
        $graphics.Clear([System.Drawing.Color]::Transparent)

        $pad = [float]($size * 0.035)
        $circle = New-Object System.Drawing.RectangleF $pad, $pad, ([float]($size - 2 * $pad)), ([float]($size - 2 * $pad))
        $background = [System.Drawing.Drawing2D.LinearGradientBrush]::new(
            $circle,
            (New-Color 255 "#1b5965"),
            (New-Color 255 "#0e3340"),
            [single]45
        )
        try {
            $graphics.FillEllipse($background, $circle)
        }
        finally {
            $background.Dispose()
        }

        $center = [float]($size / 2)
        $orbitWidth = [Math]::Max(1.15, $size * 0.043)
        $orbitPen = New-Object System.Drawing.Pen (New-Color 230 "#b8f2ed"), $orbitWidth
        $orbitPen.LineJoin = [System.Drawing.Drawing2D.LineJoin]::Round
        try {
            Draw-RotatedEllipse $graphics $center $center ([float]($size * 0.365)) ([float]($size * 0.120)) -24 $orbitPen
            Draw-RotatedEllipse $graphics $center $center ([float]($size * 0.365)) ([float]($size * 0.120)) 64 $orbitPen
            Draw-RotatedEllipse $graphics $center $center ([float]($size * 0.330)) ([float]($size * 0.110)) -82 $orbitPen
        }
        finally {
            $orbitPen.Dispose()
        }

        $dotBrush = New-Object System.Drawing.SolidBrush (New-Color 255 "#effffb")
        $accentBrush = New-Object System.Drawing.SolidBrush (New-Color 255 "#7fdad4")
        try {
            Fill-Dot $graphics ([float]($size * 0.305)) ([float]($size * 0.285)) ([float]($size * 0.030)) $dotBrush
            Fill-Dot $graphics ([float]($size * 0.715)) ([float]($size * 0.355)) ([float]($size * 0.030)) $dotBrush
            Fill-Dot $graphics ([float]($size * 0.390)) ([float]($size * 0.730)) ([float]($size * 0.030)) $dotBrush
            Fill-Dot $graphics $center $center ([float]($size * 0.045)) $dotBrush
            Fill-Dot $graphics $center $center ([float]($size * 0.020)) $accentBrush
        }
        finally {
            $dotBrush.Dispose()
            $accentBrush.Dispose()
        }
    }
    finally {
        $graphics.Dispose()
    }

    $stream = New-Object System.IO.MemoryStream
    try {
        $bitmap.Save($stream, [System.Drawing.Imaging.ImageFormat]::Png)
        return $stream.ToArray()
    }
    finally {
        $stream.Dispose()
        $bitmap.Dispose()
    }
}

function Read-Source32Png() {
    return [Convert]::FromBase64String(($source32Base64 -replace "\s", ""))
}

function Write-UInt16Le([System.IO.BinaryWriter] $writer, [int] $value) {
    $writer.Write([uint16] $value)
}

function Write-UInt32Le([System.IO.BinaryWriter] $writer, [long] $value) {
    $writer.Write([uint32] $value)
}

$entries = @(
    @{ Size = 16; Png = New-IconPng 16 },
    @{ Size = 24; Png = New-IconPng 24 },
    @{ Size = 32; Png = Read-Source32Png },
    @{ Size = 48; Png = New-IconPng 48 },
    @{ Size = 64; Png = New-IconPng 64 },
    @{ Size = 128; Png = New-IconPng 128 },
    @{ Size = 256; Png = New-IconPng 256 }
)

$output = New-Object System.IO.MemoryStream
$writer = New-Object System.IO.BinaryWriter $output
try {
    Write-UInt16Le $writer 0
    Write-UInt16Le $writer 1
    Write-UInt16Le $writer $entries.Count

    $imageOffset = 6 + (16 * $entries.Count)
    foreach ($entry in $entries) {
        $sizeByte = if ($entry.Size -eq 256) { 0 } else { $entry.Size }
        $writer.Write([byte] $sizeByte)
        $writer.Write([byte] $sizeByte)
        $writer.Write([byte] 0)
        $writer.Write([byte] 0)
        Write-UInt16Le $writer 1
        Write-UInt16Le $writer 32
        Write-UInt32Le $writer $entry.Png.Length
        Write-UInt32Le $writer $imageOffset
        $imageOffset += $entry.Png.Length
    }

    foreach ($entry in $entries) {
        $writer.Write([byte[]] $entry.Png)
    }

    [System.IO.File]::WriteAllBytes($iconPath, $output.ToArray())
}
finally {
    $writer.Dispose()
    $output.Dispose()
}

Write-Host "Icon geschrieben: $iconPath"
