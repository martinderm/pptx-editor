param(
    [Parameter(Mandatory = $true)]
    [string]$InputPath,

    [Parameter(Mandatory = $true)]
    [string]$OutputDir,

    [int]$Width = 1920,
    [int]$Height = 1080
)

$ErrorActionPreference = 'Stop'

$ppt = $null
$pres = $null

try {
    $resolvedInput = (Resolve-Path -LiteralPath $InputPath).Path
    $outputDirFull = [System.IO.Path]::GetFullPath($OutputDir)
    if (-not (Test-Path -LiteralPath $outputDirFull)) {
        New-Item -ItemType Directory -Path $outputDirFull -Force | Out-Null
    }

    $ppt = New-Object -ComObject PowerPoint.Application
    $pres = $ppt.Presentations.Open($resolvedInput, -1, 0, 0)

    $slideCount = $pres.Slides.Count
    $exportedImages = @()

    for ($i = 1; $i -le $slideCount; $i++) {
        $slide = $pres.Slides.Item($i)
        $fileName = [string]::Format("slide_{0:D2}.png", $i)
        $imagePath = Join-Path $outputDirFull $fileName
        $slide.Export($imagePath, "PNG", $Width, $Height)
        $exportedImages += $imagePath
    }

    Write-Output (ConvertTo-Json @{
        action = "export_images"
        success = $true
        slides_count = $slideCount
        output_dir = $outputDirFull.Replace('\', '/')
        images = $exportedImages | ForEach-Object { $_.Replace('\', '/') }
    } -Compress)
}
finally {
    if ($pres -ne $null) {
        $pres.Close()
    }
    if ($ppt -ne $null) {
        $ppt.Quit()
    }
}
