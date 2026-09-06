param(
    [Parameter(Mandatory = $true)]
    [string]$InputPath,

    [Parameter(Mandatory = $true)]
    [string]$OutputPath
)

$ErrorActionPreference = 'Stop'

$ppt = $null
$pres = $null

try {
    $resolvedInput = (Resolve-Path -LiteralPath $InputPath).Path
    $outputFull = [System.IO.Path]::GetFullPath($OutputPath)
    $outputDir = Split-Path -Parent $outputFull
    if (-not (Test-Path -LiteralPath $outputDir)) {
        New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
    }

    $ppt = New-Object -ComObject PowerPoint.Application
    # Presentations.Open(FileName, ReadOnly, Untitled, WithWindow)
    # -1 = msoTrue, 0 = msoFalse
    $pres = $ppt.Presentations.Open($resolvedInput, -1, 0, 0)
    # ppSaveAsPDF = 32
    $pres.SaveAs($outputFull, 32)
    Write-Output $outputFull
}
finally {
    if ($pres -ne $null) {
        $pres.Close()
    }
    if ($ppt -ne $null) {
        $ppt.Quit()
    }
}
