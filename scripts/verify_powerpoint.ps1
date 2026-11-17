param(
    [Parameter(Mandatory = $true)][string]$InputFile,
    [string]$OutputDirectory = (Join-Path (Split-Path -Parent $PSScriptRoot) 'output/powerpoint')
)

$ErrorActionPreference = 'Stop'
$inputPath = (Resolve-Path -LiteralPath $InputFile).Path
$directoryPath = [System.IO.Path]::GetFullPath($OutputDirectory)
New-Item -ItemType Directory -Path $directoryPath -Force | Out-Null
$pdfPath = Join-Path $directoryPath 'presentation.pdf'
$reportPath = Join-Path $directoryPath 'powerpoint-report.json'
$presentation = $null
$application = $null
$existingCount = 0
$previousAlerts = $null
try {
    $application = New-Object -ComObject PowerPoint.Application
    $existingCount = $application.Presentations.Count
    $previousAlerts = $application.DisplayAlerts
    $application.DisplayAlerts = 1
    # No application window or user presentation is opened/changed by this verification.
    $presentation = $application.Presentations.Open($inputPath, -1, 0, 0)
    $slideResults = @()
    foreach ($slide in $presentation.Slides) {
        $tables = 0
        $charts = 0
        $texts = 0
        $pictures = 0
        $overflow = @()
        foreach ($shape in $slide.Shapes) {
            if ($shape.Type -eq 13 -or ($shape.Type -eq 14 -and $shape.PlaceholderFormat.ContainedType -eq 13)) { $pictures++ }
            if ($shape.HasTable -eq -1) { $tables++ }
            if ($shape.HasChart -eq -1) { $charts++ }
            if ($shape.HasTextFrame -eq -1 -and $shape.TextFrame.HasText -eq -1) {
                $texts++
                # Native Office metrics: text overflows even if the package itself is valid.
                $frame = $shape.TextFrame2
                $heightAvailable = $shape.Height - $frame.MarginTop - $frame.MarginBottom
                if ($frame.TextRange.BoundHeight -gt ($heightAvailable + 1.5)) {
                    $overflow += @{ shape = $shape.Name; boundHeight = $frame.TextRange.BoundHeight; availableHeight = $heightAvailable }
                }
            }
        }
        $imagePath = Join-Path $directoryPath ('slide-{0:D2}.png' -f $slide.SlideIndex)
        $slide.Export($imagePath, 'PNG', 1600, 900)
        $slideResults += @{ slide = $slide.SlideIndex; layout = $slide.CustomLayout.Name; tables = $tables; charts = $charts; pictures = $pictures; textShapes = $texts; overflow = $overflow }
    }
    $presentation.SaveAs($pdfPath, 32)
    $result = @{
        input = $inputPath
        application = 'Microsoft PowerPoint'
        version = $application.Version
        slideCount = $presentation.Slides.Count
        opened = $true
        pdf = $pdfPath
        slides = $slideResults
    }
    $result | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $reportPath -Encoding utf8
    Write-Output ($result | ConvertTo-Json -Depth 8)
}
finally {
    if ($null -ne $presentation) { $presentation.Close() }
    if ($null -ne $application) {
        if ($null -ne $previousAlerts) { $application.DisplayAlerts = $previousAlerts }
        if ($existingCount -eq 0 -and $application.Presentations.Count -eq 0) { $application.Quit() }
        [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($application)
    }
}
