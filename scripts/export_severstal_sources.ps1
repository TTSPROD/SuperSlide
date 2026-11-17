param([string]$Root)
$ErrorActionPreference='Stop'
$app=$null
$deck=$null
$existing=0
$oldAlerts=$null
try {
 $app=New-Object -ComObject PowerPoint.Application
 $existing=$app.Presentations.Count
 $oldAlerts=$app.DisplayAlerts
 $app.DisplayAlerts=1
 $reports=@()
 foreach($source in Get-ChildItem -LiteralPath (Join-Path $Root 'SeverstalBrand\Template') -Filter '*.pptx') {
  $dest=Join-Path $Root ('output\severstal-source-previews\'+$source.BaseName)
  New-Item -ItemType Directory -Path $dest -Force | Out-Null
  $deck=$app.Presentations.Open($source.FullName,-1,0,0)
  foreach($slide in $deck.Slides) { $slide.Export((Join-Path $dest ('slide-{0:D2}.png' -f $slide.SlideIndex)),'PNG',1600,900) }
  $reports+=@{file=$source.Name;slides=$deck.Slides.Count;version=$app.Version}
  $deck.Close()
  $deck=$null
  Write-Output ('Exported '+$source.Name)
 }
 $reports | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $Root 'output\severstal-source-previews\report.json') -Encoding utf8
} finally {
 if($null -ne $deck){$deck.Close()}
 if($null -ne $app){if($null -ne $oldAlerts){$app.DisplayAlerts=$oldAlerts};if($existing -eq 0 -and $app.Presentations.Count -eq 0){$app.Quit()};[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
}
