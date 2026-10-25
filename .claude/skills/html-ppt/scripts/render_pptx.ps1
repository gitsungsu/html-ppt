# Export every slide of a .pptx to PNG using installed PowerPoint (Windows), to verify the native output.
# usage: powershell -File render_pptx.ps1 <deck.pptx> <outDir>
param([Parameter(Mandatory)][string]$Pptx, [Parameter(Mandatory)][string]$OutDir)
$Pptx = (Resolve-Path $Pptx).Path
New-Item -ItemType Directory -Force $OutDir | Out-Null
$OutDir = (Resolve-Path $OutDir).Path
$app = New-Object -ComObject PowerPoint.Application
try {
  $pres = $app.Presentations.Open($Pptx, $true, $false, $false)
  # keep the slide's own aspect ratio (4:3 templates would otherwise be stretched to 16:9)
  $height = [int][math]::Round(1280 * $pres.PageSetup.SlideHeight / $pres.PageSetup.SlideWidth)
  $pres.Export($OutDir, 'PNG', 1280, $height)
  $pres.Close()
} finally { $app.Quit() }
Get-ChildItem $OutDir -Filter *.PNG | ForEach-Object { $_.FullName }
