# Render one slide per layout of a sample PPTX with PowerPoint (COM), all filled with the same filler text.
# usage: powershell -File make-layout-slides.ps1 -Sample <sample.pptx> -Spec <template-spec.json> -Filler <filler.md> -Out <dir>
# Writes <Out>\original\<layoutName>.png (fit inside 1280x720, aspect kept; compare.mjs letterboxes it with white bars).
# Exit codes: 0 ok, 2 input error, 4 PowerPoint COM not available (comparison is skipped, decision D4).
param(
  [Parameter(Mandatory = $true)][string]$Sample,
  [Parameter(Mandatory = $true)][string]$Spec,
  [Parameter(Mandatory = $true)][string]$Filler,
  [Parameter(Mandatory = $true)][string]$Out
)
$ErrorActionPreference = 'Stop'

foreach ($p in @($Sample, $Spec, $Filler)) {
  if (-not (Test-Path -LiteralPath $p)) { Write-Error "input not found: $p"; exit 2 }
}

# Filler: optional "# Title" line, then body lines. Comment lines are dropped; text after a "---" line is ignored.
$title = ''
$body = New-Object System.Collections.Generic.List[string]
foreach ($line in (Get-Content -LiteralPath $Filler -Encoding UTF8)) {
  $t = $line.Trim()
  if ($t -match '^---+$') { if ($title -or $body.Count) { break } else { continue } }
  if ($t -eq '' -or $t -match '^<!--.*-->$') { continue }
  if (-not $title -and $t -match '^#+\s*(.+)$') { $title = $Matches[1]; continue }
  $body.Add(($t -replace '^[-*+]\s+', '' -replace '^\d+[.)]\s+', ''))
}
if (-not $title -and $body.Count) { $title = $body[0]; $body.RemoveAt(0) }
$bodyText = ($body -join "`r")

$specObj = Get-Content -LiteralPath $Spec -Raw -Encoding UTF8 | ConvertFrom-Json
$names = @($specObj.layouts | ForEach-Object { $_.name })
if ($names.Count -eq 0) { Write-Error 'spec has no layouts'; exit 2 }

$outDir = Join-Path $Out 'original'
New-Item -ItemType Directory -Force -Path $outDir | Out-Null

$tmpDir = Join-Path ([IO.Path]::GetTempPath()) ('pptlayout-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Force -Path $tmpDir | Out-Null
$tmpCopy = Join-Path $tmpDir 'sample-copy.pptx'
Copy-Item -LiteralPath $Sample -Destination $tmpCopy

$app = $null
$pres = $null
$ownApp = $false
try {
  try {
    $app = New-Object -ComObject PowerPoint.Application
  } catch {
    Write-Output "PowerPoint COM not available: $($_.Exception.Message)"
    exit 4
  }
  # Open(FileName, ReadOnly, Untitled, WithWindow): the copy, no window.
  $pres = $app.Presentations.Open($tmpCopy, 0, 0, 0)
  $ownApp = ($app.Presentations.Count -le 1)

  $wPt = [double]$pres.PageSetup.SlideWidth
  $hPt = [double]$pres.PageSetup.SlideHeight
  $scale = [Math]::Min(1280.0 / $wPt, 720.0 / $hPt)
  $pxW = [int][Math]::Round($wPt * $scale)
  $pxH = [int][Math]::Round($hPt * $scale)

  # All custom layouts of all designs, by name (first match wins).
  $layouts = @{}
  foreach ($design in $pres.Designs) {
    foreach ($cl in $design.SlideMaster.CustomLayouts) {
      if (-not $layouts.ContainsKey($cl.Name)) { $layouts[$cl.Name] = $cl }
    }
  }

  $missing = 0
  foreach ($name in $names) {
    if (-not $layouts.ContainsKey($name)) {
      Write-Output "MISSING layout: $name"
      $missing++
      continue
    }
    $slide = $pres.Slides.AddSlide($pres.Slides.Count + 1, $layouts[$name])
    foreach ($shape in $slide.Shapes) {
      if ($shape.Type -ne 14) { continue }            # msoPlaceholder
      if (-not $shape.HasTextFrame) { continue }
      switch ([int]$shape.PlaceholderFormat.Type) {
        { $_ -in 1, 3 } { $shape.TextFrame.TextRange.Text = $title; break }              # title, centre title
        { $_ -in 2, 4, 7, 5, 6, 8 } { $shape.TextFrame.TextRange.Text = $bodyText; break } # body, subtitle, object, ...
        default { }                                                                       # date, footer, slide number, picture
      }
    }
    $safe = [regex]::Replace($name, '[\\/:*?"<>|]', '_')
    $png = Join-Path $outDir ($safe + '.png')
    $slide.Export($png, 'PNG', $pxW, $pxH)
    Write-Output "OK $name -> $png"
  }
  if ($missing -gt 0) { Write-Output "$missing layout(s) not found in the sample" }
} finally {
  if ($pres) { try { $pres.Close() } catch { } }
  if ($app -and $ownApp) { try { if ($app.Presentations.Count -eq 0) { $app.Quit() } } catch { } }
  Remove-Item -LiteralPath $tmpDir -Recurse -Force -ErrorAction SilentlyContinue
}
exit 0
