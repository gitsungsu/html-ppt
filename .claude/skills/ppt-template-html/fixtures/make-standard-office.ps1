# Creates standard-office.pptx via PowerPoint COM (Windows PowerShell 5.1).
# Default Office theme, a logo PNG on the slide master, and two custom layouts:
#   TITLE_CONTENT - inherits the master decoration (logo)
#   COVER         - DisplayMasterShapes = false, FollowMasterBackground = false, solid background colour
# Usage: powershell -File make-standard-office.ps1
$ErrorActionPreference = 'Stop'

$fixtures = (Resolve-Path $PSScriptRoot).Path
$logo = Join-Path $fixtures 'images\sample.png'
$out = Join-Path $fixtures 'standard-office.pptx'

if (-not (Test-Path -LiteralPath $logo)) { throw "logo not found: $logo" }
if (Test-Path -LiteralPath $out) { Remove-Item -LiteralPath $out -Force }

$msoTrue = -1
$msoFalse = 0
$ppSaveAsOpenXMLPresentation = 24

$app = $null
$pres = $null
try {
    $app = New-Object -ComObject PowerPoint.Application
    $pres = $app.Presentations.Add($msoFalse)   # no window

    # 16:9, 13.333in x 7.5in (12192000 x 6858000 EMU)
    $pres.PageSetup.SlideWidth = 960
    $pres.PageSetup.SlideHeight = 540

    $master = $pres.SlideMaster

    # Logo picture on the slide master (top right)
    $pic = $master.Shapes.AddPicture($logo, $msoFalse, $msoTrue, 820, 10, 120, 60)
    $pic.Name = 'Logo'

    # TITLE_CONTENT: copy of the built-in "Title and Content" layout (index 2), keeps master shapes
    $content = $master.CustomLayouts.Item(2).Duplicate()
    $content.Name = 'TITLE_CONTENT'
    $content.DisplayMasterShapes = $msoTrue
    $content.FollowMasterBackground = $msoTrue

    # COVER: copy of the built-in "Title Slide" layout (index 1), no master shapes, own solid background
    $cover = $master.CustomLayouts.Item(1).Duplicate()
    $cover.Name = 'COVER'
    $cover.DisplayMasterShapes = $msoFalse
    $cover.FollowMasterBackground = $msoFalse
    $cover.Background.Fill.Solid()
    # RGB(31, 78, 121) = 1F4E79; COM RGB is R + G*256 + B*65536
    $cover.Background.Fill.ForeColor.RGB = 31 + (78 * 256) + (121 * 65536)
    # White text for the title and subtitle placeholders on the dark COVER background (RGB 255,255,255)
    $white = 255 + (255 * 256) + (255 * 65536)
    foreach ($shape in $cover.Shapes) {
        if ($shape.Type -eq 14 -and ($shape.PlaceholderFormat.Type -eq 3 -or $shape.PlaceholderFormat.Type -eq 4)) {
            # ppPlaceholderCenterTitle = 3, ppPlaceholderSubtitle = 4
            $shape.TextFrame.TextRange.Font.Color.RGB = $white
        }
    }

    # Remove every other layout so only the two custom layouts remain
    for ($i = $master.CustomLayouts.Count; $i -ge 1; $i--) {
        $layout = $master.CustomLayouts.Item($i)
        if ($layout.Name -ne 'TITLE_CONTENT' -and $layout.Name -ne 'COVER') {
            $layout.Delete()
        }
    }

    $pres.SaveAs($out, $ppSaveAsOpenXMLPresentation)
    Write-Output "wrote $out"
}
finally {
    if ($pres -ne $null) { $pres.Close() }
    if ($app -ne $null) { $app.Quit() }
}
