param(
    [string]$OutputPath = (Join-Path ([Environment]::GetFolderPath("Desktop")) "RouteGraph_continuous_arrow_regression.pptx")
)

# RouteGraph regression fixture for continuous arrow routing.
# Every logical non-branching arrow below is a single PowerPoint shape.

function RGB($r, $g, $b) {
    return ($r + ($g * 256) + ($b * 65536))
}

function Add-Label($slide, $text, $x, $y, $w, $h, $size = 15, $bold = $false) {
    $box = $slide.Shapes.AddTextbox(1, $x, $y, $w, $h)
    $box.TextFrame.TextRange.Text = $text
    $box.TextFrame.TextRange.Font.Name = "Arial"
    $box.TextFrame.TextRange.Font.Size = $size
    $box.TextFrame.TextRange.Font.Bold = $bold
    $box.TextFrame.MarginLeft = 0
    $box.TextFrame.MarginRight = 0
    $box.TextFrame.MarginTop = 0
    $box.TextFrame.MarginBottom = 0
    return $box
}

function Add-TestBox($slide, $text, $x, $y, $w, $h, $fillColor, $lineColor) {
    $shape = $slide.Shapes.AddShape(5, $x, $y, $w, $h)
    $shape.Fill.ForeColor.RGB = $fillColor
    $shape.Line.ForeColor.RGB = $lineColor
    $shape.Line.Weight = [single]1.15
    $shape.TextFrame.TextRange.Text = $text
    $shape.TextFrame.TextRange.Font.Name = "Arial"
    $shape.TextFrame.TextRange.Font.Size = 13
    $shape.TextFrame.TextRange.Font.Bold = $true
    $shape.TextFrame.TextRange.ParagraphFormat.Alignment = 2
    $shape.TextFrame.VerticalAnchor = 3
    return $shape
}

function Add-LineSegment($slide, $x1, $y1, $x2, $y2, $color, $dash = $false, $weight = 1.2) {
    $line = $slide.Shapes.AddConnector(1, $x1, $y1, $x2, $y2)
    $line.Line.ForeColor.RGB = $color
    $line.Line.Weight = [single]$weight
    $line.Line.BeginArrowheadStyle = 1
    $line.Line.EndArrowheadStyle = 1
    if ($dash) { $line.Line.DashStyle = 4 }
    return $line
}

function Add-LineArrow($slide, $x1, $y1, $x2, $y2, $color, $dash = $false, $weight = 1.2) {
    $line = $slide.Shapes.AddConnector(1, $x1, $y1, $x2, $y2)
    $line.Line.ForeColor.RGB = $color
    $line.Line.Weight = [single]$weight
    $line.Line.BeginArrowheadStyle = 1
    $line.Line.EndArrowheadStyle = 3
    $line.Line.EndArrowheadLength = 2
    $line.Line.EndArrowheadWidth = 2
    if ($dash) { $line.Line.DashStyle = 4 }
    return $line
}

function Add-ElbowArrow($slide, $x1, $y1, $x2, $y2, $color, $dash = $false, $weight = 1.2) {
    $line = $slide.Shapes.AddConnector(2, $x1, $y1, $x2, $y2)
    $line.Line.ForeColor.RGB = $color
    $line.Line.Weight = [single]$weight
    $line.Line.BeginArrowheadStyle = 1
    $line.Line.EndArrowheadStyle = 3
    $line.Line.EndArrowheadLength = 2
    $line.Line.EndArrowheadWidth = 2
    if ($dash) { $line.Line.DashStyle = 4 }
    return $line
}

function Add-PolylineArrow($slide, $points, $color, $dash = $false, $weight = 1.2) {
    if ($points.Count -lt 2) { throw "Add-PolylineArrow needs at least two points." }

    $builder = $slide.Shapes.BuildFreeform(1, [single]$points[0][0], [single]$points[0][1])
    for ($i = 1; $i -lt $points.Count; $i++) {
        $builder.AddNodes(1, 1, [single]$points[$i][0], [single]$points[$i][1])
    }

    $shape = $builder.ConvertToShape()
    $shape.Fill.Visible = 0
    $shape.Line.ForeColor.RGB = $color
    $shape.Line.Weight = [single]$weight
    $shape.Line.BeginArrowheadStyle = 1
    $shape.Line.EndArrowheadStyle = 3
    $shape.Line.EndArrowheadLength = 2
    $shape.Line.EndArrowheadWidth = 2
    if ($dash) { $shape.Line.DashStyle = 4 }
    return $shape
}

$ppt = $null
$presentation = $null
$slide = $null

try {
    $ppt = New-Object -ComObject PowerPoint.Application
    $ppt.Visible = $true
    $presentation = $ppt.Presentations.Add()
    $presentation.PageSetup.SlideWidth = 960
    $presentation.PageSetup.SlideHeight = 540
    $slide = $presentation.Slides.Add(1, 12)

    $ink = RGB 25 56 92
    $blue = RGB 47 111 191
    $purple = RGB 112 63 170
    $orange = RGB 225 133 42
    $lightBlue = RGB 229 240 252
    $lightPurple = RGB 239 232 249
    $lightOrange = RGB 252 238 220

    [void](Add-Label $slide "Continuous Arrow Routing Regression" 54 28 852 38 24 $true)

    [void](Add-Label $slide "A1  Straight solid" 55 94 190 24 14 $true)
    [void](Add-TestBox $slide "Source" 60 128 120 44 $lightBlue $blue)
    [void](Add-TestBox $slide "Target" 330 128 120 44 $lightBlue $blue)
    $a1 = Add-LineArrow $slide 180 150 330 150 $blue $false 1.8
    $a1.Name = "A1_Straight_Solid"

    [void](Add-Label $slide "A2  Native elbow dashed" 510 94 250 24 14 $true)
    [void](Add-TestBox $slide "Source" 520 128 120 44 $lightPurple $purple)
    [void](Add-TestBox $slide "Target" 770 190 120 44 $lightPurple $purple)
    $a2 = Add-ElbowArrow $slide 640 150 770 212 $purple $true 1.8
    $a2.Name = "A2_Elbow_Dashed"

    [void](Add-Label $slide "A3  Multi-bend dashed" 55 258 220 24 14 $true)
    [void](Add-TestBox $slide "Source" 60 300 120 44 $lightPurple $purple)
    [void](Add-TestBox $slide "Target" 370 390 120 44 $lightPurple $purple)
    $a3 = Add-PolylineArrow $slide @(
        @(180, 322),
        @(280, 322),
        @(280, 412),
        @(370, 412)
    ) $purple $true 1.8
    $a3.Name = "A3_Polyline_Dashed"

    [void](Add-Label $slide "A4  True branch" 550 258 180 24 14 $true)
    [void](Add-TestBox $slide "Source" 555 330 110 44 $lightOrange $orange)
    [void](Add-TestBox $slide "Target 1" 800 296 110 44 $lightOrange $orange)
    [void](Add-TestBox $slide "Target 2" 800 404 110 44 $lightOrange $orange)

    # A true one-to-two branch intentionally uses multiple shapes.
    $a4Trunk = Add-LineSegment $slide 665 352 720 352 $orange $false 1.6
    $a4Trunk.Name = "A4_Branch_Trunk"
    $a4Bus = Add-LineSegment $slide 720 318 720 426 $orange $false 1.6
    $a4Bus.Name = "A4_Branch_Bus"
    $a4Top = Add-LineArrow $slide 720 318 800 318 $orange $false 1.6
    $a4Top.Name = "A4_Branch_Target_1"
    $a4Bottom = Add-LineArrow $slide 720 426 800 426 $orange $false 1.6
    $a4Bottom.Name = "A4_Branch_Target_2"

    $presentation.SaveAs($OutputPath)
    $previewPath = [System.IO.Path]::ChangeExtension($OutputPath, ".png")
    $slide.Export($previewPath, "PNG", 960, 540)
    Write-Host "PPTX=$OutputPath"
    Write-Host "PREVIEW=$previewPath"
}
finally {
    if ($presentation -ne $null) { $presentation.Close() }
    if ($ppt -ne $null) { $ppt.Quit() }
    if ($slide -ne $null) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($slide) }
    if ($presentation -ne $null) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($presentation) }
    if ($ppt -ne $null) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt) }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
