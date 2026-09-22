# ============================================================
# RouteGraph Example: PM2.5 Counterfactual Framework
# PowerShell + PowerPoint COM
# ============================================================

function RGB($r, $g, $b) {
    return ($r + ($g * 256) + ($b * 65536))
}

function Add-TextBox($slide, $text, $x, $y, $w, $h, $fontSize = 14, $bold = $false, $color = $null) {
    $box = $slide.Shapes.AddTextbox(1, $x, $y, $w, $h)
    $box.TextFrame.TextRange.Text = $text
    $box.TextFrame.TextRange.Font.Name = "Arial"
    $box.TextFrame.TextRange.Font.Size = $fontSize
    $box.TextFrame.TextRange.Font.Bold = $bold
    if ($color -ne $null) { $box.TextFrame.TextRange.Font.Color.RGB = $color }
    $box.TextFrame.MarginLeft = 0
    $box.TextFrame.MarginRight = 0
    $box.TextFrame.MarginTop = 0
    $box.TextFrame.MarginBottom = 0
    $box.TextFrame.WordWrap = 0
    return $box
}

function Add-Box($slide, $text, $x, $y, $w, $h, $fillColor, $lineColor, $fontSize = 10, $bold = $true, $fontColor = $null) {
    $shape = $slide.Shapes.AddShape(1, $x, $y, $w, $h)
    $shape.Fill.ForeColor.RGB = $fillColor
    $shape.Line.ForeColor.RGB = $lineColor
    $shape.Line.Weight = [single]1.15
    $shape.TextFrame.TextRange.Text = $text
    $shape.TextFrame.TextRange.Font.Name = "Arial"
    $shape.TextFrame.TextRange.Font.Size = $fontSize
    $shape.TextFrame.TextRange.Font.Bold = $bold
    $shape.TextFrame.TextRange.ParagraphFormat.Alignment = 2
    if ($fontColor -ne $null) {
        $shape.TextFrame.TextRange.Font.Color.RGB = $fontColor
    } else {
        $shape.TextFrame.TextRange.Font.Color.RGB = RGB 0 0 0
    }
    $shape.TextFrame.VerticalAnchor = 3
    $shape.TextFrame.MarginLeft = 2
    $shape.TextFrame.MarginRight = 2
    $shape.TextFrame.MarginTop = 1
    $shape.TextFrame.MarginBottom = 1
    $shape.TextFrame.WordWrap = 0
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

$ppt = New-Object -ComObject PowerPoint.Application
$ppt.Visible = $true

$presentation = $ppt.Presentations.Add()
$presentation.PageSetup.SlideWidth = 1000
$presentation.PageSetup.SlideHeight = 240
$slide = $presentation.Slides.Add(1, 12)

$blue = RGB 30 100 210
$lightBlue = RGB 190 220 250
$purple = RGB 105 45 175
$lightPurple = RGB 222 214 242
$gray = RGB 225 230 235
$lightGray = RGB 235 238 242
$darkGray = RGB 65 65 65
$orange = RGB 235 145 20
$lightOrange = RGB 255 238 200
$black = RGB 0 0 0
$white = RGB 255 255 255
$Delta = [char]0x0394

$arrowBlueWeight = 1.15
$arrowPurpleWeight = 1.45

# Panel A
Add-TextBox $slide "A." 10 8 30 20 13 $true $black
Add-TextBox $slide "Training a model with inputs`nfrom non-intervention period" 18 30 230 42 13 $false $black

$bgA = $slide.Shapes.AddShape(1, 25, 86, 140, 125)
$bgA.Fill.ForeColor.RGB = $lightGray
$bgA.Line.Visible = 0

Add-Box $slide "Emission proxies"      35 96 115 23 $gray      $darkGray 9  $true $black
Add-Box $slide "Meteorological var"    35 128 115 23 $gray      $darkGray 9  $true $black
Add-Box $slide "Chemical indicator"    35 160 115 23 $gray      $darkGray 9  $true $black
Add-Box $slide "PM2.5"                 35 192 115 25 $lightBlue $blue     10 $true $black

$busAX = 220
$topAY = 107
$botAY = 205
Add-LineSegment $slide $busAX $topAY $busAX $botAY $blue $false $arrowBlueWeight
Add-LineArrow   $slide $busAX $topAY 150 $topAY $blue $false $arrowBlueWeight
Add-LineArrow   $slide $busAX $botAY 150 $botAY $blue $false $arrowBlueWeight

Add-TextBox $slide "~" 164 137 18 22 16 $false $black
Add-Box $slide "XGBoost" 180 124 78 36 $white $blue 12 $true $black

# Panel B
Add-TextBox $slide "B." 280 8 30 20 13 $true $black
Add-TextBox $slide "Predicting PM2.5 for intervention`nperiod using the trained model" 288 30 250 42 13 $false $black

Add-TextBox $slide "BAU simulations" 350 76 130 22 13 $true $purple
Add-LineArrow $slide 400 90 400 112 $black $false 1.1

$bgB = $slide.Shapes.AddShape(1, 310, 102, 190, 90)
$bgB.Fill.ForeColor.RGB = $lightPurple
$bgB.Line.Visible = 0

Add-Box $slide "Counterfactual conc." 340 118 132 32 $lightPurple $darkGray 10 $true $purple
Add-Box $slide "Observed conc."       340 164 132 32 $gray        $darkGray 10 $true $black

Add-LineArrow $slide 258 142 340 134 $purple $true $arrowPurpleWeight
Add-TextBox $slide "$Delta PM2.5 air quality" 340 213 190 22 13 $true $black

# Panel C
Add-TextBox $slide "C." 560 8 30 20 13 $true $black
Add-TextBox $slide "Re-training PM2.5 model`nfor all measurements" 568 30 220 42 13 $false $black

$bgC = $slide.Shapes.AddShape(1, 570, 86, 140, 125)
$bgC.Fill.ForeColor.RGB = $lightGray
$bgC.Line.Visible = 0

Add-Box $slide "Emission proxies"      580 96 115 23 $gray      $darkGray 9 $true $black
Add-Box $slide "Meteorological var"    580 128 115 23 $gray      $darkGray 9 $true $black
Add-Box $slide "Chemical indicator"    580 160 115 23 $gray      $darkGray 9 $true $black
Add-Box $slide "Counterfactual PM2.5"  580 192 120 25 $lightBlue $blue     9 $true $purple

$busCX = 760
$topCY = 107
$botCY = 205
Add-LineSegment $slide $busCX $topCY $busCX $botCY $blue $false $arrowBlueWeight
Add-LineArrow   $slide $busCX $topCY 695 $topCY $blue $false $arrowBlueWeight
Add-LineArrow   $slide $busCX $botCY 700 $botCY $blue $false $arrowBlueWeight

Add-TextBox $slide "~" 708 137 18 22 16 $false $black
Add-Box $slide "XGBoost" 724 124 78 36 $white $blue 12 $true $black

Add-PolylineArrow $slide @(
    @(472, 134),
    @(528, 134),
    @(528, 205),
    @(580, 205)
) $purple $true $arrowPurpleWeight

# Panel D
Add-TextBox $slide "D." 815 8 30 20 13 $true $black
Add-TextBox $slide "Predicting PM2.5 under fixed`nmeteorological states" 823 30 220 42 13 $false $black

$bgD = $slide.Shapes.AddShape(1, 830, 100, 170, 95)
$bgD.Fill.ForeColor.RGB = $lightOrange
$bgD.Line.Visible = 0

Add-Box $slide "Counterfactual`nemission strength" 848 114 145 38 $white $orange 10 $true $purple
Add-Box $slide "Emission strength"     848 164 145 34 $white $orange 10 $true $black

$forkX = 830
$topY  = 133
$midY  = 145
$botY  = 181

# True one-to-two branch: separate shapes are intentional here.
Add-LineSegment $slide 802 $midY $forkX $midY $purple $true $arrowPurpleWeight
Add-LineSegment $slide $forkX $topY $forkX $botY $purple $true $arrowPurpleWeight
Add-LineArrow   $slide $forkX $topY 848 $topY $purple $true $arrowPurpleWeight
Add-LineArrow   $slide $forkX $botY 848 $botY $purple $true $arrowPurpleWeight

Add-TextBox $slide "$Delta Emission strength" 848 213 190 22 13 $true $black

$bgA.ZOrder(1)
$bgB.ZOrder(1)
$bgC.ZOrder(1)
$bgD.ZOrder(1)

$outputPath = "$env:USERPROFILE\Desktop\RouteGraph_PM25_counterfactual.pptx"
$presentation.SaveAs($outputPath)

Write-Host "Done! PPT saved to: $outputPath"
