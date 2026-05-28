# ============================================================
# RouteGraph Four-Panel Technical Route Template
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
    $shape.Line.Weight = 1.15
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
    $line.Line.Weight = $weight
    $line.Line.BeginArrowheadStyle = 1
    $line.Line.EndArrowheadStyle = 1
    if ($dash) { $line.Line.DashStyle = 4 }
    return $line
}

function Add-LineArrow($slide, $x1, $y1, $x2, $y2, $color, $dash = $false, $weight = 1.2) {
    $line = $slide.Shapes.AddConnector(1, $x1, $y1, $x2, $y2)
    $line.Line.ForeColor.RGB = $color
    $line.Line.Weight = $weight
    $line.Line.BeginArrowheadStyle = 1
    $line.Line.EndArrowheadStyle = 3
    $line.Line.EndArrowheadLength = 2
    $line.Line.EndArrowheadWidth = 2
    if ($dash) { $line.Line.DashStyle = 4 }
    return $line
}

$ppt = New-Object -ComObject PowerPoint.Application
$ppt.Visible = $true

$presentation = $ppt.Presentations.Add()
$presentation.PageSetup.SlideWidth = 1000
$presentation.PageSetup.SlideHeight = 260
$slide = $presentation.Slides.Add(1, 12)

$blue = RGB 30 100 210
$lightBlue = RGB 190 220 250
$purple = RGB 105 45 175
$lightPurple = RGB 222 214 242
$gray = RGB 225 230 235
$lightGray = RGB 235 238 242
$darkGray = RGB 65 65 65
$black = RGB 0 0 0
$white = RGB 255 255 255
$Delta = [char]0x0394

$arrowBlueWeight = 1.15
$arrowPurpleWeight = 1.45

# TODO: Define panel coordinates and elements here.

$outputPath = "$env:USERPROFILE\Desktop\RouteGraph_output.pptx"
$presentation.SaveAs($outputPath)

Write-Host "Done! PPT saved to: $outputPath"
