# RouteGraph PowerShell COM 规范

## 1. 脚本结构

每个 RouteGraph 脚本应尽量使用固定结构：

```powershell
# 1. Metadata
# 2. Helper functions
# 3. Create PowerPoint
# 4. Colors and global settings
# 5. Layout constants
# 6. Draw panels and shapes
# 7. Draw arrows
# 8. Send backgrounds to back
# 9. Save file
```

## 2. 基础 RGB 函数

PowerPoint COM 使用 BGR-like 整数表示颜色。统一使用：

```powershell
function RGB($r, $g, $b) {
    return ($r + ($g * 256) + ($b * 65536))
}
```

## 3. Helper functions

RouteGraph 生成脚本应优先使用 helper functions，避免到处散写 COM 命令。

推荐核心函数：

```powershell
Add-TextBox
Add-Box
Add-LineArrow
Add-ElbowArrow
Add-PolylineArrow
```

PowerShell 7 的 COM 绑定对数值类型更严格。`Line.Weight`、`BuildFreeform` 和 `AddNodes` 的浮点参数应显式转换为 `[single]`，避免 `System.Double` 到 `System.Int32` 的绑定错误。

可选函数：

```powershell
Add-RoundedBox
Add-ElbowArrow
Add-PanelLabel
Add-PanelTitle
Add-BackgroundBlock
```

## 4. 文本框函数规范

```powershell
function Add-TextBox($slide, $text, $x, $y, $w, $h, $fontSize = 14, $bold = $false, $color = $null) {
    $box = $slide.Shapes.AddTextbox(1, $x, $y, $w, $h)

    $box.TextFrame.TextRange.Text = $text
    $box.TextFrame.TextRange.Font.Name = "Arial"
    $box.TextFrame.TextRange.Font.Size = $fontSize
    $box.TextFrame.TextRange.Font.Bold = $bold

    if ($color -ne $null) {
        $box.TextFrame.TextRange.Font.Color.RGB = $color
    }

    $box.TextFrame.MarginLeft = 0
    $box.TextFrame.MarginRight = 0
    $box.TextFrame.MarginTop = 0
    $box.TextFrame.MarginBottom = 0

    # 禁止自动乱换行
    $box.TextFrame.WordWrap = 0

    return $box
}
```

## 5. 方框函数规范

```powershell
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
    }
    else {
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
```

## 6. 箭头函数规范

### 连续箭头原则

每个逻辑箭头应优先对应一个 PowerPoint shape。不要把同一条箭头路由拆成多个独立短线对象，否则虚线会在每段重新开始，视觉上像被短线硬拼出来。

优先级：

1. 直线路由：`Add-LineArrow`
2. 简单折线路由：`Add-ElbowArrow`
3. 多折线路由：`Add-PolylineArrow`
4. 只有真实分叉或刻意遮挡断点，才使用多个对象。

### 无箭头线段

```powershell
function Add-LineSegment($slide, $x1, $y1, $x2, $y2, $color, $dash = $false, $weight = 1.2) {
    $line = $slide.Shapes.AddConnector(1, $x1, $y1, $x2, $y2)
    $line.Line.ForeColor.RGB = $color
    $line.Line.Weight = [single]$weight
    $line.Line.BeginArrowheadStyle = 1
    $line.Line.EndArrowheadStyle = 1
    if ($dash) { $line.Line.DashStyle = 4 }
    return $line
}
```

`Add-LineSegment` 只用于真实分叉的主干、辅助参考线、或确实没有箭头头部的短线。不要用它拼接一条完整虚线箭头。

### 带箭头线段

```powershell
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
```

### 原生折线箭头

```powershell
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
```

### 可控连续多折线箭头

```powershell
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
```

## 7. 箭头路由规范

复杂连接优先使用一条连续多折线：

```powershell
Add-PolylineArrow $slide @(
    @(472, 134),
    @(528, 134),
    @(528, 205),
    @(580, 205)
) $purple $true $arrowPurpleWeight
```

避免默认一根斜线穿过其他框。

## 8. 图层顺序规范

PowerPoint 图层顺序非常重要。

常见顺序：

1. 背景块；
2. 普通框；
3. 箭头；
4. 需要覆盖箭头的模型框；
5. panel 标签和标题。

如果要让模型框遮住路线中段，应：

```text
先画线，再画模型框。
```

## 9. 字体规范

默认：

```powershell
$fontName = "Arial"
$fontPanelLabel = 13
$fontPanelTitle = 13
$fontBox = 9
$fontModel = 12
$fontBottomLabel = 13
```

避免使用小数字号，如 `8.5`、`11.5`，不同 PowerPoint 环境可能表现不稳定。

## 10. 坐标规范

函数统一使用：

```text
x, y, w, h
```

复杂模板中建议使用语义变量：

```powershell
$panelA_x = 25
$panelA_y = 86
$modelA_x = 180
$modelA_y = 124
$busA_x = 220
```

减少魔法数字，方便二次修改。
