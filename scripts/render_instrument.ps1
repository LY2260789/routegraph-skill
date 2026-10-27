param(
    [Parameter(Mandatory=$true)][string]$PlanPath,
    [Parameter(Mandatory=$true)][string]$OutputPath
)
$ErrorActionPreference = 'Stop'
$plan = Get-Content -LiteralPath $PlanPath -Raw -Encoding UTF8 | ConvertFrom-Json
if ($plan.schema_version -ne 'routegraph.render-plan.v1') { throw 'Unsupported render plan' }
$OutputPath = [IO.Path]::GetFullPath($OutputPath)
$preview = [IO.Path]::ChangeExtension($OutputPath, '.png')
$receipt = [IO.Path]::ChangeExtension($OutputPath, '.render.json')
foreach($p in @($OutputPath,$preview,$receipt)) { if(Test-Path -LiteralPath $p) { throw "Refusing overwrite: $p" } }
foreach($g in $plan.groups) {
    if($g.native) {
        if((Get-FileHash -LiteralPath $g.native.path -Algorithm SHA256).Hash.ToLower() -ne $g.native.sha256) { throw 'Native reference changed after compilation' }
    }
}
$null = New-Item -ItemType Directory -Force -Path ([IO.Path]::GetDirectoryName($OutputPath))
function RG-Color([string]$hex) {
    return [int]([Convert]::ToInt32($hex.Substring(0,2),16) + 256*[Convert]::ToInt32($hex.Substring(2,2),16) + 65536*[Convert]::ToInt32($hex.Substring(4,2),16))
}
function RG-Fill($shape,$item) {
    $shape.Line.Visible=0
    $shape.Fill.ForeColor.RGB=RG-Color $item.fill
    if($item.gradient) {
        $shape.Fill.TwoColorGradient(1,1)
        $shape.Fill.ForeColor.RGB=RG-Color $item.gradient[0]
        $shape.Fill.BackColor.RGB=RG-Color $item.gradient[1]
    }
}
function RG-Text($slide,$name,$text,$box,$size,$bold=$false) {
    $s=$slide.Shapes.AddTextbox(1,[single]$box[0],[single]$box[1],[single]$box[2],[single]$box[3])
    $s.Name=$name
    $tf=$s.TextFrame
    $tf.AutoSize=0
    $tf.MarginLeft=0; $tf.MarginRight=0; $tf.MarginTop=0; $tf.MarginBottom=0; $tf.WordWrap=0
    $tf.TextRange.Text=$text; $tf.TextRange.Font.Name=$plan.font
    $tf.TextRange.Font.Size=[single]$size; $tf.TextRange.Font.Color.RGB=0; $tf.TextRange.Font.Bold=[int]$bold * -1
    $s.Left=[single]$box[0]; $s.Top=[single]$box[1]; $s.Width=[single]$box[2]; $s.Height=[single]$box[3]
    return $s
}
function RG-Group($slide,$names,$id) {
    # Even single imported objects receive a semantic name. Multi-shape parts are groups.
    if($names.Count -eq 1) { $g=$slide.Shapes.Item([string]$names[0]) }
    else { $g=$slide.Shapes.Range([object[]]$names).Group() }
    $g.Name='component/'+$id
    return $g
}
$app=$null; $pres=$null; $sources=@{}; $manifest=@(); $textChecks=@()
try {
    $app=New-Object -ComObject PowerPoint.Application
    $pres=$app.Presentations.Add(0)
    $pres.PageSetup.SlideWidth=[single]$plan.canvas[0]
    $pres.PageSetup.SlideHeight=[single]$plan.canvas[1]
    $slide=$pres.Slides.Add(1,12)
    $slide.Background.Fill.ForeColor.RGB=16777215
    foreach($item in $plan.groups) {
        $names=@()
        if($item.native) {
            $path=[string]$item.native.path
            if(-not $sources.ContainsKey($path)) { $sources[$path]=$app.Presentations.Open($path,-1,0,0) }
            $source=$sources[$path].Slides.Item([int]$item.native.slide)
            # Group in the read-only in-memory source to preserve relative coordinates and order.
            $range=$source.Shapes.Range([object[]]$item.native.shape_names)
            if($range.Count -gt 1) { $sourceGroup=$range.Group() } else { $sourceGroup=$range.Item(1) }
            $left=$sourceGroup.Left; $top=$sourceGroup.Top
            $sourceGroup.Copy()
            $pasted=$slide.Shapes.Paste()
            $group=$pasted.Item(1)
            $group.Left=[single]($left+$item.origin[0]-$item.native.source_bounds[0])
            $group.Top=[single]($top+$item.origin[1]-$item.native.source_bounds[1])
            $group.Name='component/'+$item.id
        } else {
            foreach($part in $item.shapes) {
                if($part.kind -eq 'ellipse' -or $part.kind -eq 'rectangle') {
                    $b=$part.box
                    $shapeType=if($part.kind -eq 'ellipse'){9}else{1}
                    $s=$slide.Shapes.AddShape($shapeType,[single]$b[0],[single]$b[1],[single]$b[2],[single]$b[3])
                } elseif($part.kind -eq 'polygon') {
                    $pts=$part.points
                    $builder=$slide.Shapes.BuildFreeform(0,[single]$pts[0][0],[single]$pts[0][1])
                    for($i=1;$i -lt $pts.Count;$i++) { $builder.AddNodes(0,0,[single]$pts[$i][0],[single]$pts[$i][1]) }
                    $builder.AddNodes(0,0,[single]$pts[0][0],[single]$pts[0][1])
                    $s=$builder.ConvertToShape()
                } else { throw "Unsupported primitive $($part.kind)" }
                $s.Name=$item.id+'/'+$part.name
                RG-Fill $s $part
                $names += $s.Name
            }
            $group=RG-Group $slide $names $item.id
            $b=$item.local_bounds; $cx=$b[0]+$b[2]/2; $cy=$b[1]+$b[3]/2
            $rad=[double]$item.rotation*[Math]::PI/180
            $group.Rotation=[single](($item.rotation+360)%360)
            $group.Left=[single]($item.origin[0]+$cx*[Math]::Cos($rad)-$cy*[Math]::Sin($rad)-$group.Width/2)
            $group.Top=[single]($item.origin[1]+$cx*[Math]::Sin($rad)+$cy*[Math]::Cos($rad)-$group.Height/2)
        }
        $group.AlternativeText=($item.source | ConvertTo-Json -Compress -Depth 10)
        $angle=[double]$group.Rotation*[Math]::PI/180
        $bw=[Math]::Abs($group.Width*[Math]::Cos($angle))+[Math]::Abs($group.Height*[Math]::Sin($angle))
        $bh=[Math]::Abs($group.Width*[Math]::Sin($angle))+[Math]::Abs($group.Height*[Math]::Cos($angle))
        $bounds=@(($group.Left+$group.Width/2-$bw/2),($group.Top+$group.Height/2-$bh/2),$bw,$bh)
        $manifest += @{id=$item.id; bounds=$bounds; z_order=$group.ZOrderPosition; shape_type=[int]$group.Type}
    }
    foreach($label in $plan.labels) {
        $s=RG-Text $slide ('label/'+$label.id) $label.text $label.box $label.font_size
        $textChecks += @{id=$label.id; width=$s.Width; height=$s.Height; text_width=$s.TextFrame.TextRange.BoundWidth; text_height=$s.TextFrame.TextRange.BoundHeight; overflow=($s.TextFrame.TextRange.BoundWidth -gt $s.Width+1 -or $s.TextFrame.TextRange.BoundHeight -gt $s.Height+1)}
        $a=$slide.Shapes.AddLine([single]$label.start[0],[single]$label.start[1],[single]$label.target[0],[single]$label.target[1])
        $a.Name='leader/'+$label.id; $a.Line.ForeColor.RGB=0; $a.Line.Weight=[single]1.3; $a.Line.EndArrowheadStyle=2
        $index=0
        foreach($target in $label.additional_targets) {
            $index++
            $a=$slide.Shapes.AddLine([single]$label.start[0],[single]$label.start[1],[single]$target[0],[single]$target[1])
            $a.Name='leader/'+$label.id+'/'+$index; $a.Line.ForeColor.RGB=0; $a.Line.Weight=[single]1.3; $a.Line.EndArrowheadStyle=2
        }
    }
    foreach($a in $plan.annotations) {
        $s=RG-Text $slide ('annotation/'+$a.id) $a.text $a.box $a.font_size $a.bold
        $textChecks += @{id=('annotation/'+$a.id); width=$s.Width; height=$s.Height; text_width=$s.TextFrame.TextRange.BoundWidth; text_height=$s.TextFrame.TextRange.BoundHeight; overflow=($s.TextFrame.TextRange.BoundWidth -gt $s.Width+1 -or $s.TextFrame.TextRange.BoundHeight -gt $s.Height+1)}
    }
    $pres.SaveAs($OutputPath,24)
    $slide.Export($preview,'PNG',[int]($plan.canvas[0]*2),[int]($plan.canvas[1]*2))
    @{schema_version='routegraph.render-receipt.v1'; pptx=$OutputPath; sha256=(Get-FileHash -LiteralPath $OutputPath).Hash.ToLower(); plan_sha256=(Get-FileHash -LiteralPath $PlanPath).Hash.ToLower(); groups=$manifest; text_checks=$textChecks; top_level_objects=$slide.Shapes.Count} | ConvertTo-Json -Depth 20 | Set-Content -LiteralPath $receipt -Encoding UTF8
    Write-Output "PPTX=$OutputPath`nPREVIEW=$preview`nRECEIPT=$receipt"
} finally {
    foreach($source in $sources.Values) { $source.Saved=-1; $source.Close() }
    if($pres) { $pres.Saved=-1; $pres.Close() }
    # PowerPoint is a shared COM server. Never Quit an application with user presentations.
    if($app) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($app) }
}
