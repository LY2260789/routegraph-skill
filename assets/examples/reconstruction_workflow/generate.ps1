param([Parameter(Mandatory=$true)][string]$OutputPath)
$ErrorActionPreference='Stop'
$OutputPath=[IO.Path]::GetFullPath($OutputPath)
$preview=[IO.Path]::ChangeExtension($OutputPath,'.png')
foreach($file in @($OutputPath,$preview)){if(Test-Path -LiteralPath $file){throw "Refusing overwrite: $file"}}
$null=New-Item -ItemType Directory -Force -Path ([IO.Path]::GetDirectoryName($OutputPath))
function RG-Color([string]$hex){return [int]([Convert]::ToInt32($hex.Substring(0,2),16)+256*[Convert]::ToInt32($hex.Substring(2,2),16)+65536*[Convert]::ToInt32($hex.Substring(4,2),16))}
function RG-Text($id,$value,$x,$y,$w,$h,$size,$color,$bold=$false){
 $s=$slide.Shapes.AddTextbox(1,[single]$x,[single]$y,[single]$w,[single]$h);$s.Name=$id
 $tf=$s.TextFrame;$tf.AutoSize=0;$tf.WordWrap=0
 $tf.MarginLeft=0;$tf.MarginRight=0;$tf.MarginTop=0;$tf.MarginBottom=0
 $tf.TextRange.Text=$value;$tf.TextRange.Font.Name='Arial';$tf.TextRange.Font.Size=[single]$size
 $tf.TextRange.Font.Bold=([int]$bold)*-1;$tf.TextRange.Font.Color.RGB=RG-Color $color
 return $s
}
function RG-Box($id,$x,$y,$w,$h,$fill){
 $s=$slide.Shapes.AddShape(1,[single]$x,[single]$y,[single]$w,[single]$h);$s.Name=$id
 $s.Fill.ForeColor.RGB=RG-Color $fill;$s.Line.Visible=0;return $s
}
function RG-Arrow($id,$points,$color,$dashed=$false){
 $b=$slide.Shapes.BuildFreeform(0,[single]$points[0][0],[single]$points[0][1])
 for($i=1;$i -lt $points.Count;$i++){$b.AddNodes(0,0,[single]$points[$i][0],[single]$points[$i][1])}
 $s=$b.ConvertToShape();$s.Name=$id;$s.Fill.Visible=0
 $s.Line.ForeColor.RGB=RG-Color $color;$s.Line.Weight=[single]2.5;$s.Line.EndArrowheadStyle=2
 if($dashed){$s.Line.DashStyle=4};return $s
}
$app=$null;$pres=$null
try{
 $app=New-Object -ComObject PowerPoint.Application;$pres=$app.Presentations.Add(0)
 $pres.PageSetup.SlideWidth=1200;$pres.PageSetup.SlideHeight=600
 $slide=$pres.Slides.Add(1,12);$slide.Background.Fill.ForeColor.RGB=RG-Color 'FFFFFF'
 $null=RG-Text 'title' 'From reference figure to editable PowerPoint' 36 30 1128 48 32 '172B4D' $true
 $null=RG-Text 'subtitle' 'RouteGraph  /  Agent-guided scientific figure reconstruction' 36 91 1128 30 18 '52627A'
 $stages=@(
  @{id='reference';x=36;num='01';title='Inspect reference';color='176B8A';light='EFF7FA';tag='AGENT';a='Read labels and arrows';b='Identify visible components';c='Record uncertain details'},
  @{id='structure';x=328;num='02';title='Record structure';color='4B5BB5';light='F1F2FA';tag='STRUCTURE';a='Instrument scene or topology';b='Geometry, labels and layers';c='Keep source evidence'},
  @{id='drawing';x=620;num='03';title='Draw native PPT';color='267466';light='EFF7F4';tag='POWERPOINT COM';a='Shapes, text and connectors';b='Materials and named groups';c='Editable component surfaces'},
  @{id='review';x=912;num='04';title='Render and review';color='9A6419';light='FBF6ED';tag='CHECK + AGENT REVIEW';a='Check structure and labels';b='Inspect overlap and arrows';c='Deliver PPTX, PNG and notes'}
 )
 foreach($p in $stages){
  $parts=@();$x=$p.x;$id=$p.id
  $parts+=RG-Box "$id/panel" $x 158 252 268 $p.light
  $parts+=RG-Text "$id/number" $p.num ($x+20) 176 40 28 19 $p.color $true
  $parts+=RG-Text "$id/tag" $p.tag ($x+20) 211 218 20 11 $p.color $true
  $parts+=RG-Text "$id/title" $p.title ($x+20) 246 221 34 21 '172B4D' $true
  $parts+=RG-Box "$id/rule" ($x+20) 295 34 3 $p.color
  $parts+=RG-Text "$id/detail-1" $p.a ($x+20) 316 221 26 15 '34465F'
  $parts+=RG-Text "$id/detail-2" $p.b ($x+20) 347 221 26 15 '34465F'
  $parts+=RG-Text "$id/detail-3" $p.c ($x+20) 378 221 26 15 '34465F'
  $names=[object[]]@($parts|ForEach-Object{$_.Name});$g=$slide.Shapes.Range($names).Group();$g.Name="stage/$id"
 }
 $null=RG-Arrow 'flow/reference-to-structure' @(@(290,263),@(325,263)) '73839B'
 $null=RG-Arrow 'flow/structure-to-drawing' @(@(582,263),@(617,263)) '73839B'
 $null=RG-Arrow 'flow/drawing-to-review' @(@(874,263),@(909,263)) '73839B'
 $null=RG-Arrow 'repair/review-to-structure' @(@(1038,430),@(1038,491),@(454,491),@(454,430)) '4B5BB5' $true
 $null=RG-Text 'repair/label' 'Local correction: update the affected component, then render again' 487 454 539 24 15 '4B5BB5'
 $null=RG-Box 'footer/rule' 36 540 1128 1 'DCE3EC'
 $null=RG-Text 'footer/priorities' 'Structure first. Native editability next. Visual polish after verification.' 36 556 1128 25 16 '52627A'
 $slide.NotesPage.Shapes.Placeholders.Item(2).TextFrame.TextRange.Text='Illustrative workflow for RouteGraph. Image interpretation and scientific judgement remain Agent-led. Instrument scenes use the scene compiler; ordinary diagrams use the existing topology-first script workflow. This figure does not claim automatic OCR, segmentation or scientific validation.'
 $pres.SaveAs($OutputPath,24);$slide.Export($preview,'PNG',2400,1200)
 Write-Output "PPTX=$OutputPath`nPREVIEW=$preview"
}finally{if($pres){$pres.Close()};if($app){[void][Runtime.InteropServices.Marshal]::ReleaseComObject($app)}}
