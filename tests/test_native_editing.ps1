param(
    [Parameter(Mandatory=$true)][string]$InputPath,
    [Parameter(Mandatory=$true)][string]$OutputPath,
    [string]$ComponentId='spacer',
    [string]$SurfaceName='spacer/outer_wall',
    [string]$LabelName='label/spacer',
    [int]$ExpectedPartCount=5
)
$ErrorActionPreference='Stop'
$InputPath=[IO.Path]::GetFullPath($InputPath); $OutputPath=[IO.Path]::GetFullPath($OutputPath)
if(Test-Path -LiteralPath $OutputPath) { throw 'Refusing overwrite' }
$hash=(Get-FileHash -LiteralPath $InputPath).Hash
$app=$null; $pres=$null
try {
    $app=New-Object -ComObject PowerPoint.Application
    $pres=$app.Presentations.Open($InputPath,-1,0,0)
    $slide=$pres.Slides.Item(1)
    $groupName='component/'+$ComponentId
    $g=$slide.Shapes.Item($groupName)
    if($g.Type -ne 6) { throw 'Component is not a native group' }
    $parts=$g.Ungroup()
    $count=$parts.Count
    if($count -ne $ExpectedPartCount) { throw 'Component lost expected native surfaces' }
    $wall=$slide.Shapes.Item($SurfaceName)
    $wall.Fill.ForeColor.RGB=255
    $expectedWidth=[single]($wall.Width+3)
    $wall.Width=$expectedWidth
    $g=$parts.Group(); $g.Name=$groupName
    $label=$slide.Shapes.Item($LabelName); $label.TextFrame.TextRange.Text='Editable component'
    $pres.SaveAs($OutputPath,24); $pres.Close(); $pres=$null
    $pres=$app.Presentations.Open($OutputPath,-1,0,0)
    if($pres.Slides.Item(1).Shapes.Item($LabelName).TextFrame.TextRange.Text -ne 'Editable component') { throw 'Text edit did not survive save/reopen' }
    $g=$pres.Slides.Item(1).Shapes.Item($groupName)
    if($g.GroupItems.Item($SurfaceName).Fill.ForeColor.RGB -ne 255) { throw 'Native fill edit did not survive save/reopen' }
    if([Math]::Abs($g.GroupItems.Item($SurfaceName).Width-$expectedWidth) -gt 0.05) { throw 'Native size edit did not survive save/reopen' }
    if((Get-FileHash -LiteralPath $InputPath).Hash -ne $hash) { throw 'Input file changed' }
    @{status='pass'; component=$ComponentId; native_surface_count=$ExpectedPartCount; checks=@('ungroup native component surfaces','change surface size and fill','edit plain text','regroup and save','reopen and verify'); source_unchanged=((Get-FileHash -LiteralPath $InputPath).Hash -eq $hash)} | ConvertTo-Json | Set-Content -LiteralPath ([IO.Path]::ChangeExtension($OutputPath,'.json')) -Encoding UTF8
} finally {
    if($pres) { $pres.Saved=-1; $pres.Close() }
    if($app) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($app) }
}
