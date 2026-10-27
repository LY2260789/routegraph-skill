param(
    [Parameter(Mandatory=$true)][string]$OutputDirectory,
    [string]$PythonExecutable='python'
)
$ErrorActionPreference='Stop'
$repo=Split-Path $PSScriptRoot -Parent
$out=[IO.Path]::GetFullPath($OutputDirectory)
if(Test-Path -LiteralPath $out) { throw 'Use a new output directory; regression artifacts are never overwritten' }
$null=New-Item -ItemType Directory -Path $out
function Invoke-CheckedPython([string[]]$Arguments) {
    & $PythonExecutable @Arguments
    if($LASTEXITCODE -ne 0) { throw "Python check failed: $Arguments" }
}
Invoke-CheckedPython @('-m','unittest','discover','-s',(Join-Path $repo 'tests'),'-v')
$scene=Join-Path $repo 'assets\examples\instrument_reconstruction\charging_collection.reviewed.scene.json'
$plan=Join-Path $out 'charging_collection.plan.json'
Invoke-CheckedPython @((Join-Path $repo 'scripts\reconstruction\scene.py'),$scene,'--out',$plan)
& (Join-Path $repo 'scripts\render_instrument.ps1') -PlanPath $plan -OutputPath (Join-Path $out 'charging_collection.pptx')
Invoke-CheckedPython @((Join-Path $repo 'scripts\reconstruction\verify.py'),(Join-Path $out 'charging_collection.pptx'),'--scene',$scene,'--plan',$plan,'--receipt',(Join-Path $out 'charging_collection.render.json'),'--out',(Join-Path $out 'charging_collection.validation.json'))
Invoke-CheckedPython @((Join-Path $repo 'tests\make_gallery.py'),$out)
& (Join-Path $repo 'scripts\render_instrument.ps1') -PlanPath (Join-Path $out 'primitives.plan.json') -OutputPath (Join-Path $out 'primitives.pptx')
Invoke-CheckedPython @((Join-Path $repo 'scripts\reconstruction\verify.py'),(Join-Path $out 'primitives.pptx'),'--scene',(Join-Path $out 'primitives.scene.json'),'--plan',(Join-Path $out 'primitives.plan.json'),'--receipt',(Join-Path $out 'primitives.render.json'),'--out',(Join-Path $out 'primitives.validation.json'))
# Do not change the legacy template. Only remove its process-wide Quit in a test copy.
$legacy=Get-Content -LiteralPath (Join-Path $repo 'assets\templates\continuous_arrow_regression.ps1') -Raw
$legacy=$legacy.Replace('$ppt.Quit()', '$null')
$legacyPath=Join-Path $out 'legacy_arrow_regression.ps1'
Set-Content -LiteralPath $legacyPath -Value $legacy -Encoding UTF8
# Legacy Boolean COM setters behave differently after typed COM bindings have
# been loaded by the new renderer in PowerShell 7. Run the unchanged drawing
# body in a fresh shell, as users normally run the standalone legacy script.
$shellExecutable=(Get-Process -Id $PID).Path
& $shellExecutable -NoProfile -ExecutionPolicy Bypass -File $legacyPath -OutputPath (Join-Path $out 'legacy_arrows.pptx')
if($LASTEXITCODE -ne 0) { throw 'Legacy template process failed' }
Invoke-CheckedPython @((Join-Path $repo 'tests\check_legacy_arrows.py'),(Join-Path $out 'legacy_arrows.pptx'),(Join-Path $out 'legacy_arrows.validation.json'))
& (Join-Path $repo 'tests\test_native_editing.ps1') -InputPath (Join-Path $out 'charging_collection.pptx') -OutputPath (Join-Path $out 'native_edit_roundtrip.pptx')
Write-Output "Regression artifacts: $out"
