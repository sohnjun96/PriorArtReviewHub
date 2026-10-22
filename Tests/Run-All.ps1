param([switch]$IncludeGui)
$ErrorActionPreference='Stop'
$taskEngine=(Get-Process -Id $PID).Path
foreach ($taskTestName in @('Storage.Tests.ps1','Workflow.Tests.ps1','Selection.Tests.ps1','Export.Tests.ps1','Gui.Structure.Tests.ps1','Deployment.Tests.ps1','ExecutionPolicy.Tests.ps1')) {
    $taskTestPath=Join-Path $PSScriptRoot $taskTestName
    if (-not (Test-Path -LiteralPath $taskTestPath)) {throw "필수 검증 파일이 없습니다: $taskTestName"}
    Write-Output "RUN $taskTestName"
    & $taskEngine -NoLogo -NoProfile -STA -ExecutionPolicy RemoteSigned -File $taskTestPath
    if ($LASTEXITCODE -ne 0) {throw "검증 실패: $taskTestName"}
}
if ($IncludeGui) {
    foreach ($taskGuiTest in @('Gui.Layout.Tests.ps1','Gui.Runtime.Tests.ps1')) {
        & $taskEngine -NoLogo -NoProfile -STA -ExecutionPolicy RemoteSigned -File (Join-Path $PSScriptRoot $taskGuiTest)
        if ($LASTEXITCODE -ne 0) {throw "화면 실행 검증 실패: $taskGuiTest"}
    }
}
Write-Output 'PASS all requested checks'
