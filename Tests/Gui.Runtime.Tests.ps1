$ErrorActionPreference='Stop'
$taskProject=Split-Path $PSScriptRoot -Parent
Import-Module (Join-Path $taskProject 'App\Modules\Workflow.psm1') -Force -DisableNameChecking
$taskRoot=Join-Path ([IO.Path]::GetTempPath()) ('AIReview-Gui-'+[guid]::NewGuid().ToString('N'))
$taskContext=Initialize-Demo -Root $taskRoot
$taskPeriod=Get-Date -Format 'yyyy-MM'
$taskCase=New-Submission -Context $taskContext -Period $taskPeriod -ApplicationNumber '10-2026-0000123' -RequestId ([guid]::NewGuid().ToString('N'))
$taskVersions=@{}
$taskRatings=@()
$taskRankings=@()
$taskIndex=0
foreach ($taskSource in @('S01','S02','S03')) {
    $taskIndex++
    $taskDoc=[pscustomobject]@{documentId=('D'+$taskIndex);type='특허';identifier=('KR10-000000'+$taskIndex);title=('화면 검증 문헌 '+$taskIndex);description='동일 출원의 조사 방식별 납품';attachments=@()}
    $taskDelivery=Save-Delivery -Context $taskContext -Case $taskCase -SourceId $taskSource -ExpectedVersion 0 -RequestId ([guid]::NewGuid().ToString('N')) -Status Published -Documents @($taskDoc)
    $taskVersions[$taskSource]=$taskDelivery.version
    $taskRatings += [pscustomobject]@{sourceId=$taskSource;documentId=$taskDoc.documentId;score=($taskIndex-1);cited=$false;note='0점과 X 보존 확인'}
    $taskRankings += [pscustomobject]@{sourceId=$taskSource;rank=$taskIndex;status='Ranked'}
}
$taskEvaluation=Save-Evaluation -Context $taskContext -Case $taskCase -ExpectedVersion 0 -RequestId ([guid]::NewGuid().ToString('N')) -DeliveryVersions $taskVersions -DocumentRatings $taskRatings -Rankings $taskRankings
if ($null -eq $taskEvaluation.data.documentRatings[0].score -or $taskEvaluation.data.documentRatings[0].score -ne 0) {throw '실제 0점이 보존되지 않았습니다.'}
$taskOutput=Join-Path $taskProject 'TestResults'
[void][IO.Directory]::CreateDirectory($taskOutput)
$taskScreenshot=Join-Path $taskOutput 'Gui.Smoke.png'
$taskStart=Join-Path $taskProject 'App\Start.ps1'
$taskProcessInfo=New-Object Diagnostics.ProcessStartInfo
$taskProcessInfo.FileName='powershell.exe'
$taskProcessInfo.Arguments=('-NoLogo -NoProfile -STA -ExecutionPolicy RemoteSigned -File "{0}" -Demo -DataRoot "{1}" -SmokeTest -ExerciseUiWrites -ScreenshotPath "{2}"' -f $taskStart,$taskRoot,$taskScreenshot)
$taskProcessInfo.UseShellExecute=$false
$taskProcessInfo.CreateNoWindow=$true
$taskProcessInfo.RedirectStandardOutput=$true
$taskProcessInfo.RedirectStandardError=$true
$taskProcess=[Diagnostics.Process]::Start($taskProcessInfo)
try {
    if (-not $taskProcess.WaitForExit(45000)) {$taskProcess.Kill();throw 'WPF 화면 검증이 45초 안에 완료되지 않았습니다.'}
    if ($taskProcess.ExitCode -ne 0) {throw ($taskProcess.StandardError.ReadToEnd()+$taskProcess.StandardOutput.ReadToEnd())}
    if (-not [IO.File]::Exists($taskScreenshot)) {throw '화면 이미지가 생성되지 않았습니다.'}
    $taskNewEvaluation=Get-Evaluation -Context $taskContext -Case $taskCase
    $taskNewDelivery=Get-DeliveryForEdit -Context $taskContext -Case $taskCase -SourceId S01
    if ($taskNewDelivery.version -ne 3 -or $taskNewEvaluation.version -ne 2 -or $taskNewEvaluation.needsReconfirmation) {throw '화면의 임시저장·재납품·평가 저장이 반영되지 않았습니다.'}
    if ($null -eq $taskNewEvaluation.data.documentRatings[0].score -or $taskNewEvaluation.data.documentRatings[0].score -ne 0 -or $taskNewEvaluation.data.documentRatings[0].cited -ne $false) {throw '화면 평가 저장에서 0점 또는 X가 바뀌었습니다.'}
    Write-Output 'PASS WPF runtime: load submission, draft save, republish, stale evaluation review, save evaluation, zero score, screenshot'
    Write-Output "Screenshot: $taskScreenshot"
} finally {$taskProcess.Dispose()}
