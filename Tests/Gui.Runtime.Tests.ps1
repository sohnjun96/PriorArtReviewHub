$ErrorActionPreference='Stop'
$taskProject=Split-Path $PSScriptRoot -Parent
Import-Module (Join-Path $taskProject 'App\Modules\Workflow.psm1') -Force -DisableNameChecking
$taskRoot=Join-Path ([IO.Path]::GetTempPath()) ('AIReview-Gui-'+[guid]::NewGuid().ToString('N'))
$taskContext=Initialize-Demo -Root $taskRoot
$taskProfile=@(Get-Users -Context $taskContext | Where-Object {$_.userId -eq 'U001'})[0]
$taskProfile | Add-Member -NotePropertyName bureau -NotePropertyValue '검증 심사국' -Force
$taskProfile | Add-Member -NotePropertyName division -NotePropertyValue '검증 심사과' -Force
[void](Save-MasterRecord -Context $taskContext -Kind Users -Id U001 -Data $taskProfile -ExpectedVersion $taskProfile.version -RequestId ([guid]::NewGuid().ToString('N')))
$taskContext=Get-Context -Root $taskRoot
$taskPeriod=Get-Date -Format 'yyyy-MM'
$taskCase=New-Submission -Context $taskContext -Period $taskPeriod -ApplicationNumber '10-2026-0000123' -Cpc 'G06F 16/00' -RequestId ([guid]::NewGuid().ToString('N'))
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
    $taskNewUser=@(Get-Users -Context $taskContext | Where-Object {$_.name -eq '화면 등록 심사관'})[0]
    if ($taskNewUser.name -ne '화면 등록 심사관' -or $taskNewUser.bureau -ne '화면 등록 심사국' -or $taskNewUser.division -ne '화면 등록 심사과') {throw '사용자 폼의 심사국·심사과가 저장되지 않았습니다.'}
    if ($taskNewUser.userId -notmatch '^U[a-f0-9]{32}$' -or $taskNewUser.sid -notmatch '^APP-[a-f0-9]{32}$') {throw '사용자 식별번호가 자동 생성되지 않았습니다.'}
    $taskNewCase=@(Get-Cases -Context $taskContext -Period $taskPeriod -Scope Mine | Where-Object {$_.applicationNumber -eq '10-2026-0000456'})[0]
    if ($taskNewCase.cpc -ne 'H04L 9/00' -or $taskNewCase.examinerBureau -ne '검증 심사국' -or $taskNewCase.examinerDivision -ne '검증 심사과') {throw '출원 화면의 CPC·소속 스냅샷이 저장되지 않았습니다.'}
    if ($taskNewDelivery.version -ne 3 -or $taskNewEvaluation.version -ne 2 -or $taskNewEvaluation.needsReconfirmation) {throw '화면의 임시저장·재납품·평가 저장이 반영되지 않았습니다.'}
    foreach ($taskIdentifier in @('10-2025-0000007','10-2025-0000008')) {
        $taskNumberOnly=@($taskNewDelivery.data.documents | Where-Object {$_.identifier -eq $taskIdentifier})
        if ($taskNumberOnly.Count -ne 1 -or $taskNumberOnly[0].title) {throw '출원번호만 입력한 AI 검색 결과가 저장되지 않았습니다.'}
    }
    if ($null -eq $taskNewEvaluation.data.documentRatings[0].score -or $taskNewEvaluation.data.documentRatings[0].score -ne 0 -or $taskNewEvaluation.data.documentRatings[0].cited -ne $false) {throw '화면 평가 저장에서 0점 또는 X가 바뀌었습니다.'}
    Write-Output 'PASS WPF runtime: selection login, automatic user IDs, bureau/division, CPC, number-only AI results, failed-row retention, evaluation, zero score, screenshot'
    Write-Output "Screenshot: $taskScreenshot"
} finally {$taskProcess.Dispose()}

# Exercise a second first-screen selection on the same PC. Its main window must
# receive the selected examiner's context and cases, without the admin's cases.
$taskOtherContext=Get-Context -Root $taskRoot -UserId U002
$taskOtherCase=New-Submission -Context $taskOtherContext -Period $taskPeriod -ApplicationNumber '10-2026-0000789' -Cpc 'G06N 3/00' -RequestId ([guid]::NewGuid().ToString('N'))
$taskOtherReport=Join-Path $taskOutput 'Gui.OtherUser.json'
$taskOtherScreenshot=Join-Path $taskOutput 'Gui.OtherUser.png'
$taskProcessInfo.Arguments=('-NoLogo -NoProfile -STA -ExecutionPolicy RemoteSigned -File "{0}" -Demo -DataRoot "{1}" -SmokeTest -SmokeLoginUserId U002 -SmokeReportPath "{2}" -ScreenshotPath "{3}"' -f $taskStart,$taskRoot,$taskOtherReport,$taskOtherScreenshot)
$taskProcess=[Diagnostics.Process]::Start($taskProcessInfo)
try {
    if (-not $taskProcess.WaitForExit(45000)) {$taskProcess.Kill();throw '두 번째 사용자 선택 화면 검증이 완료되지 않았습니다.'}
    if ($taskProcess.ExitCode -ne 0) {throw ($taskProcess.StandardError.ReadToEnd()+$taskProcess.StandardOutput.ReadToEnd())}
    $taskReport=ConvertFrom-Json -InputObject ([IO.File]::ReadAllText($taskOtherReport))
    if ($taskReport.userId -ne 'U002' -or $taskReport.adminVisible -or $taskOtherCase.caseId -notin $taskReport.caseIds -or $taskCase.caseId -in $taskReport.caseIds) {throw '첫 화면에서 선택한 심사관의 목록으로 전환되지 않았습니다.'}
    Write-Output 'PASS WPF selection: two examiner identities on one PC show their own cases; admin controls hidden for examiner'
} finally {$taskProcess.Dispose()}
