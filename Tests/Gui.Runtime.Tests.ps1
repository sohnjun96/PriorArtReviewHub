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
$taskBlankDocumentId='D-Blank'
foreach ($taskSource in @('S01','S02','S03')) {
    $taskIndex++
    $taskDoc=[pscustomobject]@{documentId=('D'+$taskIndex);type='특허';identifier=('KR10-000000'+$taskIndex);title=('화면 검증 문헌 '+$taskIndex);description='동일 출원의 조사 방식별 납품';attachments=@()}
    $taskDocuments=@($taskDoc)
    if ($taskSource -eq 'S03') {$taskDocuments += [pscustomobject]@{documentId=$taskBlankDocumentId;type='특허';identifier='KR10-0000099';title='미평가 유지 문헌';description='점수와 인용 여부의 공란 보존 확인';attachments=@()}}
    $taskDelivery=Save-Delivery -Context $taskContext -Case $taskCase -SourceId $taskSource -ExpectedVersion 0 -RequestId ([guid]::NewGuid().ToString('N')) -Status Published -Documents $taskDocuments
    $taskVersions[$taskSource]=$taskDelivery.version
    $taskRatings += [pscustomobject]@{sourceId=$taskSource;documentId=$taskDoc.documentId;score=($taskIndex-1);cited=$false;note='0점과 X 보존 확인'}
    $taskRankings += [pscustomobject]@{sourceId=$taskSource;rank=$taskIndex;status='Ranked'}
}
$taskEvaluation=Save-Evaluation -Context $taskContext -Case $taskCase -ExpectedVersion 0 -RequestId ([guid]::NewGuid().ToString('N')) -DeliveryVersions $taskVersions -DocumentRatings $taskRatings -Rankings $taskRankings
if ($null -eq $taskEvaluation.data.documentRatings[0].score -or $taskEvaluation.data.documentRatings[0].score -ne 0) {throw '실제 0점이 보존되지 않았습니다.'}
$taskOutput=Join-Path $taskProject 'TestResults'
[void][IO.Directory]::CreateDirectory($taskOutput)
$taskScreenshot=Join-Path $taskOutput 'Gui.Smoke.png'
$taskAdminReport=Join-Path $taskOutput 'Gui.AdminMode.json'
$taskStart=Join-Path $taskProject 'App\Start.ps1'
$taskProcessInfo=New-Object Diagnostics.ProcessStartInfo
$taskProcessInfo.FileName=(Get-Process -Id $PID).Path
$taskProcessInfo.Arguments=('-NoLogo -NoProfile -STA -ExecutionPolicy RemoteSigned -File "{0}" -Demo -AdminMode -DataRoot "{1}" -SmokeTest -ExerciseUiWrites -ScreenshotPath "{2}" -SmokeReportPath "{3}"' -f $taskStart,$taskRoot,$taskScreenshot,$taskAdminReport)
$taskProcessInfo.UseShellExecute=$false
$taskProcessInfo.CreateNoWindow=$true
$taskProcessInfo.RedirectStandardOutput=$true
$taskProcessInfo.RedirectStandardError=$true
# Exercise the actual GUI in the same engine as the test host. Hidden Windows
# PowerShell consoles use the OEM code page; PowerShell 7 writes UTF-8 instead.
$taskConsoleEncoding=if ([IO.Path]::GetFileName($taskProcessInfo.FileName) -ieq 'pwsh.exe') {New-Object Text.UTF8Encoding($false)} else {[Text.Encoding]::GetEncoding([Globalization.CultureInfo]::InstalledUICulture.TextInfo.OEMCodePage)}
$taskProcessInfo.StandardOutputEncoding=$taskConsoleEncoding
$taskProcessInfo.StandardErrorEncoding=$taskConsoleEncoding
$taskProcess=[Diagnostics.Process]::Start($taskProcessInfo)
try {
    if (-not $taskProcess.WaitForExit(60000)) {$taskProcess.Kill();throw 'WPF 화면 검증이 60초 안에 완료되지 않았습니다.'}
    if ($taskProcess.ExitCode -ne 0) {throw ($taskProcess.StandardError.ReadToEnd()+$taskProcess.StandardOutput.ReadToEnd())}
    if (-not [IO.File]::Exists($taskScreenshot)) {throw '화면 이미지가 생성되지 않았습니다.'}
    $taskAdminSession=ConvertFrom-Json -InputObject ([IO.File]::ReadAllText($taskAdminReport))
    if ($taskAdminSession.launchMode -ne 'Admin' -or $taskAdminSession.accessMode -ne 'Admin' -or -not $taskAdminSession.adminVisible -or $taskAdminSession.loginDialogShown) {throw '관리자용 실행이 사용자 선택 없이 관리자 화면으로 진입하지 않았습니다.'}
    if (($taskAdminSession.visibleMenus -join '|') -cne '내 심사 업무|AI 담당 업무|운영 관리' -or -not $taskAdminSession.deliveryVisible) {throw '관리자 실행의 개인·AI·운영 업무 구분이 적용되지 않았습니다.'}
    if (($taskAdminSession.personalSteps -join '|') -cne '1. 대상건 선정|2. AI 검색 결과 평가') {throw '심사 업무의 대상건 선정과 결과 평가 단계가 분리되지 않았습니다.'}
    if (-not $taskAdminSession.dropdownBindingsVerified) {throw '평가 드롭다운을 선택한 값이 실제 평가 행에 즉시 반영되지 않았습니다.'}
    if (-not $taskAdminSession.singleClickEditingVerified) {throw '문헌번호 셀을 한 번 클릭해 바로 수정할 수 없습니다.'}
    if (-not $taskAdminSession.automaticDeliveryRefreshVerified) {throw 'AI 검색 결과를 저장한 후 선택 출원의 결과·평가 화면이 자동으로 갱신되지 않았습니다.'}
    if (-not $taskAdminSession.automaticUserLoadVerified) {throw '사용자를 드롭다운에서 선택한 뒤 추가 불러오기 없이 수정할 수 없습니다.'}
    if (-not $taskAdminSession.adminFormDisplayVerified) {throw '월 운영의 참여자·AI 배정 폼을 조회한 것만으로 미저장 변경이 발생했습니다.'}
    if (-not $taskAdminSession.caseScrollVerified) {throw '다른 출원으로 이동할 때 화면 위치 초기화 또는 같은 출원 재조회 시 위치 유지에 실패했습니다.'}
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
    $taskBlankRating=@($taskNewEvaluation.data.documentRatings | Where-Object {$_.sourceId -eq 'S03' -and $_.documentId -eq $taskBlankDocumentId})
    if ($taskBlankRating.Count -ne 1 -or $null -ne $taskBlankRating[0].score -or $null -ne $taskBlankRating[0].cited) {throw '드롭다운의 미평가 값이 0점이나 X로 저장되었습니다.'}
    Write-Output ('PASS WPF admin mode (PowerShell {0}): direct login, user choice loads editable form, automatic user IDs, CPC, number-only results, automatic refresh, dropdown evaluation, zero/X and unevaluated values, screenshot' -f $PSVersionTable.PSVersion)
    Write-Output "Screenshot: $taskScreenshot"
} finally {$taskProcess.Dispose()}

# Start with no submissions, persist a draft, close the app, then confirm the restored draft.
$taskSelectionRoot=Join-Path ([IO.Path]::GetTempPath()) ('AIReview-GuiSelection-'+[guid]::NewGuid().ToString('N'))
$taskSelectionAdmin=Initialize-Demo -Root $taskSelectionRoot
$taskSelectionUser=Get-Context -Root $taskSelectionRoot -UserId U002 -AccessMode User
foreach ($taskFlow in @('Draft','Confirm')) {
    $taskSelectionReport=Join-Path $taskOutput ("Gui.Selection-$taskFlow.json")
    $taskSelectionImage=Join-Path $taskOutput ("Gui.Selection-$taskFlow.png")
    $taskProcessInfo.Arguments=('-NoLogo -NoProfile -STA -ExecutionPolicy RemoteSigned -File "{0}" -Demo -DataRoot "{1}" -SmokeTest -SmokeLoginUserId U002 -ExerciseTargetSelection {2} -SmokeReportPath "{3}" -ScreenshotPath "{4}"' -f $taskStart,$taskSelectionRoot,$taskFlow,$taskSelectionReport,$taskSelectionImage)
    $taskProcess=[Diagnostics.Process]::Start($taskProcessInfo)
    try {
        if (-not $taskProcess.WaitForExit(60000)) {$taskProcess.Kill();throw '대상건 임시저장·확정 화면 검증 시간이 초과되었습니다.'}
        if ($taskProcess.ExitCode -ne 0) {throw ($taskProcess.StandardError.ReadToEnd()+$taskProcess.StandardOutput.ReadToEnd())}
        $taskSelectionResult=ConvertFrom-Json -InputObject ([IO.File]::ReadAllText($taskSelectionReport))
        if (-not $taskSelectionResult.selectionFlowVerified) {throw '대상건 임시저장 또는 재로그인 후 확정 흐름을 검증하지 못했습니다.'}
        $taskSavedSelection=Get-TargetSelection $taskSelectionUser $taskPeriod
        if ($taskFlow -eq 'Draft') {
            if ($taskSavedSelection.status -ne 'Draft' -or $taskSelectionResult.sidebarVisible -or ($taskSelectionResult.visiblePersonalSteps -join '|') -ne '1. 대상건 선정' -or @(Get-Cases $taskSelectionAdmin $taskPeriod All).Count -ne 0) {throw '미확정 화면 또는 초안 비공개 동작이 올바르지 않습니다.'}
        } else {
            if ($taskSavedSelection.status -ne 'Confirmed' -or -not $taskSelectionResult.evaluationOpened -or $taskSelectionResult.ratingsHeight -le $taskSelectionResult.rankHeight -or ($taskSelectionResult.rankSourceIds -join '|') -ne 'S01|S02|S03') {throw '확정 뒤 평가 화면·전체 조사 방식·평가 영역 크기가 올바르지 않습니다.'}
        }
        Write-Output "PASS WPF monthly selection $taskFlow, persistent draft, stage visibility and evaluation layout"
    } finally {$taskProcess.Dispose()}
}
$taskSelectionCase=@(Get-Cases $taskSelectionUser $taskPeriod Mine)[0]
foreach ($taskSource in @('S02','S03')) {
    [void](Save-Delivery -Context $taskSelectionAdmin -Case $taskSelectionCase -SourceId $taskSource -ExpectedVersion 0 -RequestId ([guid]::NewGuid().ToString('N')) -Status Published -Documents @([pscustomobject]@{documentId='D1';identifier='10-2025-0000001';type='특허';attachments=@()}))
}
$taskSelectionReport=Join-Path $taskOutput 'Gui.Selection-Delivered.json'
$taskProcessInfo.Arguments=('-NoLogo -NoProfile -STA -ExecutionPolicy RemoteSigned -File "{0}" -Demo -DataRoot "{1}" -SmokeTest -SmokeLoginUserId U002 -SmokeReportPath "{2}"' -f $taskStart,$taskSelectionRoot,$taskSelectionReport)
$taskProcess=[Diagnostics.Process]::Start($taskProcessInfo)
try {
    if (-not $taskProcess.WaitForExit(60000)) {$taskProcess.Kill();throw '확정 후 재로그인 검증 시간이 초과되었습니다.'}
    if ($taskProcess.ExitCode -ne 0) {throw ($taskProcess.StandardError.ReadToEnd()+$taskProcess.StandardOutput.ReadToEnd())}
    $taskSelectionResult=ConvertFrom-Json -InputObject ([IO.File]::ReadAllText($taskSelectionReport))
    if (-not $taskSelectionResult.evaluationOpened -or ($taskSelectionResult.caseStatusLabels -join '|') -ne '납품 완료' -or ($taskSelectionResult.rankSourceIds -join '|') -ne 'S01|S02|S03') {throw '확정 후 재로그인 또는 AI 납품 완료 상태와 선행기술조사 행이 올바르지 않습니다.'}
    Write-Output 'PASS WPF confirmed re-login opens evaluation with completed AI status and undelivered general-search comparator'
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
    if (-not $taskProcess.WaitForExit(60000)) {$taskProcess.Kill();throw '두 번째 사용자 선택 화면 검증이 완료되지 않았습니다.'}
    if ($taskProcess.ExitCode -ne 0) {throw ($taskProcess.StandardError.ReadToEnd()+$taskProcess.StandardOutput.ReadToEnd())}
    $taskReport=ConvertFrom-Json -InputObject ([IO.File]::ReadAllText($taskOtherReport))
    if ($taskReport.userId -ne 'U002' -or $taskReport.adminVisible -or $taskOtherCase.caseId -notin $taskReport.caseIds -or $taskCase.caseId -in $taskReport.caseIds -or $taskReport.launchMode -ne 'User' -or $taskReport.accessMode -ne 'User' -or -not $taskReport.loginDialogShown) {throw '일반사용자용 실행에서 본인 선택과 업무 범위가 적용되지 않았습니다.'}
    if ($taskReport.deliveryVisible -or ($taskReport.visibleMenus -join '|') -cne '내 심사 업무' -or $taskReport.workspaceScope -ne 'Mine' -or ($taskReport.scopeIds -join '|') -cne 'Mine') {throw '일반 심사관에게 AI 담당 업무가 보이거나 본인 외의 출원 범위가 노출되었습니다.'}
    if (($taskReport.personalSteps -join '|') -cne '1. 대상건 선정|2. AI 검색 결과 평가' -or -not $taskReport.evaluationOpened -or ($taskReport.visiblePersonalSteps -join '|') -cne '2. AI 검색 결과 평가') {throw '이미 제출한 운영월은 평가 화면으로 바로 진입해야 합니다.'}
    Write-Output 'PASS WPF examiner mode: only personal workspace, own cases, separate selection/evaluation steps, no AI result entry'
} finally {$taskProcess.Dispose()}

# An AI examiner has two independent jobs. The AI workspace must include the
# assigned group's cases and only the assigned AI source. Returning to personal
# selection must restore only that examiner's own cases.
$taskAiContext=Get-Context -Root $taskRoot -UserId U003 -AccessMode User
$taskAiCase=New-Submission -Context $taskAiContext -Period $taskPeriod -ApplicationNumber '10-2026-0000987' -Cpc 'G06F 40/30' -RequestId ([guid]::NewGuid().ToString('N'))
$taskExpectedAssignedIds=@(Get-Cases -Context $taskAiContext -Period $taskPeriod -Scope Assigned | ForEach-Object {$_.caseId} | Sort-Object)
$taskExpectedPersonalIds=@(Get-Cases -Context $taskAiContext -Period $taskPeriod -Scope Mine | ForEach-Object {$_.caseId} | Sort-Object)
$taskAiReport=Join-Path $taskOutput 'Gui.AiWorkspace.json'
$taskAiScreenshot=Join-Path $taskOutput 'Gui.AiWorkspace.png'
$taskProcessInfo.Arguments=('-NoLogo -NoProfile -STA -ExecutionPolicy RemoteSigned -File "{0}" -Demo -DataRoot "{1}" -SmokeTest -SmokeLoginUserId U003 -ExerciseWorkspaceSwitch -SmokeReportPath "{2}" -ScreenshotPath "{3}"' -f $taskStart,$taskRoot,$taskAiReport,$taskAiScreenshot)
$taskProcess=[Diagnostics.Process]::Start($taskProcessInfo)
try {
    if (-not $taskProcess.WaitForExit(60000)) {$taskProcess.Kill();throw 'AI 담당 업무와 개인 업무 전환 검증이 60초 안에 완료되지 않았습니다.'}
    if ($taskProcess.ExitCode -ne 0) {throw ($taskProcess.StandardError.ReadToEnd()+$taskProcess.StandardOutput.ReadToEnd())}
    $taskReport=ConvertFrom-Json -InputObject ([IO.File]::ReadAllText($taskAiReport))
    if ($taskReport.userId -ne 'U003' -or $taskReport.adminVisible -or -not $taskReport.deliveryVisible -or ($taskReport.visibleMenus -join '|') -cne '내 심사 업무|AI 담당 업무') {throw 'AI 담당자에게 개인 업무와 AI 담당 업무가 따로 표시되지 않았습니다.'}
    if (-not $taskReport.workspaceSeparationVerified -or $taskReport.personalScope -ne 'Mine' -or $taskReport.aiScope -ne 'Assigned' -or $taskReport.workspaceScope -ne 'Mine') {throw 'AI 담당 업무와 개인 업무의 조회 범위가 전환에 따라 분리되지 않았습니다.'}
    if ((@($taskReport.aiCaseIds | Sort-Object) -join '|') -cne ($taskExpectedAssignedIds -join '|') -or $taskOtherCase.caseId -notin $taskReport.aiCaseIds) {throw 'AI 담당 업무에서 배정 그룹의 다른 심사관 출원을 조회하지 못했습니다.'}
    if ((@($taskReport.personalCaseIds | Sort-Object) -join '|') -cne ($taskExpectedPersonalIds -join '|') -or $taskAiCase.caseId -notin $taskReport.personalCaseIds -or $taskOtherCase.caseId -in $taskReport.personalCaseIds -or (@($taskReport.caseIds | Sort-Object) -join '|') -cne ($taskExpectedPersonalIds -join '|')) {throw '개인 업무로 돌아온 뒤 본인 출원 외의 담당 그룹 출원이 남아 있습니다.'}
    if (($taskReport.aiSourceIds -join '|') -cne 'S03') {throw 'AI 담당 업무에서 배정되지 않은 AI 또는 일반 조사 방식이 노출되었습니다.'}
    if (($taskReport.personalSteps -join '|') -cne '1. 대상건 선정|2. AI 검색 결과 평가' -or -not $taskReport.evaluationOpened -or ($taskReport.visiblePersonalSteps -join '|') -cne '2. AI 검색 결과 평가') {throw 'AI 담당자의 개인 대상건 선정과 결과 평가 단계가 섞였습니다.'}
    Write-Output 'PASS WPF AI examiner mode: assigned group/source entry, independent personal selection/evaluation, own-case scope restored after workspace switch'
} finally {$taskProcess.Dispose()}

# Selecting a user who holds both roles through the user launcher must still
# remain ordinary mode. The admin launcher is the only UI entry for admin work.
$taskRestrictedReport=Join-Path $taskOutput 'Gui.DualRoleUser.json'
$taskRestrictedScreenshot=Join-Path $taskOutput 'Gui.DualRoleUser.png'
$taskProcessInfo.Arguments=('-NoLogo -NoProfile -STA -ExecutionPolicy RemoteSigned -File "{0}" -Demo -DataRoot "{1}" -SmokeTest -SmokeLoginUserId U001 -SmokeReportPath "{2}" -ScreenshotPath "{3}"' -f $taskStart,$taskRoot,$taskRestrictedReport,$taskRestrictedScreenshot)
$taskProcess=[Diagnostics.Process]::Start($taskProcessInfo)
try {
    if (-not $taskProcess.WaitForExit(60000)) {$taskProcess.Kill();throw '중복 역할 일반사용자 화면 검증이 완료되지 않았습니다.'}
    if ($taskProcess.ExitCode -ne 0) {throw ($taskProcess.StandardError.ReadToEnd()+$taskProcess.StandardOutput.ReadToEnd())}
    $taskReport=ConvertFrom-Json -InputObject ([IO.File]::ReadAllText($taskRestrictedReport))
    if ($taskReport.userId -ne 'U001' -or $taskReport.adminVisible -or $taskReport.accessMode -ne 'User' -or 'All' -in $taskReport.scopeIds -or -not $taskReport.loginDialogShown) {throw '일반사용자용 실행에서 관리자 역할이 활성화되었습니다.'}
    if (($taskReport.visibleMenus -join '|') -cne '내 심사 업무|AI 담당 업무' -or $taskReport.workspaceScope -ne 'Mine') {throw '중복 역할 사용자도 개인 업무와 AI 담당 업무를 구분해야 합니다.'}
    Write-Output 'PASS WPF user mode: selecting an admin/examiner account does not enable administrator menus or full-scope access'
} finally {$taskProcess.Dispose()}
