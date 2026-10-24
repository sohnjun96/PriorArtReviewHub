[CmdletBinding()]
param(
    [string]$OutputDirectory,
    [ValidatePattern('^\d{4}-(0[1-9]|1[0-2])$')][string]$Period=(Get-Date -Format 'yyyy-MM'),
    [switch]$IncludeGui
)
$ErrorActionPreference='Stop'
Set-StrictMode -Version 2.0
$project=Split-Path $PSScriptRoot -Parent
Import-Module (Join-Path $project 'App/Modules/Workflow.psm1') -Force -DisableNameChecking
if (-not $OutputDirectory) {$OutputDirectory=Join-Path $project ('TestResults/FullFlow-'+(Get-Date -Format 'yyyyMMdd-HHmmss')+'-'+[guid]::NewGuid().ToString('N').Substring(0,6))}
$output=[IO.Path]::GetFullPath($OutputDirectory)
if (Test-Path -LiteralPath $output) {throw '새 테스트 폴더를 지정하세요. 기존 데이터는 덮어쓰지 않습니다.'}
[void][IO.Directory]::CreateDirectory($output)
$root=Join-Path $output '04-Evaluated/Data'
$utf8=New-Object Text.UTF8Encoding($true)
$checks=New-Object 'System.Collections.Generic.List[string]'
$stages=New-Object 'System.Collections.Generic.List[object]'
$gui=New-Object 'System.Collections.Generic.List[object]'
function Request {[guid]::NewGuid().ToString('N')}
function Check([bool]$Condition,[string]$Message) {
    if (-not $Condition) {throw "FAIL $Message"}
    $checks.Add($Message)
}
function Reject([scriptblock]$Action,[string]$Message) {
    $failed=$false;try {& $Action | Out-Null} catch {$failed=$true}
    Check $failed $Message
}
function Write-Launchers([string]$Folder) {
    $entry=(Join-Path $project 'App/Start.ps1').Replace("'","''")
    $scriptText="param([switch]`$AdminMode)`r`n& '$entry' -DataRoot (Join-Path `$PSScriptRoot 'Data') -AdminMode:`$AdminMode`r`n"
    [IO.File]::WriteAllText((Join-Path $Folder 'Open.ps1'),$scriptText,$utf8)
    foreach ($mode in @('User','Admin')) {
        $flag=if ($mode -eq 'Admin') {' -AdminMode'} else {''}
        $cmd="@echo off`r`npowershell.exe -NoLogo -NoProfile -STA -ExecutionPolicy RemoteSigned -File `"%~dp0Open.ps1`"$flag`r`nif errorlevel 1 pause`r`n"
        [IO.File]::WriteAllText((Join-Path $Folder ($mode+'.cmd')),$cmd,[Text.Encoding]::ASCII)
    }
}
function Save-Stage([string]$FolderName,[string]$Label) {
    $folder=Join-Path $output $FolderName
    if ($FolderName -ne '04-Evaluated') {
        [void][IO.Directory]::CreateDirectory($folder)
        Copy-Item -LiteralPath $root -Destination (Join-Path $folder 'Data') -Recurse
    }
    Write-Launchers $folder
    $view=Select-AdminStatistics (Get-AdminStatistics $admin $Period)
    $stages.Add([pscustomobject]@{stage=$Label;folder=$FolderName;summary=$view.summary;groups=$view.groups})
    Write-Output "PASS $Label"
}
function Documents([int]$CaseIndex,[int]$SourceIndex) {
    # Deliberately synthetic, number-only identifiers; no AI search is performed.
    @(1..7 | ForEach-Object {[pscustomobject]@{documentId=('D'+$_);identifier=('10-2025-{0:0000000}' -f (9000000+$CaseIndex*100+$SourceIndex*10+$_));type='특허';attachments=@()}})
}
function Test-Window([string]$Stage,[string]$UserId,[string]$Workspace,[string]$Label,[switch]$Admin) {
    $folder=Join-Path $output 'Screenshots';[void][IO.Directory]::CreateDirectory($folder)
    $report=Join-Path $folder ($Label+'.json');$image=Join-Path $folder ($Label+'.png')
    $info=New-Object Diagnostics.ProcessStartInfo
    $info.FileName=(Get-Process -Id $PID).Path
    $info.Arguments=('-NoLogo -NoProfile -STA -ExecutionPolicy RemoteSigned -File "{0}" -DataRoot "{1}" -SmokeTest -SmokeLoginUserId "{2}" -SmokeWorkspace {3} -SmokeReportPath "{4}" -ScreenshotPath "{5}"' -f (Join-Path $project 'App/Start.ps1'),(Join-Path $output "$Stage/Data"),$UserId,$Workspace,$report,$image)
    if ($Admin) {$info.Arguments+=' -AdminMode'}
    $info.UseShellExecute=$false;$info.CreateNoWindow=$true
    $info.RedirectStandardError=$true;$info.RedirectStandardOutput=$true
    $encoding=if ([IO.Path]::GetFileName($info.FileName) -ieq 'pwsh.exe') {New-Object Text.UTF8Encoding($false)} else {[Text.Encoding]::GetEncoding([Globalization.CultureInfo]::InstalledUICulture.TextInfo.OEMCodePage)}
    $info.StandardErrorEncoding=$encoding;$info.StandardOutputEncoding=$encoding
    $process=[Diagnostics.Process]::Start($info)
    try {
        $errors=$process.StandardError.ReadToEndAsync();$stdout=$process.StandardOutput.ReadToEndAsync()
        if (-not $process.WaitForExit(45000)) {$process.Kill();throw "WPF 검증 시간 초과: $Label"}
        if ($process.ExitCode -ne 0) {throw ($errors.Result+$stdout.Result)}
    } finally {$process.Dispose()}
    $result=ConvertFrom-Json ([IO.File]::ReadAllText($report))
    $gui.Add([pscustomobject]@{label=$Label;report=$report;image=$image})
    Write-Host "PASS WPF $Label"
    return $result
}

$null=& (Join-Path $PSScriptRoot 'Initialize-Store.ps1') -Root $root -AdminName '테스트 관리자' -Period $Period
$admin=Get-Context -Root $root -AccessMode Admin
$profile=@(Get-Users $admin | Where-Object {$_.userId -eq $admin.User.userId})[0]
$profile.roles=@('admin')
$null=Save-MasterRecord $admin Users $profile.userId $profile $profile.version (Request)
$admin=Get-Context -Root $root -AccessMode Admin
$group=Get-Groups $admin | Where-Object {$_.groupId -eq 'G01'}
$group.name='1그룹 · 전기·정보'
$null=Save-MasterRecord $admin Groups G01 $group $group.version (Request)
$null=Save-MasterRecord $admin Groups G02 @{groupId='G02';name='2그룹 · 기계·화학';active=$true} 0 (Request)
$people=New-Object 'System.Collections.Generic.List[object]'
$names=@('김가람','이서준','박민서','최지우','정하늘','강도윤','조수빈','윤서아','장현우','임다은')
$divisions=@('전기심사과','정보심사과','정보심사과','전기심사과','정보심사과','기계심사과','화학심사과','기계심사과','화학심사과','기계심사과')
for ($i=0;$i -lt 10;$i++) {
    $groupId=if ($i -lt 5) {'G01'} else {'G02'}
    $sourceId=if ($i -in @(0,5)) {'S02'} elseif ($i -in @(1,6)) {'S03'} else {''}
    $roles=if ($sourceId) {@('examiner','ai')} else {@('examiner')}
    $bureau=if ($i -lt 5) {'테스트 전기정보심사국'} else {'테스트 기계화학심사국'}
    $user=New-User $admin -Name ('테스트 '+$names[$i]) -Bureau $bureau -Division $divisions[$i] -Roles $roles -RequestId (Request)
    $people.Add([pscustomobject]@{userId=$user.userId;name=$user.name;division=$divisions[$i];groupId=$groupId;sourceId=$sourceId})
}
$settings=Get-Period $admin $Period
$settings.members=@($people | ForEach-Object {@{userId=$_.userId;groupId=$_.groupId}})
$settings.assignments=@($people | Where-Object {$_.sourceId} | ForEach-Object {@{userId=$_.userId;groupId=$_.groupId;sourceId=$_.sourceId}})
$null=Save-Period $admin $Period $settings $settings.version (Request)
Check (@(Get-Users $admin | Where-Object {'examiner' -in $_.roles}).Count -eq 10) '심사관 10명, 별도 관리자 1명'
Check (@($settings.assignments).Count -eq 4) '두 그룹에 AI 담당자 각 2명'

# Stage 1: everyone saves and reloads their own two-case draft.
for ($i=0;$i -lt 10;$i++) {
    $ctx=Get-Context -Root $root -UserId $people[$i].userId -AccessMode User
    $rows=@(1..2 | ForEach-Object {@{applicationNumber=('10-{0}-{1:0000000}' -f $Period.Substring(0,4),(8000000+$i*2+$_));cpc=$(if ($i -lt 5) {'G06F 16/00'} else {'B25J 9/00'})}})
    $saved=Save-TargetSelection $ctx $Period 0 (Request) Draft $rows
    $fresh=Get-TargetSelection (Get-Context -Root $root -UserId $people[$i].userId -AccessMode User) $Period
    Check ($saved.status -eq 'Draft' -and $fresh.rows.Count -eq 2 -and $fresh.rows[0].cpc -eq $rows[0].cpc) ($people[$i].name+' 대상 건 임시저장·재조회')
}
Check (@(Get-Cases $admin $Period All).Count -eq 0) '미확정 대상 건은 관리자·AI 대상 목록에 나타나지 않음'
foreach ($person in @($people | Where-Object {$_.sourceId})) {
    $ctx=Get-Context -Root $root -UserId $person.userId -AccessMode User
    Check (@(Get-Cases $ctx $Period Assigned).Count -eq 0) ($person.name+' 미확정 목록 비공개')
}
Save-Stage '01-SelectionDraft' '대상 건 선정 · 임시저장'

# Stage 2: confirmation is idempotent and each handler sees only their group.
foreach ($person in $people) {
    $ctx=Get-Context -Root $root -UserId $person.userId -AccessMode User
    $draft=Get-TargetSelection $ctx $Period;$request=Request
    $saved=Save-TargetSelection $ctx $Period $draft.version $request Confirmed $draft.rows
    $retry=Save-TargetSelection $ctx $Period $draft.version $request Confirmed $draft.rows
    Check ($saved.caseIds.Count -eq 2 -and $retry.version -eq $saved.version -and @(Get-Cases $ctx $Period Mine).Count -eq 2) ($person.name+' 선정 확정·중복 재시도')
}
$cases=@(Get-Cases $admin $Period All | Sort-Object applicationNumber)
Check ($cases.Count -eq 20) '선정 확정 총 20건'
foreach ($person in @($people | Where-Object {$_.sourceId})) {
    $ctx=Get-Context -Root $root -UserId $person.userId -AccessMode User
    $assigned=@(Get-Cases $ctx $Period Assigned)
    Check ($assigned.Count -eq 10 -and @($assigned | Where-Object {$_.groupId -ne $person.groupId}).Count -eq 0) ($person.name+' 담당 그룹 10건만 조회')
}
$g1ai=Get-Context -Root $root -UserId $people[0].userId -AccessMode User
$ordinary=Get-Context -Root $root -UserId $people[2].userId -AccessMode User
$g2case=@($cases | Where-Object {$_.groupId -eq 'G02'})[0]
$g1case=$cases[0];$docs=Documents 0 2
Reject {Save-Delivery $g1ai $g2case S02 0 (Request) Published -Documents $docs} '다른 그룹 AI 결과 입력 거부'
Reject {Save-Delivery $g1ai $g1case S03 0 (Request) Published -Documents $docs} '미배정 AI 방식 입력 거부'
Reject {Save-Delivery $ordinary $g1case S02 0 (Request) Published -Documents $docs} '일반 심사관의 AI 담당 업무 입력 거부'
$tooMany=@($docs)+@([pscustomobject]@{documentId='D8';identifier='10-2025-9999999';type='특허';attachments=@()})
Reject {Save-Delivery $g1ai $g1case S02 0 (Request) Published -Documents $tooMany} '조사 방식별 선행기술 문헌 8건 입력 거부'
Save-Stage '02-SelectionConfirmed' '선정 확정 · AI 검색 결과 대기'

# Stage 3: handlers enter seven number-only documents per AI and target case.
for ($i=0;$i -lt $cases.Count;$i++) {
    $case=$cases[$i];$owner=Get-Context -Root $root -UserId $case.userId -AccessMode User
    foreach ($sourceId in @('S02','S03')) {
        $person=@($people | Where-Object {$_.groupId -eq $case.groupId -and $_.sourceId -eq $sourceId})[0]
        $ctx=Get-Context -Root $root -UserId $person.userId -AccessMode User
        $docs=Documents $i ([int]$sourceId.Substring(1))
        $draft=Save-Delivery $ctx $case $sourceId 0 (Request) Draft -Documents $docs
        Check (@(Get-Deliveries $owner $case | Where-Object {$_.data.sourceId -eq $sourceId}).Count -eq 0) ($case.applicationNumber+' '+$sourceId+' 임시저장 비공개')
        $saved=Save-Delivery $ctx $case $sourceId $draft.version (Request) Published -Documents $docs
        Check ($saved.actorId -eq $person.userId -and $saved.data.documents.Count -eq 7) ($case.applicationNumber+' '+$sourceId+' 담당자 입력 완료')
    }
    # Existing prior-art search is entered by the examiner, for the three-way ranking.
    $null=Save-Delivery $owner $case S01 0 (Request) Published -Documents (Documents $i 1)
    Check ((Get-CaseStatus $owner $case).deliveryStatusLabel -eq '입력 완료') ($case.applicationNumber+' AI 검색 결과 입력 완료 상태')
}
$view=Select-AdminStatistics (Get-AdminStatistics $admin $Period)
Check ($view.summary.deliveredCount -eq 20 -and $view.summary.unreviewedCount -eq 20) 'AI 검색 결과 입력 20건 완료 · 아직 평가 0건'
Save-Stage '03-ResultsEntered' 'AI 검색 결과 입력 완료 · 평가 전'

# Stage 4: every examiner, including the four handlers, evaluates their own cases.
foreach ($person in $people) {
    $ctx=Get-Context -Root $root -UserId $person.userId -AccessMode User
    foreach ($case in @(Get-Cases $ctx $Period Mine)) {
        $versions=@{};$ratings=@();$rankings=@();$index=0
        foreach ($delivery in @(Get-Deliveries $ctx $case | Sort-Object {$_.data.sourceId})) {
            $sourceId=$delivery.data.sourceId;$versions[$sourceId]=$delivery.version
            $rankings+=@{sourceId=$sourceId;rank=([int]$sourceId.Substring(1));status='Ranked'}
            foreach ($doc in $delivery.data.documents) {
                $ratings+=@{sourceId=$sourceId;documentId=$doc.documentId;score=($index%6);cited=($index%2 -eq 1);note='테스트용 가상 선행기술 문헌 평가'};$index++
            }
        }
        $saved=Save-Evaluation $ctx $case 0 (Request) $versions -DocumentRatings $ratings -Rankings $rankings
        Check ($saved.data.complete -and -not $saved.needsReconfirmation -and $saved.data.documentRatings.Count -eq 21) ($person.name+' '+$case.applicationNumber+' 평가 완료')
        $zero=@($saved.data.documentRatings | Where-Object {$_.score -eq 0})
        Check ($zero.Count -gt 0 -and $zero[0].cited -eq $false) ($case.applicationNumber+' 실제 0점·인용 X 보존')
    }
}
Reject {Save-Evaluation $ordinary $cases[0] 0 (Request) @{} -DocumentRatings @() -Rankings @()} '다른 심사관의 평가 수정 거부'
$snapshot=Get-AdminStatistics $admin $Period;$view=Select-AdminStatistics $snapshot
Check ($view.summary.participantCount -eq 10 -and $view.summary.caseCount -eq 20 -and $view.summary.completeCount -eq 20 -and $view.summary.evaluationRate -eq '100%') '전체 통계: 심사관 10명 · 대상 건 20건 · 평가 완료율 100%'
foreach ($groupId in @('G01','G02')) {
    $filtered=Select-AdminStatistics $snapshot -GroupId $groupId
    Check ($filtered.summary.participantCount -eq 5 -and $filtered.summary.caseCount -eq 10 -and $filtered.summary.completeCount -eq 10) ($groupId+' 통계: 심사관 5명 · 대상 건 10건 · 평가 완료 10건')
}
foreach ($person in $people) {
    $filtered=Select-AdminStatistics $snapshot -UserId $person.userId
    Check ($filtered.summary.caseCount -eq 2 -and $filtered.summary.completeCount -eq 2) ($person.name+' 계정별 본인 대상 건·평가 2건')
    if ($person.sourceId) {Check ($filtered.accounts[0].aiAssigned -eq 10 -and $filtered.accounts[0].aiPublished -eq 10) ($person.name+' 본인 평가와 담당 AI 결과 입력 10건 분리 집계')}
}
Check (($view.sources | Measure-Object documentCount -Sum).Sum -eq 420 -and ($view.sources | Measure-Object scoreCount -Sum).Sum -eq 420) '선행기술조사 140건 + AI 검색 결과 280건 = 문헌별 평가 420건'
Save-Stage '04-Evaluated' 'AI 검색 결과 평가 완료'

if ($IncludeGui) {
    $id=$people[2].userId
    $screen=Test-Window '01-SelectionDraft' $id Selection '01-SelectionDraft'
    Check ($screen.selectionRowCount -eq 2 -and -not $screen.sidebarVisible -and $screen.visiblePersonalSteps.Count -eq 1 -and -not $screen.deliveryVisible) 'WPF 임시저장 복원 · 일반 심사관에게 선정 화면만 표시'
    $screen=Test-Window '02-SelectionConfirmed' $id Evaluation '02-SelectionConfirmed'
    Check ($screen.evaluationOpened -and $screen.caseIds.Count -eq 2 -and @($screen.caseStatusLabels | Where-Object {$_ -ne '대기 중'}).Count -eq 0) 'WPF 선정 확정 후 본인 대상 건 2건 · 대기 중'
    $screen=Test-Window '03-ResultsEntered' $id Evaluation '03-ResultsEntered'
    Check ($screen.ratingCount -eq 21 -and ($screen.rankSourceIds -join '|') -eq 'S01|S02|S03') 'WPF 선행기술조사와 AI 2개 방식 · 평가할 문헌 21건'
    for ($i=0;$i -lt $people.Count;$i++) {
        $person=$people[$i]
        $screen=Test-Window '04-Evaluated' $person.userId Evaluation ('04-Examiner-{0:00}' -f ($i+1))
        Check ($screen.caseIds.Count -eq 2 -and $screen.ratingCount -eq 21 -and $screen.ratingsHeight -ge 300 -and $screen.ratingsHeight -gt $screen.rankHeight -and $screen.deliveryVisible -eq [bool]$person.sourceId) ($person.name+' WPF 본인 평가·역할별 메뉴·문헌 목록 높이')
        if ($person.sourceId) {
            $screen=Test-Window '04-Evaluated' $person.userId AI ('05-AiHandler-{0:00}' -f ($i+1))
            Check ($screen.workspace -eq 'AI' -and $screen.caseIds.Count -eq 10 -and $screen.documentCount -eq 7 -and ($screen.sourceIds -join '|') -eq $person.sourceId) ($person.name+' WPF 담당 그룹 10건·지정 AI·문헌 7건')
        }
    }
    $screen=Test-Window '04-Evaluated' $admin.User.userId Statistics '06-Administrator' -Admin
    Check ($screen.statisticsSummary.caseCount -eq 20 -and $screen.statisticsSummary.completeCount -eq 20) 'WPF 관리자 전체 통계 20건 완료'
}
$report=[ordered]@{period=$Period;notice='모든 이름·출원번호·선행기술 문헌·평가는 테스트용 가상 데이터입니다. 실제 AI 검색을 실행하지 않았습니다.';passed=$true;checkCount=$checks.Count;guiWindowCount=$gui.Count;participants=@($people.ToArray());stages=@($stages.ToArray());summary=$view.summary;groups=$view.groups;accounts=$view.accounts;sources=$view.sources;checks=@($checks.ToArray());screenshots=@($gui.ToArray())}
[IO.File]::WriteAllText((Join-Path $output 'Report.json'),(ConvertTo-Json $report -Depth 15),$utf8)
$lines=@('전체 업무 흐름 테스트 결과',('운영월: '+$Period),('검증: '+$checks.Count+'개 통과 / WPF 화면 '+$gui.Count+'개'),'심사관 10명(두 그룹 각 5명), AI 담당자 4명(그룹별 2명)','대상 건 20건 / AI 검색 결과 280문헌 / 선행기술조사 140문헌 / 문헌 평가 420건','모든 데이터는 테스트용 가상 자료이며 실제 AI 검색을 수행하지 않았습니다.','','각 단계 폴더의 User.cmd에서 심사과 - 이름을 선택하세요. Admin.cmd는 관리자 통계 화면을 엽니다.','단계별 데이터는 독립된 복사본입니다. 앞 단계에서 입력해도 다른 단계는 바뀌지 않습니다.','프로그램 위치를 옮겼다면 이 도구를 다시 실행해 실행 파일을 새로 만드세요.','')
$lines+=@($people | ForEach-Object {('{0} | {1} - {2} | {3}' -f $_.groupId,$_.division,$_.name,$(if ($_.sourceId -eq 'S02') {'ChatGPT 담당자 겸 심사관'} elseif ($_.sourceId -eq 'S03') {'클로드 담당자 겸 심사관'} else {'심사관'}))})
[IO.File]::WriteAllLines((Join-Path $output '안내.txt'),[string[]]$lines,$utf8)
Write-Output "PASS full flow: $($checks.Count) checks, $($gui.Count) WPF windows"
Write-Output "Report: $output"
