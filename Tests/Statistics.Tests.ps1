$ErrorActionPreference='Stop'
Set-StrictMode -Version 2.0
Import-Module (Join-Path $PSScriptRoot '../App/Modules/Workflow.psm1') -Force -DisableNameChecking
$taskRoot=Join-Path ([IO.Path]::GetTempPath()) ('AIReview-Statistics-'+[guid]::NewGuid().ToString('N'))
function Assert-Stat([bool]$Value,[string]$Message) {if (-not $Value) {throw $Message}}
function Request {return [guid]::NewGuid().ToString('N')}
function Document([string]$Id) {return [pscustomobject]@{documentId=$Id;identifier=$Id;type='특허';attachments=@()}}
$admin=Initialize-Demo $taskRoot;$period=Get-Date -Format 'yyyy-MM'
$examiner=Get-Context $taskRoot -UserId U002 -AccessMode User
$snapshot=Get-AdminStatistics $admin $period;$view=Select-AdminStatistics $snapshot
Assert-Stat ($view.summary.caseCount -eq 0 -and $view.summary.participantCount -eq 3 -and $view.summary.unselectedCount -eq 3 -and $view.summary.evaluationRate -eq '—') 'Empty month must show all zero-submission participants and no fake completion rate.'
$denied=$false;try {Get-AdminStatistics $examiner $period | Out-Null} catch {$denied=$_.Exception.Message -like '*관리자*'}
Assert-Stat $denied 'Ordinary users must not load administrator statistics.'
Write-Output 'PASS empty month, zero-submission participants and admin-only access'

[void](Save-MasterRecord $admin -Kind Groups -Id G02 -ExpectedVersion 0 -RequestId (Request) -Data ([pscustomobject]@{groupId='G02';name='두 번째 그룹';active=$true}))
$other=New-User $admin -Name '다른 그룹 심사관' -Division '통계과' -Roles @('examiner','ai') -RequestId (Request)
$unused=New-User $admin -Name '미참여 계정' -Roles @('examiner') -Active $false -RequestId (Request)
$settings=Get-Period $admin $period
$settings.members+=@([pscustomobject]@{userId=$other.userId;groupId='G02'})
$settings.assignments+=@([pscustomobject]@{userId='U001';groupId='G02';sourceId='S02'},[pscustomobject]@{userId=$other.userId;groupId='G02';sourceId='S03'})
[void](Save-Period $admin $period -Data $settings -ExpectedVersion $settings.version -RequestId (Request))
$otherContext=Get-Context $taskRoot -UserId $other.userId
$ai=Get-Context $taskRoot -UserId U003
[void](Save-TargetSelection $ai $period -ExpectedVersion 0 -RequestId (Request) -Status Draft -Rows @([pscustomobject]@{applicationNumber='10-2026-0000099';cpc=''}))
$a=New-Submission $admin $period -ApplicationNumber '10-2026-0000001' -RequestId (Request)
$b=New-Submission $examiner $period -ApplicationNumber '10-2026-0000002' -RequestId (Request)
$c=New-Submission $otherContext $period -ApplicationNumber '10-2026-0000003' -RequestId (Request)
$cancelled=New-Submission $admin $period -ApplicationNumber '10-2026-0000004' -RequestId (Request)
[void](Cancel-Submission $admin $cancelled -ExpectedVersion $cancelled.version -RequestId (Request))
$versions=@{};$ratings=@();$rankings=@();$n=0
foreach ($source in @('S01','S02','S03')) {
    $n++;$saved=Save-Delivery $admin $a $source -ExpectedVersion 0 -RequestId (Request) -Status Published -Documents @((Document "D$n"))
    $versions[$source]=$saved.version
    $ratings+=[pscustomobject]@{sourceId=$source;documentId="D$n";score=$(if ($n -eq 1) {0} elseif ($n -eq 2) {5} else {$null});cited=$(if ($n -eq 3) {$null} else {$n -eq 2});note=''}
    $rankings+=[pscustomobject]@{sourceId=$source;status='Ranked';rank=$(if ($n -eq 1) {2} elseif ($n -eq 2) {1} else {3})}
}
[void](Save-Delivery $admin $b S02 -ExpectedVersion 0 -RequestId (Request) -Status Draft -Documents @((Document 'Draft')))
[void](Save-Delivery $admin $b S03 -ExpectedVersion 0 -RequestId (Request) -Status Published -Documents @((Document 'B3')))
$evaluation=Save-Evaluation $admin $a -ExpectedVersion 0 -RequestId (Request) -DeliveryVersions $versions -DocumentRatings $ratings -Rankings $rankings
$snapshot=Get-AdminStatistics $admin $period;$view=Select-AdminStatistics $snapshot
Assert-Stat ($view.summary.caseCount -eq 3 -and $snapshot.cancelledCount -eq 1 -and $view.summary.deliveredCount -eq 1 -and $view.summary.waitingCount -eq 2 -and $view.summary.partialCount -eq 1) 'Draft deliveries and cancelled cases must not count as completed work.'
Assert-Stat (@($view.accounts | Where-Object {$_.userId -eq 'U003'})[0].selectionStatus -eq '임시저장') 'A saved selection draft must remain distinguishable from a confirmed selection.'
Assert-Stat (@($view.accounts | Where-Object {$_.userId -eq $unused.userId})[0].selectionStatus -eq '미참여') 'Nonparticipating inactive accounts must remain visible.'
$s1=@($view.sources | Where-Object {$_.sourceId -eq 'S01'})[0];$s3=@($view.sources | Where-Object {$_.sourceId -eq 'S03'})[0]
Assert-Stat ($s1.averageScore -eq '0.00' -and $s1.scoreCount -eq 1 -and $s1.citationRate -eq '0%' -and $s3.scoreCount -eq 0 -and $s3.citationRate -eq '—') 'A real zero/X must count, while null values must not enter either denominator.'
Write-Output 'PASS cancelled/draft exclusion, partial evaluation, zero/X and null denominators'

$ratings[2].score=4;$ratings[2].cited=$false
$evaluation=Save-Evaluation $admin $a -ExpectedVersion $evaluation.version -RequestId (Request) -DeliveryVersions $versions -DocumentRatings $ratings -Rankings $rankings
$snapshot=Get-AdminStatistics $admin $period;$view=Select-AdminStatistics $snapshot
Assert-Stat ($view.summary.completeCount -eq 1 -and $view.summary.evaluationRate -eq '33.3%' -and ($view.groups | Measure-Object caseCount -Sum).Sum -eq $view.summary.caseCount) 'Overall and grouped totals must agree.'
$g2=Select-AdminStatistics $snapshot -GroupId G02
Assert-Stat ($g2.summary.caseCount -eq 1 -and $g2.summary.participantCount -eq 1) 'Group filter must not count AI assignees as participants in the wrong group.'
$u1=Select-AdminStatistics $snapshot -UserId U001
Assert-Stat ($u1.cases.Count -eq 1 -and $u1.accounts[0].aiAssigned -eq 3 -and $u1.accounts[0].aiPublished -eq 1 -and $u1.accounts[0].aiPending -eq 2) 'Assigned AI work must remain separate from the account own submissions.'
$g2u1=Select-AdminStatistics $snapshot -GroupId G02 -UserId U001
Assert-Stat ($g2u1.summary.participantCount -eq 0 -and $g2u1.summary.caseCount -eq 0 -and $g2u1.accounts[0].aiAssigned -eq 1) 'Cross-group AI assignments must not become personal submissions.'
Write-Output 'PASS overall/group/account drilldown, complete rate and separate AI responsibilities'

[void](Save-Delivery $admin $a S02 -ExpectedVersion 1 -RequestId (Request) -Status Published -Documents @((Document 'New2')))
[void](Save-Delivery $admin $a S03 -ExpectedVersion 1 -RequestId (Request) -Status Draft -Documents @((Document 'NewDraft3')))
$snapshot=Get-AdminStatistics $admin $period;$view=Select-AdminStatistics $snapshot
$s2=@($view.sources | Where-Object {$_.sourceId -eq 'S02'})[0];$s3=@($view.sources | Where-Object {$_.sourceId -eq 'S03'})[0]
Assert-Stat ($view.summary.staleCount -eq 1 -and $view.summary.completeCount -eq 0 -and $view.summary.deliveredCount -eq 1 -and $s2.scoreCount -eq 0 -and $s2.firstCount -eq 0 -and $s3.scoreCount -eq 1) 'Stale scores/ranks must be excluded while a newer draft must not replace the published result.'
$emptyPeriod=(Get-Date).AddMonths(1).ToString('yyyy-MM');$settings=Get-Period $admin $period;$settings.period=$emptyPeriod
[void](Save-Period $admin $emptyPeriod -Data $settings -ExpectedVersion 0 -RequestId (Request))
$empty=Select-AdminStatistics (Get-AdminStatistics $admin $emptyPeriod)
Assert-Stat ($empty.summary.caseCount -eq 0 -and $empty.summary.unselectedCount -eq 4) 'Operating months must not share selection/completion counts.'
Write-Output 'PASS stale evaluation, published-with-draft handling and operating-month isolation'

$broken=Join-Path $taskRoot ('Data/'+$period+'/G01/Deliveries/S02/U001/'+$a.caseId+'/revisions/000003.json')
[IO.File]::WriteAllText($broken,'broken',(New-Object Text.UTF8Encoding($false)))
$failed=$false;try {Get-AdminStatistics $admin $period | Out-Null} catch {$failed=$true}
Assert-Stat $failed 'Corrupt records must fail the snapshot instead of looking like zero or complete statistics.'
Write-Output 'PASS corrupt input is reported instead of silently entering statistics'
