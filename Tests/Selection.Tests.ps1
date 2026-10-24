$ErrorActionPreference='Stop'
Set-StrictMode -Version 2.0
Import-Module (Join-Path $PSScriptRoot '../App/Modules/Workflow.psm1') -Force -DisableNameChecking
Import-Module (Join-Path $PSScriptRoot '../App/Modules/Storage.psm1') -Force -DisableNameChecking
$testRoot=Join-Path ([IO.Path]::GetTempPath()) ('AIReview-SelectionTests-'+[guid]::NewGuid().ToString('N'))
function Request { [guid]::NewGuid().ToString('N') }
function Check([bool]$Condition,[string]$Message) {if (-not $Condition) {throw $Message}}
function Reject([scriptblock]$Action,[string]$Pattern) {
    $failure=$null;try {& $Action | Out-Null} catch {$failure=$_}
    if (-not $failure -or $failure.Exception.Message -notmatch $Pattern) {throw "Expected $Pattern, received $failure"}
}
try {
    $admin=Initialize-Demo -Root $testRoot
    $user=Get-Context -Root $testRoot -UserId U002 -AccessMode User
    $ai=Get-Context -Root $testRoot -UserId U003 -AccessMode User
    $month=Get-Date -Format yyyy-MM
    Check ((Get-TargetSelection $user $month).status -eq 'NotStarted') 'New month must start at selection.'
    $rows=@([pscustomobject]@{applicationNumber='unfinished';cpc='G06F 16/00'})
    $draftRequest=Request
    $draft=Save-TargetSelection $user $month 0 $draftRequest Draft $rows
    Check ($draft.version -eq 1 -and $draft.status -eq 'Draft') 'Draft revision missing.'
    $fresh=Get-Context -Root $testRoot -UserId U002 -AccessMode User
    Check ((Get-TargetSelection $fresh $month).rows[0].cpc -eq 'G06F 16/00') 'Draft not restored from disk.'
    Check (@(Get-Cases $user $month Mine).Count -eq 0 -and @(Get-Cases $ai $month Assigned).Count -eq 0 -and @(Get-Cases $admin $month All).Count -eq 0) 'Draft leaked into a case list.'
    Check ((Save-TargetSelection $user $month 0 $draftRequest Draft $rows).version -eq 1) 'Draft retry created a duplicate revision.'
    Reject {Save-TargetSelection $user $month 0 (Request) Draft @()} '다른 창'
    Reject {Save-TargetSelection $user $month 1 (Request) Confirmed $rows} '출원번호'
    Reject {Save-TargetSelection $user $month 1 (Request) Confirmed @()} '한 건 이상'
    Reject {Save-TargetSelection $user $month 0 $draftRequest Draft @()} '요청 ID'
    $duplicates=@([pscustomobject]@{applicationNumber='1020260000001';cpc=''},[pscustomobject]@{applicationNumber='10-2026-0000001';cpc=''})
    Reject {Save-TargetSelection $user $month 1 (Request) Confirmed $duplicates} '중복'
    Check (@(Get-Cases $admin $month All).Count -eq 0) 'Failed validation published a partial set.'
    Write-Output 'PASS draft persistence, privacy, stale version, idempotency and whole-list validation'

    $settings=Get-Period $admin $month
    $next='2027-01';$nextSettings=$settings | Select-Object *;$nextSettings.period=$next
    [void](Save-Period $admin $next $nextSettings 0 (Request))
    Check ((Get-TargetSelection $user $next).status -eq 'NotStarted') 'Drafts crossed months.'
    $nextDraft=Save-TargetSelection $user $next 0 (Request) Draft @([pscustomobject]@{applicationNumber='10-2027-0000001';cpc='H04L'})
    Check ((Get-TargetSelection $user $month).rows[0].applicationNumber -eq 'unfinished') 'Second month replaced first draft.'
    $limited=$settings | Select-Object *;$limited.submissionLimit=1
    $limited=Save-Period $admin $month $limited $settings.version (Request)
    $rows=@([pscustomobject]@{applicationNumber='1020260000001';cpc='G06F 16/00'},[pscustomobject]@{applicationNumber='10-2026-0000002';cpc='H04L 9/00'})
    Reject {Save-TargetSelection $user $month 1 (Request) Confirmed $rows} '상한'
    $limited.submissionLimit=0
    [void](Save-Period $admin $month $limited $limited.version (Request))
    Write-Output 'PASS month isolation and submission limits'

    # Hold only the final manifest lock. Preparation succeeds, but the publish commit fails.
    $confirmRequest=Request
    $lockPath=Join-Path $draft.path write.lock
    $handle=[IO.File]::Open($lockPath,[IO.FileMode]::OpenOrCreate,[IO.FileAccess]::ReadWrite,[IO.FileShare]::None)
    try {Reject {Save-TargetSelection $user $month 1 $confirmRequest Confirmed $rows} 'Timed out'} finally {$handle.Dispose()}
    $prepared=@(Get-ChildItem -LiteralPath (Join-Path $testRoot "Data/$month/G01/Submissions/U002") -Directory)
    Check ($prepared.Count -eq 2) 'Expected two prepared immutable case records.'
    Check (@(Get-Cases $admin $month All).Count -eq 0 -and @(Get-Cases $ai $month Assigned).Count -eq 0) 'Interrupted confirmation published part of the set.'
    $hidden=(Read-Revision -Path $prepared[0].FullName).data
    Reject {Get-Deliveries $user $hidden} '확정되지'
    $confirmed=Save-TargetSelection $user $month 1 $confirmRequest Confirmed $rows
    Check ($confirmed.status -eq 'Confirmed' -and $confirmed.caseIds.Count -eq 2 -and $confirmed.version -eq 2) 'Confirmation failed.'
    $cases=@(Get-Cases $user $month Mine)
    Check ($cases.Count -eq 2 -and @(Get-Cases $ai $month Assigned).Count -eq 2) 'Confirmed cases not published together.'
    Check (@($cases | Where-Object {$_.applicationNumber -eq '10-2026-0000001' -and $_.cpc -eq 'G06F 16/00'}).Count -eq 1) 'Confirmation lost normalized number or CPC.'
    Check ((Save-TargetSelection $user $month 1 $confirmRequest Confirmed $rows).version -eq 2) 'Confirmation retry duplicated cases.'
    Reject {Save-TargetSelection $user $month 2 (Request) Draft @()} '이미 대상 건'
    Reject {Save-TargetSelection $user $month 1 (Request) Confirmed $rows} '다른 창'
    Write-Output 'PASS interrupted confirmation hides all prepared cases; same request recovers exactly once'

    $legacy=New-Submission $admin $month '10-2026-0000999' (Request)
    Check ((Get-TargetSelection $admin $month).status -eq 'Confirmed') 'Existing submissions must enter evaluation directly.'
    $other=Save-TargetSelection $ai $month 0 (Request) Confirmed @([pscustomobject]@{applicationNumber='10-2026-0000999';cpc=''})
    Check ((@(Get-Cases $ai $month Mine)[0]).duplicateWarning) 'Cross-examiner duplicate warning lost.'
    $case=$cases[0]
    Check ((Get-CaseStatus $user $case).deliveryStatusLabel -eq '대기 중') 'Missing AI results should wait.'
    $doc=[pscustomobject]@{documentId='D1';identifier='10-2025-0000001';type='특허';attachments=@()}
    [void](Save-Delivery -Context $admin -Case $case -SourceId S02 -ExpectedVersion 0 -RequestId (Request) -Status Published -Documents @($doc))
    Check ((Get-CaseStatus $user $case).deliveryStatusLabel -eq '대기 중') 'Partial AI delivery should wait.'
    [void](Save-Delivery -Context $admin -Case $case -SourceId S03 -ExpectedVersion 0 -RequestId (Request) -Status Draft -Documents @($doc))
    Check ((Get-CaseStatus $user $case).deliveryStatusLabel -eq '대기 중') 'Draft AI delivery should wait.'
    [void](Save-Delivery -Context $admin -Case $case -SourceId S03 -ExpectedVersion 1 -RequestId (Request) -Status Published -NoResults $true -NoResultsReason '검색 완료, 적합 문헌 없음')
    Check ((Get-CaseStatus $user $case).deliveryStatusLabel -eq '입력 완료') 'Completed no-results delivery must count as delivered.'
    [void](Save-Delivery -Context $admin -Case $case -SourceId S02 -ExpectedVersion 1 -RequestId (Request) -Status Draft -Documents @($doc))
    Check ((Get-CaseStatus $user $case).deliveryStatusLabel -eq '입력 완료') 'A new draft must not replace the published delivery state.'
    $nextSettings=Get-Period $admin $next
    [void](Set-PeriodStatus $admin $next Closed $nextSettings.version (Request) '검증 마감')
    Reject {Save-TargetSelection $user $next 1 (Request) Draft @()} '마감'
    Reject {Save-TargetSelection $user $next 1 (Request) Confirmed $nextDraft.rows} '마감'
    Write-Output 'PASS legacy submissions, duplicate warnings, waiting/completed delivery labels and month closure'
} finally {
    $resolved=[IO.Path]::GetFullPath($testRoot)
    $prefix=[IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\')+'\'
    if (-not $resolved.StartsWith($prefix,[StringComparison]::OrdinalIgnoreCase) -or [IO.Path]::GetFileName($resolved) -notmatch '^AIReview-SelectionTests-[a-f0-9]{32}$') {throw 'Unsafe test cleanup path.'}
    if ([IO.Directory]::Exists($resolved)) {Remove-Item -LiteralPath $resolved -Recurse -Force}
}
