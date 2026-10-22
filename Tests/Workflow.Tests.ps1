param([switch]$KeepArtifacts)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0
$modulePath = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../App/Modules/Workflow.psm1'))
Import-Module $modulePath -Force -DisableNameChecking
$testRoot = Join-Path ([IO.Path]::GetTempPath()) ('AiSearch-WorkflowTests-' + [guid]::NewGuid().ToString('N'))
$script:passed = 0

function Assert-True { param([bool]$Condition,[string]$Message) if (-not $Condition) { throw "Assertion failed: $Message" } }
function Assert-Throws {
    param([scriptblock]$Action,[string]$Pattern='.')
    $caught=$null
    try { & $Action | Out-Null } catch { $caught=$_ }
    if ($null -eq $caught) { throw "Expected exception: $Pattern" }
    if ($caught.Exception.Message -notmatch $Pattern) { throw "Expected '$Pattern', received '$($caught.Exception.Message)'" }
}
function Test-Case { param([string]$Name,[scriptblock]$Action) & $Action; $script:passed++; Write-Host "PASS $Name" }
function New-Request { return [guid]::NewGuid().ToString('N') }
function New-Document { param([string]$Identifier) return [pscustomobject]@{type='특허';identifier=$Identifier;title='검증 문헌';description='관련 구성';attachments=@()} }

try {
    $admin=Initialize-Demo -Root $testRoot
    $examiner=Get-Context -Root $testRoot -UserId U002
    $ai=Get-Context -Root $testRoot -UserId U003
    $month=Get-Date -Format yyyy-MM

    Test-Case 'demo initialization is idempotent and identities are separate' {
        $null=Initialize-Demo -Root $testRoot
        Assert-True ((Get-Users $admin).Count -eq 3) 'three demo users'
        Assert-True ((Get-Period $admin $month).version -eq 1) 'setup never rewrites existing month'
        Assert-True ($examiner.user.userId -eq 'U002') 'explicit demo identity'
    }

    $submissionRequest=New-Request
    $case=New-Submission -Context $examiner -Period $month -ApplicationNumber 1020260000001 -RequestId $submissionRequest
    Test-Case 'submission normalization, duplicate protection and idempotent retry' {
        Assert-True ($case.applicationNumber -eq '10-2026-0000001') 'normalize 13 digits'
        $again=New-Submission -Context $examiner -Period $month -ApplicationNumber '10-2026-0000001' -RequestId $submissionRequest
        Assert-True ($again.caseId -eq $case.caseId) 'same request returns same case'
        Assert-Throws {New-Submission $examiner $month '10-2026-0000001' (New-Request)} '이미 제출'
        Assert-Throws {New-Submission $examiner $month '10-2026-0000002' $submissionRequest} '요청 ID'
        Assert-Throws {New-Submission $examiner $month 'bad' (New-Request)} '출원번호'
    }

    Test-Case 'role and assignment boundaries use authoritative records' {
        Assert-Throws {Get-Cases $examiner $month All} '관리자'
        Assert-Throws {Save-Delivery -Context $ai -Case $case -SourceId S02 -ExpectedVersion 0 -RequestId (New-Request) -Status Published -Documents @((New-Document P1))} '권한'
        Assert-Throws {Save-Delivery -Context $examiner -Case $case -SourceId S03 -ExpectedVersion 0 -RequestId (New-Request) -Status Published -Documents @((New-Document P1))} '권한'
        Assert-Throws {Get-Evaluation $ai $case} '권한'
        Assert-True (@(Get-Cases $ai $month Assigned).Count -eq 1) 'AI sees assigned group cases'
    }

    $docs=@((New-Document P1))
    $publishedRequest=New-Request
    $s02=Save-Delivery -Context $admin -Case $case -SourceId S02 -ExpectedVersion 0 -RequestId $publishedRequest -Status Published -Documents $docs
    $s03=Save-Delivery -Context $ai -Case $case -SourceId S03 -ExpectedVersion 0 -RequestId (New-Request) -Status Published -Documents @((New-Document P2))
    $s01=Save-Delivery -Context $examiner -Case $case -SourceId S01 -ExpectedVersion 0 -RequestId (New-Request) -Status Published -NoResults $true -NoResultsReason '관련 문헌 없음'

    Test-Case 'publication idempotency compares exact request content' {
        $again=Save-Delivery -Context $admin -Case $case -SourceId S02 -ExpectedVersion 0 -RequestId $publishedRequest -Status Published -Documents $docs
        Assert-True ($again.version -eq 1) 'retry does not create another version or another document ID'
        $changed=@((New-Document Different))
        Assert-Throws {Save-Delivery -Context $admin -Case $case -SourceId S02 -ExpectedVersion 0 -RequestId $publishedRequest -Status Published -Documents $changed} '요청 ID'
        Assert-True (@(Get-Deliveries $ai $case).Count -eq 1) 'nonowner AI sees only assigned source'
    }

    Test-Case 'seven-document cap and no-results validation' {
        $tooMany=@(1..8 | ForEach-Object {New-Document ('P'+$_)})
        Assert-Throws {Save-Delivery -Context $admin -Case $case -SourceId S02 -ExpectedVersion 1 -RequestId (New-Request) -Status Draft -Documents $tooMany} '최대 7'
        Assert-Throws {Save-Delivery -Context $admin -Case $case -SourceId S02 -ExpectedVersion 1 -RequestId (New-Request) -Status Draft -NoResults $true} '사유'
        Assert-Throws {Save-Delivery -Context $admin -Case $case -SourceId S02 -ExpectedVersion 1 -RequestId (New-Request) -Status Draft -NoResults $true -NoResultsReason why -Documents $docs} '동시에'
    }

    $versions=@{S01=1;S02=1;S03=1}
    $ratings=@(
        [pscustomobject]@{sourceId='S02';documentId=$s02.data.documents[0].documentId;score=0;cited=$false;note='실제 0점'},
        [pscustomobject]@{sourceId='S03';documentId=$s03.data.documents[0].documentId;score=$null;cited=$null;note='미평가'}
    )
    $ranks=@(@{sourceId='S01';rank=$null;status='NoResults'},@{sourceId='S02';rank=1;status='Ranked'},@{sourceId='S03';rank=2;status='Ranked'})
    $evaluationRequest=New-Request
    $eval=Save-Evaluation -Context $examiner -Case $case -ExpectedVersion 0 -RequestId $evaluationRequest -DeliveryVersions $versions -DocumentRatings $ratings -Rankings $ranks

    Test-Case 'zero and false differ from unselected evaluation values' {
        Assert-True ($eval.data.documentRatings[0].score -eq 0 -and $null -ne $eval.data.documentRatings[0].score) 'zero survives JSON'
        Assert-True ($eval.data.documentRatings[0].cited -eq $false -and $null -ne $eval.data.documentRatings[0].cited) 'false survives JSON'
        Assert-True ($null -eq $eval.data.documentRatings[1].score -and $null -eq $eval.data.documentRatings[1].cited) 'unselected survives as null'
        Assert-True (-not $eval.data.complete) 'partial input is not completed'
    }

    Test-Case 'draft remains private while examiner sees previous published version' {
        $draftDocs=@($s02.data.documents)
        $draftDocs[0].title='수정 초안'
        $draft=Save-Delivery -Context $admin -Case $case -SourceId S02 -ExpectedVersion 1 -RequestId (New-Request) -Status Draft -Documents $draftDocs
        Assert-True ($draft.version -eq 2) 'draft revision appended'
        $visible=@(Get-Deliveries $examiner $case | Where-Object {$_.data.sourceId -eq 'S02'})[0]
        Assert-True ($visible.version -eq 1 -and $visible.data.documents[0].title -eq '검증 문헌') 'examiner still sees publication'
        Assert-True ((Get-DeliveryForEdit $admin $case S02).version -eq 2) 'assigned editor sees draft'
        Assert-True (-not (Get-Evaluation $examiner $case).needsReconfirmation) 'draft does not invalidate evaluation'
        Assert-Throws {Get-DeliveryForEdit $examiner $case S02} '권한'
        Assert-Throws {Save-Delivery -Context $admin -Case $case -SourceId S02 -ExpectedVersion 1 -RequestId (New-Request) -Status Draft -Documents $docs} '변경'
    }

    $newDocs=@((Get-DeliveryForEdit $admin $case S02).data.documents)
    $s02v3=Save-Delivery -Context $admin -Case $case -SourceId S02 -ExpectedVersion 2 -RequestId (New-Request) -Status Published -Documents $newDocs -Reason '문헌 설명 정정'
    Test-Case 'republishing preserves IDs and invalidates old evaluation snapshot' {
        Assert-True ($s02v3.data.documents[0].documentId -eq $s02.data.documents[0].documentId) 'document identity remains stable'
        $old=Get-Evaluation $examiner $case
        Assert-True ($old.needsReconfirmation -and $old.staleSourceIds -contains 'S02') 'stale publication detected'
        Assert-Throws {Save-Evaluation -Context $examiner -Case $case -ExpectedVersion 1 -RequestId (New-Request) -DeliveryVersions $versions -DocumentRatings $ratings -Rankings $ranks} '납품자료가 변경'
        $again=Save-Evaluation -Context $examiner -Case $case -ExpectedVersion 0 -RequestId $evaluationRequest -DeliveryVersions $versions -DocumentRatings $ratings -Rankings $ranks
        Assert-True ($again.version -eq 1) 'idempotent old evaluation returns its original revision'
        $changedRatings=@($ratings | ForEach-Object {$_ | Select-Object *})
        $changedRatings[0].note='다른 내용'
        Assert-Throws {Save-Evaluation -Context $examiner -Case $case -ExpectedVersion 0 -RequestId $evaluationRequest -DeliveryVersions $versions -DocumentRatings $changedRatings -Rankings $ranks} '요청 ID'
    }

    Test-Case 'ranking and individual rating validation reject inconsistent values' {
        $currentVersions=@{S01=1;S02=3;S03=1}
        $badRanks=@(@{sourceId='S01';rank=$null;status='NoResults'},@{sourceId='S02';rank=1;status='Ranked'},@{sourceId='S03';rank=1;status='Ranked'})
        Assert-Throws {Save-Evaluation -Context $examiner -Case $case -ExpectedVersion 1 -RequestId (New-Request) -DeliveryVersions $currentVersions -DocumentRatings $ratings -Rankings $badRanks} '동순위'
        $bad=@([pscustomobject]@{sourceId='S02';documentId=$s02.data.documents[0].documentId;score=6;cited=$true;note=''})
        Assert-Throws {Save-Evaluation -Context $examiner -Case $case -ExpectedVersion 1 -RequestId (New-Request) -DeliveryVersions $currentVersions -DocumentRatings $bad -Rankings $ranks} '0~5'
        $bad[0].score=1;$bad[0].cited='X'
        Assert-Throws {Save-Evaluation -Context $examiner -Case $case -ExpectedVersion 1 -RequestId (New-Request) -DeliveryVersions $currentVersions -DocumentRatings $bad -Rankings $ranks} '인용 여부'
        $ratings[1].score=5;$ratings[1].cited=$true
        $done=Save-Evaluation -Context $examiner -Case $case -ExpectedVersion 1 -RequestId (New-Request) -DeliveryVersions $currentVersions -DocumentRatings $ratings -Rankings $ranks
        Assert-True ($done.data.complete -and -not $done.needsReconfirmation) 'completed snapshot matches publications'
    }

    Test-Case 'attachment paths are restricted to owning case and source' {
        $file=Join-Path $testRoot 'attachment.txt';[IO.File]::WriteAllText($file,'attachment test')
        $attachment=Import-DeliveryAttachment -Context $ai -Case $case -SourceId S03 -Path $file
        $resolved=Resolve-DeliveryAttachment $ai $case S03 $attachment
        Assert-True ([IO.File]::Exists($resolved)) 'copied attachment exists'
        Assert-True ($attachment.fileName -eq 'attachment.txt' -and $attachment.sha256.Length -eq 64) 'metadata preserved'
        Assert-True ([IO.File]::ReadAllText($resolved) -eq 'attachment test') 'original attachment bytes copied'
        Assert-True ((Get-FileHash -LiteralPath $file -Algorithm SHA256).Hash.ToLowerInvariant() -eq $attachment.sha256) 'attachment checksum matches selected input'
        Assert-Throws {Resolve-DeliveryAttachment $ai $case S03 @{relativePath='..\outside.txt'}} '벗어났'
        Assert-Throws {Resolve-DeliveryAttachment $ai $case S02 $attachment} '벗어났'
    }

    Test-Case 'closed month blocks writes, and explicit reopen restores them' {
        $closeRequest=New-Request
        $closed=Set-PeriodStatus -Context $admin -Period $month -Status Closed -ExpectedVersion 1 -RequestId $closeRequest -Reason '월 마감'
        Assert-Throws {Save-Delivery -Context $ai -Case $case -SourceId S03 -ExpectedVersion 1 -RequestId (New-Request) -Status Draft -Documents @($s03.data.documents)} '마감'
        Assert-Throws {New-Submission $examiner $month '10-2026-0000002' (New-Request)} '마감'
        Assert-True ((Set-PeriodStatus -Context $admin -Period $month -Status Closed -ExpectedVersion 1 -RequestId $closeRequest -Reason '월 마감').version -eq 2) 'close retry uses original revision'
        Assert-Throws {Set-PeriodStatus -Context $admin -Period $month -Status Open -ExpectedVersion 1 -RequestId $closeRequest -Reason '월 마감'} '요청 ID'
        $opened=Set-PeriodStatus -Context $admin -Period $month -Status Open -ExpectedVersion 2 -RequestId (New-Request) -Reason '정정 재개방'
        Assert-True ($opened.status -eq 'Open') 'month reopened'
    }

    Test-Case 'cancellation preserves original submission and blocks subsequent writes' {
        $extra=New-Submission $examiner $month '10-2026-0000002' (New-Request)
        $cancelRequest=New-Request
        $cancelled=Cancel-Submission -Context $examiner -Case $extra -ExpectedVersion 1 -RequestId $cancelRequest -Reason '번호 정정'
        Assert-True ($cancelled.status -eq 'Cancelled' -and $cancelled.version -eq 2) 'cancel appends revision'
        Assert-True ((Cancel-Submission -Context $examiner -Case $extra -ExpectedVersion 1 -RequestId $cancelRequest -Reason '번호 정정').version -eq 2) 'cancel retry'
        Assert-Throws {Cancel-Submission -Context $examiner -Case $extra -ExpectedVersion 1 -RequestId $cancelRequest -Reason '다른 사유'} '요청 ID'
        Assert-Throws {Save-Delivery -Context $examiner -Case $extra -SourceId S01 -ExpectedVersion 0 -RequestId (New-Request) -Status Published -NoResults $true -NoResultsReason none} '취소'
    }

    Test-Case 'month closure waits for shared save handles while shared handles coexist' {
        $gatePath=Join-Path $testRoot ('Locks\'+$month+'\lifecycle.lock')
        $first=New-Object IO.FileStream($gatePath,[IO.FileMode]::Open,[IO.FileAccess]::Read,[IO.FileShare]::Read)
        $second=New-Object IO.FileStream($gatePath,[IO.FileMode]::Open,[IO.FileAccess]::Read,[IO.FileShare]::Read)
        $ready=Join-Path $testRoot 'close-ready.txt'
        $moduleLiteral=$modulePath.Replace("'","''");$rootLiteral=$testRoot.Replace("'","''");$readyLiteral=$ready.Replace("'","''")
        $childCode="`$ErrorActionPreference='Stop';Import-Module '$moduleLiteral' -Force -DisableNameChecking;`$ctx=Get-Context '$rootLiteral';`$p=Get-Period `$ctx '$month';[IO.File]::WriteAllText('$readyLiteral','ready');Set-PeriodStatus -Context `$ctx -Period '$month' -Status Closed -ExpectedVersion `$p.version -RequestId ([guid]::NewGuid().ToString('N')) -Reason 'concurrency verification'|Out-Null"
        $info=New-Object Diagnostics.ProcessStartInfo
        $info.FileName=[Diagnostics.Process]::GetCurrentProcess().MainModule.FileName
        $info.Arguments='-NoProfile -NonInteractive -ExecutionPolicy Bypass -EncodedCommand '+[Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($childCode))
        $info.UseShellExecute=$false;$info.CreateNoWindow=$true;$info.RedirectStandardError=$true;$info.RedirectStandardOutput=$true
        $child=New-Object Diagnostics.Process;$child.StartInfo=$info
        try {
            $null=$child.Start();$clock=[Diagnostics.Stopwatch]::StartNew()
            while (-not [IO.File]::Exists($ready) -and -not $child.HasExited -and $clock.ElapsedMilliseconds -lt 10000) {Start-Sleep -Milliseconds 20}
            Assert-True ([IO.File]::Exists($ready)) 'child reached month close'
            Start-Sleep -Milliseconds 120
            Assert-True (-not $child.HasExited) 'exclusive month close waits for shared save gate'
            $first.Dispose();$first=$null;$second.Dispose();$second=$null
            if (-not $child.WaitForExit(10000)) {$child.Kill();throw 'Month-close child did not finish.'}
            $childError=$child.StandardError.ReadToEnd()
            Assert-True ($child.ExitCode -eq 0) ('child closes after shared saves finish: '+$childError)
            $period=Get-Period $admin $month
            Assert-True ($period.status -eq 'Closed') 'exclusive close committed'
            $null=Set-PeriodStatus -Context $admin -Period $month -Status Open -ExpectedVersion $period.version -RequestId (New-Request) -Reason 'test reopen'
        } finally {
            if ($first) {$first.Dispose()};if ($second) {$second.Dispose()}
            if (-not $child.HasExited) {$child.Kill();$child.WaitForExit()};$child.Dispose()
        }
    }

    Test-Case 'admin period edits preserve source snapshots and require lifecycle reasons' {
        $period=Get-Period $admin $month
        $bad=$period | Select-Object *
        $bad.sourceIds=@('S01','S02');$bad.assignments=@($bad.assignments | Where-Object {$_.sourceId -ne 'S03'})
        Assert-Throws {Save-Period -Context $admin -Period $month -Data $bad -ExpectedVersion $period.version -RequestId (New-Request) -Reason 'remove source'} '제거할 수 없'
        $bad=$period | Select-Object *;$bad.status='Closed'
        Assert-Throws {Save-Period -Context $admin -Period $month -Data $bad -ExpectedVersion $period.version -RequestId (New-Request)} '사유'
        $period.submissionLimit=3
        $request=New-Request
        $saved=Save-Period -Context $admin -Period $month -Data $period -ExpectedVersion $period.version -RequestId $request -Reason '상한 설정'
        Assert-True ($saved.submissionLimit -eq 3) 'period definitions persist through lifecycle gate'
        $again=Save-Period -Context $admin -Period $month -Data $period -ExpectedVersion $period.version -RequestId $request -Reason '상한 설정'
        Assert-True ($again.version -eq $saved.version) 'master period save is idempotent'
    }

    Test-Case 'inactive sources retain readable history while refusing new writes' {
        $source=@(Get-Sources $admin | Where-Object {$_.sourceId -eq 'S02'})[0]
        $source.active=$false
        $inactive=Save-MasterRecord -Context $admin -Kind Sources -Id S02 -Data $source -ExpectedVersion $source.version -RequestId (New-Request) -Reason '사용 중지'
        Assert-True (@(Get-DeliveryHistory $admin $case S02).Count -eq 3) 'historical versions remain readable'
        Assert-Throws {Save-Delivery -Context $admin -Case $case -SourceId S02 -ExpectedVersion 3 -RequestId (New-Request) -Status Draft -Documents $newDocs} '활성 조사 방식'
        $inactive.active=$true
        $null=Save-MasterRecord -Context $admin -Kind Sources -Id S02 -Data $inactive -ExpectedVersion $inactive.version -RequestId (New-Request) -Reason '사용 재개'
    }

    Test-Case 'revoked role and inactive user immediately invalidate cached context' {
        $record=@(Get-Users $admin | Where-Object {$_.userId -eq 'U003'})[0]
        $record.roles=@('examiner')
        $changed=Save-MasterRecord -Context $admin -Kind Users -Id U003 -Data $record -ExpectedVersion $record.version -RequestId (New-Request) -Reason '담당 해제'
        Assert-Throws {Save-Delivery -Context $ai -Case $case -SourceId S03 -ExpectedVersion 1 -RequestId (New-Request) -Status Draft -Documents @($s03.data.documents)} '권한'
        $changed.active=$false
        $null=Save-MasterRecord -Context $admin -Kind Users -Id U003 -Data $changed -ExpectedVersion $changed.version -RequestId (New-Request) -Reason '비활성화'
        Assert-Throws {Get-Cases $ai $month Assigned} '활성 사용자'
    }

    Test-Case 'production roots disallow demo user override and forged cached identity' {
        $marker=Join-Path $testRoot '.demo-mode.json';$renamed=Join-Path $testRoot '.demo-mode.disabled'
        [IO.File]::Move($marker,$renamed)
        try {
            Assert-Throws {Get-Context -Root $testRoot -UserId U002} '데모'
            Assert-Throws {Get-Cases $examiner $month Mine} 'Windows 계정'
        } finally {[IO.File]::Move($renamed,$marker)}
    }

    Write-Host "Workflow tests: $script:passed passed."
} finally {
    if ($KeepArtifacts) {Write-Host "Artifacts: $testRoot"}
    elseif ([IO.Directory]::Exists($testRoot)) {
        $resolvedRoot=[IO.Path]::GetFullPath($testRoot)
        $tempPrefix=[IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\')+'\'
        if (-not $resolvedRoot.StartsWith($tempPrefix,[StringComparison]::OrdinalIgnoreCase) -or [IO.Path]::GetFileName($resolvedRoot) -notmatch '^AiSearch-WorkflowTests-[0-9a-f]{32}$') {throw 'Unsafe test cleanup path.'}
        Remove-Item -LiteralPath $resolvedRoot -Recurse -Force
    }
}
