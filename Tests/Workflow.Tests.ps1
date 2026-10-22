param([switch]$KeepArtifacts)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0
$modulePath = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../App/Modules/Workflow.psm1'))
Import-Module $modulePath -Force -DisableNameChecking
Import-Module (Join-Path $PSScriptRoot '../App/Modules/Storage.psm1') -Force -DisableNameChecking
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
        Assert-True ($examiner.user.bureau -eq '' -and $examiner.user.division -eq '') 'new demo user has blank affiliation'
    }

    Test-Case 'user affiliation is saved, trimmed and editable with validation' {
        $profile=[pscustomobject]@{userId='U004';sid='DEMO-U004';name='등록 검증';bureau=' 전기통신심사국 ';division=' 통신심사과 ';roles=@('examiner');active=$true}
        $saved=Save-MasterRecord -Context $admin -Type Users -Id U004 -Data $profile -ExpectedVersion 0 -RequestId (New-Request)
        Assert-True ($saved.bureau -eq '전기통신심사국' -and $saved.division -eq '통신심사과') 'new user persists trimmed affiliation'
        $saved.division='전자심사과'
        $updated=Save-MasterRecord -Context $admin -Type Users -Id U004 -Data $saved -ExpectedVersion $saved.version -RequestId (New-Request)
        $loaded=Get-Context -Root $testRoot -UserId U004
        Assert-True ($updated.version -eq 2 -and $loaded.user.division -eq '전자심사과') 'updated affiliation reloads'
        $first=@(Get-Revisions -Path $saved.path)[0]
        Assert-True ($first.data.division -eq '통신심사과') 'original user revision is preserved'
        $invalid=$updated | Select-Object *;$invalid.bureau=@('not','text')
        Assert-Throws {Save-MasterRecord -Context $admin -Type Users -Id U004 -Data $invalid -ExpectedVersion 2 -RequestId (New-Request)} '문자열'
        $invalid=$updated | Select-Object *;$invalid.bureau=@('one entry is still an array')
        Assert-Throws {Save-MasterRecord -Context $admin -Type Users -Id U004 -Data $invalid -ExpectedVersion 2 -RequestId (New-Request)} '문자열'
        $invalid=$updated | Select-Object *;$invalid.division=('가'*101)
        Assert-Throws {Save-MasterRecord -Context $admin -Type Users -Id U004 -Data $invalid -ExpectedVersion 2 -RequestId (New-Request)} '최대 100'
        $owner=@(Get-Users $admin | Where-Object {$_.userId -eq 'U002'})[0]
        $owner.bureau='기계금속심사국';$owner.division='운송기계심사과'
        $null=Save-MasterRecord -Context $admin -Type Users -Id U002 -Data $owner -ExpectedVersion $owner.version -RequestId (New-Request)
    }

    Test-Case 'legacy users without affiliation load as blank without rewriting history' {
        $path=Join-Path $testRoot 'Master/Users/ULegacy'
        $legacy=[pscustomobject]@{userId='ULegacy';sid='DEMO-ULegacy';name='기존 사용자';roles=@('examiner');active=$true}
        $null=Write-Revision -Path $path -ExpectedVersion 0 -RequestId (New-Request) -ActorId U001 -Kind Users -TargetId ULegacy -Data $legacy
        $revisionFile=Join-Path $path 'revisions/000001.json';$before=Get-FileHash -LiteralPath $revisionFile
        $loaded=Get-Context -Root $testRoot -UserId ULegacy
        $listed=@(Get-Users $admin | Where-Object {$_.userId -eq 'ULegacy'})[0]
        Assert-True ($loaded.user.bureau -eq '' -and $loaded.user.division -eq '' -and $listed.bureau -eq '' -and $listed.division -eq '') 'missing affiliation projects as blank'
        Assert-True ((Get-FileHash -LiteralPath $revisionFile).Hash -eq $before.Hash) 'legacy user bytes stay intact'
        Assert-True ($null -eq (Read-Revision -Path $path).data.PSObject.Properties['bureau']) 'blank read projection is not persisted'
    }

    Test-Case 'new users receive automatic identifiers and retries preserve the first registration' {
        $request=New-Request
        $created=New-User -Context $admin -Name ' 신규 심사관 ' -Bureau ' 특허심사국 ' -Division ' 정보처리심사과 ' -Roles @('examiner','ai') -RequestId $request -Reason '참여자 등록'
        Assert-True ($created.userId -match '^U[0-9a-f]{32}$' -and $created.sid -match '^APP-[0-9a-f]{32}$') 'internal user ID and SID are generated'
        Assert-True ($created.name -eq '신규 심사관' -and $created.bureau -eq '특허심사국' -and $created.division -eq '정보처리심사과') 'registration trims profile text'
        $again=New-User -Context $admin -Name '신규 심사관' -Bureau '특허심사국' -Division '정보처리심사과' -Roles @('ai','examiner') -RequestId $request -Reason '참여자 등록'
        Assert-True ($again.userId -eq $created.userId -and $again.sid -eq $created.sid -and $again.version -eq 1) 'same intent registers only once'
        Assert-Throws {New-User -Context $admin -Name '다른 심사관' -Bureau '특허심사국' -Division '정보처리심사과' -Roles @('examiner','ai') -RequestId $request -Reason '참여자 등록'} '요청 ID'
        Assert-Throws {New-User -Context $admin -Name '신규 심사관' -Bureau '특허심사국' -Division '정보처리심사과' -Roles @('examiner','ai') -RequestId $request -Reason '다른 사유'} '요청 ID'
        Assert-Throws {New-User -Context $examiner -Name '권한 없음' -RequestId (New-Request)} '관리자'
    }

    Test-Case 'master registration supplies a missing SID and preserves legacy identifiers on update' {
        $request=New-Request
        $profile=[pscustomobject]@{userId='UAutoSid';name='자동 식별자';bureau='';division='';roles=@('examiner');active=$true}
        $created=Save-MasterRecord -Context $admin -Type Users -Id UAutoSid -Data $profile -ExpectedVersion 0 -RequestId $request
        $again=Save-MasterRecord -Context $admin -Type Users -Id UAutoSid -Data $profile -ExpectedVersion 0 -RequestId $request
        Assert-True ($created.sid -match '^APP-[0-9a-f]{32}$' -and $created.sid -eq $again.sid -and $again.version -eq 1) 'missing SID creation is idempotent'
        $legacy=@(Get-Users $admin | Where-Object {$_.userId -eq 'ULegacy'})[0]
        $legacy.sid='caller-attempted-change';$legacy.name='수정한 기존 사용자'
        $updated=Save-MasterRecord -Context $admin -Type Users -Id ULegacy -Data $legacy -ExpectedVersion $legacy.version -RequestId (New-Request)
        Assert-True ($updated.userId -eq 'ULegacy' -and $updated.sid -eq 'DEMO-ULegacy' -and $updated.name -eq '수정한 기존 사용자') 'legacy IDs and SID stay unchanged'
        $withoutSid=$updated | Select-Object * -ExcludeProperty sid
        $again=Save-MasterRecord -Context $admin -Type Users -Id ULegacy -Data $withoutSid -ExpectedVersion $updated.version -RequestId (New-Request)
        Assert-True ($again.sid -eq 'DEMO-ULegacy') 'updates without a SID keep the stored value'
        Assert-Throws {Save-MasterRecord -Context $admin -Type Users -Id 'DifferentId' -Data $again -ExpectedVersion 0 -RequestId (New-Request)} 'ID'
    }

    Test-Case 'login choices sort by division and name and disambiguate duplicate names' {
        $one=New-User -Context $admin -Name '동명이인' -Bureau '기계심사국' -Division '같은심사과' -RequestId (New-Request)
        $two=New-User -Context $admin -Name '동명이인' -Bureau '전기심사국' -Division '같은심사과' -RequestId (New-Request)
        $three=New-User -Context $admin -Name '동명이인' -Bureau '기계심사국' -Division '같은심사과' -RequestId (New-Request)
        $plain=New-User -Context $admin -Name '단독 이름' -Division '단독심사과' -RequestId (New-Request)
        $inactive=New-User -Context $admin -Name '비활성 숨김' -Division '숨김심사과' -RequestId (New-Request) -Active $false
        $list=@(Get-LoginUsers -Root $testRoot)
        $duplicates=@($list | Where-Object {$_.name -eq '동명이인'})
        Assert-True ($duplicates.Count -eq 3 -and @($duplicates.displayName | Select-Object -Unique).Count -eq 3) 'all duplicate labels are distinct'
        Assert-True (@($duplicates | Where-Object {$_.displayName -like '같은심사과 - 동명이인 (*'}).Count -eq 3) 'disambiguation preserves department-name prefix'
        Assert-True ((@($list | Where-Object {$_.userId -eq $plain.userId})[0]).displayName -eq '단독심사과 - 단독 이름') 'unique name has no suffix'
        Assert-True (@($list | Where-Object {$_.userId -eq $inactive.userId}).Count -eq 0) 'inactive users are hidden'
        Assert-True ((@($list | Where-Object {$_.userId -eq 'U001'})[0]).division -eq '미지정') 'blank division has a usable label'
        Assert-True (($list.userId -join ',') -eq ((@($list | Sort-Object division,name,userId).userId) -join ',')) 'list has stable department/name/ID ordering'
        Assert-True (($list.displayName -join '|') -eq ((@(Get-LoginUsers -Root $testRoot).displayName) -join '|')) 'labels stay stable between reads'
        Assert-Throws {Get-Context -Root $testRoot -UserId $inactive.userId} '활성 사용자'
    }

    $submissionRequest=New-Request
    $case=New-Submission -Context $examiner -Period $month -ApplicationNumber 1020260000001 -RequestId $submissionRequest -Cpc ' H04L 9/00; G06F 21/00 '
    Test-Case 'submission normalization, duplicate protection and idempotent retry' {
        Assert-True ($case.applicationNumber -eq '10-2026-0000001') 'normalize 13 digits'
        $again=New-Submission -Context $examiner -Period $month -ApplicationNumber '10-2026-0000001' -RequestId $submissionRequest -Cpc 'H04L 9/00; G06F 21/00'
        Assert-True ($again.caseId -eq $case.caseId) 'same request returns same case'
        Assert-True ($case.cpc -eq 'H04L 9/00; G06F 21/00') 'CPC is trimmed and persisted as text'
        Assert-True ($case.examinerBureau -eq '기계금속심사국' -and $case.examinerDivision -eq '운송기계심사과') 'submission snapshots authoritative user affiliation'
        Assert-Throws {New-Submission -Context $examiner -Period $month -ApplicationNumber '10-2026-0000001' -RequestId $submissionRequest -Cpc 'H04L 1/00'} '요청 ID'
        Assert-Throws {New-Submission -Context $examiner -Period $month -ApplicationNumber '10-2026-0000001' -RequestId $submissionRequest} '요청 ID'
        Assert-Throws {New-Submission -Context $examiner -Period $month -ApplicationNumber '10-2026-0000002' -RequestId (New-Request) -Cpc ('A'*501)} '최대 500'
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

    Test-Case 'affiliation updates do not change existing submission snapshots' {
        $owner=@(Get-Users $admin | Where-Object {$_.userId -eq 'U002'})[0]
        $owner.bureau='전기통신심사국';$owner.division='통신심사과'
        $null=Save-MasterRecord -Context $admin -Type Users -Id U002 -Data $owner -ExpectedVersion $owner.version -RequestId (New-Request)
        $reloaded=@(Get-Cases $examiner $month Mine | Where-Object {$_.caseId -eq $case.caseId})[0]
        Assert-True ($reloaded.examinerBureau -eq '기계금속심사국' -and $reloaded.examinerDivision -eq '운송기계심사과') 'record retains submission-time affiliation'
        Assert-True ($reloaded.cpc -eq $case.cpc) 'CPC remains attached to its case'
        $again=New-Submission -Context $examiner -Period $month -ApplicationNumber $case.applicationNumber -RequestId $submissionRequest -Cpc $case.cpc
        Assert-True ($again.caseId -eq $case.caseId -and $again.examinerDivision -eq '운송기계심사과') 'retry returns original snapshot despite profile change'
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

    Test-Case 'seven-document cap and legacy no-results validation' {
        Assert-Throws {Save-Delivery -Context $admin -Case $case -SourceId S02 -ExpectedVersion 1 -RequestId (New-Request) -Status Published -Documents @()} '한 건 이상'
        $tooMany=@(1..8 | ForEach-Object {New-Document ('P'+$_)})
        Assert-Throws {Save-Delivery -Context $admin -Case $case -SourceId S02 -ExpectedVersion 1 -RequestId (New-Request) -Status Draft -Documents $tooMany} '최대 7'
        Assert-Throws {Save-Delivery -Context $admin -Case $case -SourceId S02 -ExpectedVersion 1 -RequestId (New-Request) -Status Draft -NoResults $true} '사유'
        Assert-Throws {Save-Delivery -Context $admin -Case $case -SourceId S02 -ExpectedVersion 1 -RequestId (New-Request) -Status Draft -NoResults $true -NoResultsReason why -Documents $docs} '동시에'
    }

    Test-Case 'identifier-only AI results publish without a title and preserve hidden legacy text on edit' {
        $numberCase=New-Submission -Context $admin -Period $month -ApplicationNumber '10-2026-0000099' -RequestId (New-Request)
        $numberOnly=Save-Delivery -Context $admin -Case $numberCase -SourceId S02 -ExpectedVersion 0 -RequestId (New-Request) -Status Published -Documents @(@{identifier=' 10-2025-0000123 '})
        Assert-True ($numberOnly.data.documents[0].identifier -eq '10-2025-0000123' -and $numberOnly.data.documents[0].title -eq '') 'number alone can be published'
        Assert-Throws {Save-Delivery -Context $admin -Case $numberCase -SourceId S02 -ExpectedVersion 1 -RequestId (New-Request) -Status Draft -Documents @(@{identifier='  '})} '식별정보'
        $withText=@($numberOnly.data.documents);$withText[0].title='보존할 명칭';$withText[0].description='보존할 관련 설명';$withText[0].type='LegacyCustom'
        $legacy=Save-Delivery -Context $admin -Case $numberCase -SourceId S02 -ExpectedVersion 1 -RequestId (New-Request) -Status Draft -Documents $withText
        $edited=Save-Delivery -Context $admin -Case $numberCase -SourceId S02 -ExpectedVersion 2 -RequestId (New-Request) -Status Published -Documents @(@{documentId=$legacy.data.documents[0].documentId;identifier='10-2025-0000124'})
        Assert-True ($edited.data.documents[0].title -eq '보존할 명칭' -and $edited.data.documents[0].description -eq '보존할 관련 설명' -and $edited.data.documents[0].type -eq 'LegacyCustom') 'number-only editing retains legacy fields'
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
    $s02v3=Save-Delivery -Context $admin -Case $case -SourceId S02 -ExpectedVersion 2 -RequestId (New-Request) -Status Published -Documents $newDocs
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
        $closed=Set-PeriodStatus -Context $admin -Period $month -Status Closed -ExpectedVersion 1 -RequestId $closeRequest
        Assert-Throws {Save-Delivery -Context $ai -Case $case -SourceId S03 -ExpectedVersion 1 -RequestId (New-Request) -Status Draft -Documents @($s03.data.documents)} '마감'
        Assert-Throws {New-Submission $examiner $month '10-2026-0000002' (New-Request)} '마감'
        Assert-True ((Set-PeriodStatus -Context $admin -Period $month -Status Closed -ExpectedVersion 1 -RequestId $closeRequest).version -eq 2) 'close retry uses original revision'
        Assert-Throws {Set-PeriodStatus -Context $admin -Period $month -Status Open -ExpectedVersion 1 -RequestId $closeRequest} '요청 ID'
        $opened=Set-PeriodStatus -Context $admin -Period $month -Status Open -ExpectedVersion 2 -RequestId (New-Request)
        Assert-True ($opened.status -eq 'Open') 'month reopened'
    }

    Test-Case 'cancellation preserves original submission and blocks subsequent writes' {
        $extra=New-Submission $examiner $month '10-2026-0000002' (New-Request) '' 'G06F 16/00'
        $owner=@(Get-Users $admin | Where-Object {$_.userId -eq 'U002'})[0]
        $owner.bureau='변경된 심사국';$owner.division='변경된 심사과'
        $null=Save-MasterRecord -Context $admin -Type Users -Id U002 -Data $owner -ExpectedVersion $owner.version -RequestId (New-Request)
        $cancelRequest=New-Request
        $cancelled=Cancel-Submission -Context $examiner -Case $extra -ExpectedVersion 1 -RequestId $cancelRequest
        Assert-True ($cancelled.status -eq 'Cancelled' -and $cancelled.version -eq 2) 'cancel appends revision'
        Assert-True ($cancelled.cpc -eq 'G06F 16/00' -and $cancelled.examinerBureau -eq $extra.examinerBureau -and $cancelled.examinerDivision -eq $extra.examinerDivision) 'cancel preserves original CPC and affiliation snapshot'
        Assert-True ((Cancel-Submission -Context $examiner -Case $extra -ExpectedVersion 1 -RequestId $cancelRequest).version -eq 2) 'cancel retry'
        Assert-Throws {Cancel-Submission -Context $examiner -Case $extra -ExpectedVersion 1 -RequestId $cancelRequest -Reason '다른 사유'} '요청 ID'
        Assert-Throws {Save-Delivery -Context $examiner -Case $extra -SourceId S01 -ExpectedVersion 0 -RequestId (New-Request) -Status Published -NoResults $true -NoResultsReason none} '취소'
    }

    Test-Case 'legacy submissions retry with the original fingerprint and blank new fields' {
        $legacyRequest=New-Request
        # This is the exact sorted JSON fingerprint input used before CPC existed.
        $legacyInput='{"applicationNumber":"10-2026-0000003","operation":"NewSubmission","period":"'+$month+'","reason":"","userId":"U002"}'
        $hash=[Security.Cryptography.SHA256]::Create()
        try {$legacyFingerprint=([BitConverter]::ToString($hash.ComputeHash([Text.Encoding]::UTF8.GetBytes($legacyInput)))).Replace('-','').ToLowerInvariant()}
        finally {$hash.Dispose()}
        $path=Join-Path $testRoot ('Data/'+$month+'/G01/Submissions/U002/CLegacyCpc')
        $legacy=[pscustomobject]@{caseId='CLegacyCpc';period=$month;groupId='G01';userId='U002';examinerName='기존 심사관';applicationNumber='10-2026-0000003';status='Submitted';submittedAt=[datetimeoffset]::UtcNow.ToString('o');duplicateWarning=$false;requestFingerprint=$legacyFingerprint}
        $null=Write-Revision -Path $path -ExpectedVersion 0 -RequestId $legacyRequest -ActorId U002 -Kind Submission -TargetId CLegacyCpc -Data $legacy
        $revisionFile=Join-Path $path 'revisions/000001.json';$before=Get-FileHash -LiteralPath $revisionFile
        $loaded=@(Get-Cases $examiner $month Mine | Where-Object {$_.caseId -eq 'CLegacyCpc'})[0]
        Assert-True ($loaded.cpc -eq '' -and $loaded.examinerBureau -eq '' -and $loaded.examinerDivision -eq '') 'legacy case stays blank instead of taking current user affiliation'
        $retry=New-Submission -Context $examiner -Period $month -ApplicationNumber '10-2026-0000003' -RequestId $legacyRequest -Cpc '  '
        Assert-True ($retry.caseId -eq $loaded.caseId -and $retry.version -eq 1 -and $retry.cpc -eq '') 'blank CPC retry uses exact legacy fingerprint'
        Assert-Throws {New-Submission -Context $examiner -Period $month -ApplicationNumber '10-2026-0000003' -RequestId $legacyRequest -Cpc 'H04L 9/00'} '요청 ID'
        Assert-True ((Get-FileHash -LiteralPath $revisionFile).Hash -eq $before.Hash) 'legacy submission bytes stay intact after reads and retry'
        $cancelled=Cancel-Submission -Context $examiner -Case $loaded -ExpectedVersion 1 -RequestId (New-Request) -Reason '기존 건 취소'
        Assert-True ($cancelled.cpc -eq '' -and $cancelled.examinerBureau -eq '' -and $cancelled.examinerDivision -eq '') 'legacy cancellation keeps historical fields blank'
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

    Test-Case 'admin period edits preserve source snapshots and allow lifecycle changes without reasons' {
        $period=Get-Period $admin $month
        $bad=$period | Select-Object *
        $bad.sourceIds=@('S01','S02');$bad.assignments=@($bad.assignments | Where-Object {$_.sourceId -ne 'S03'})
        Assert-Throws {Save-Period -Context $admin -Period $month -Data $bad -ExpectedVersion $period.version -RequestId (New-Request) -Reason 'remove source'} '제거할 수 없'
        $bad=$period | Select-Object *;$bad.status='Closed'
        $closed=Save-Period -Context $admin -Period $month -Data $bad -ExpectedVersion $period.version -RequestId (New-Request)
        Assert-True ($closed.status -eq 'Closed') 'period can close without a reason'
        $period=Set-PeriodStatus -Context $admin -Period $month -Status Open -ExpectedVersion $closed.version -RequestId (New-Request)
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

    Test-Case 'production roots allow active user selection without matching a Windows SID' {
        $marker=Join-Path $testRoot '.demo-mode.json';$renamed=Join-Path $testRoot '.demo-mode.disabled'
        [IO.File]::Move($marker,$renamed)
        try {
            $selected=Get-Context -Root $testRoot -UserId U002
            Assert-True ($selected.user.userId -eq 'U002' -and -not $selected.demoMode) 'department/name selector can choose an active production user'
            Assert-True (@(Get-Cases $selected $month Mine).Count -gt 0) 'selected examiner can use existing case data'
            Assert-True ((Get-Context -Root $testRoot).user.userId -eq 'U001') 'script default chooses an active admin deterministically'
            Assert-True (@(Get-LoginUsers -Root $testRoot | Where-Object {$_.userId -eq 'U003'}).Count -eq 0) 'deactivated user is absent from production login choices'
            Assert-Throws {Get-Context -Root $testRoot -UserId U003} '활성 사용자'
            Assert-Throws {Get-Cases $ai $month Assigned} '활성 사용자'
        } finally {[IO.File]::Move($renamed,$marker)}
    }

    Test-Case 'launcher access modes separate general-user and automatic administrator login in production' {
        $modeRoot=Join-Path $testRoot 'LauncherModes'
        $setup=Initialize-Demo -Root $modeRoot
        [IO.File]::Move((Join-Path $modeRoot '.demo-mode.json'),(Join-Path $modeRoot '.demo-mode.disabled'))
        $onlyAdmin=New-User -Context $setup -Name '관리 전용' -Roles @('admin') -RequestId (New-Request)
        $inactiveAdmin=New-User -Context $setup -Name '비활성 관리자' -Roles @('admin') -Active $false -RequestId (New-Request)
        $adminSession=Get-Context -Root $modeRoot -AccessMode Admin
        $normalSession=Get-Context -Root $modeRoot -UserId U001 -AccessMode User
        $peer=Get-Context -Root $modeRoot -UserId U002 -AccessMode User
        Assert-True ($adminSession.accessMode -eq 'Admin' -and $adminSession.user.userId -eq 'U001' -and $adminSession.user.roles -contains 'admin') 'admin launcher chooses the first active admin without a selector'
        Assert-True ((Get-Context -Root $modeRoot -AccessMode Admin).user.userId -eq $adminSession.user.userId) 'automatic admin selection stays stable across reads'
        Assert-True ($normalSession.accessMode -eq 'User' -and -not $normalSession.demoMode -and $normalSession.user.roles -notcontains 'admin') 'normal production session hides admin UI roles even for a multi-role administrator'
        Assert-True ($normalSession.user.roles -contains 'examiner' -and $normalSession.user.roles -contains 'ai') 'normal session retains ordinary work roles'
        $normalChoices=@(Get-LoginUsers -Root $modeRoot -AccessMode User)
        $adminChoices=@(Get-LoginUsers -Root $modeRoot -AccessMode Admin)
        Assert-True (@($normalChoices | Where-Object {$_.userId -eq 'U001'}).Count -eq 1 -and @($normalChoices | Where-Object {$_.userId -eq $onlyAdmin.userId}).Count -eq 0) 'general-user choices include working administrators and exclude admin-only accounts'
        Assert-True (@($adminChoices | Where-Object {$_.userId -eq 'U002' -or $_.userId -eq $inactiveAdmin.userId}).Count -eq 0 -and $adminChoices.Count -eq 2) 'admin choices contain active administrators only'
        Assert-Throws {Get-Context -Root $modeRoot -UserId $onlyAdmin.userId -AccessMode User} '일반 사용자'
        Assert-Throws {Get-Context -Root $modeRoot -UserId U002 -AccessMode Admin} '총괄관리자'
        $ownCase=New-Submission -Context $normalSession -Period $month -ApplicationNumber '10-2026-0000151' -RequestId (New-Request)
        $peerCase=New-Submission -Context $peer -Period $month -ApplicationNumber '10-2026-0000152' -RequestId (New-Request)
        Assert-True (@(Get-Cases -Context $normalSession -Period $month -Scope Mine).Count -eq 1) 'normal session can submit and list its own work'
        Assert-True (@(Get-Cases -Context $adminSession -Period $month -Scope All).Count -eq 2) 'admin mode can inspect all work'
        Assert-Throws {Get-Cases -Context $normalSession -Period $month -Scope All} '관리자'
        Assert-Throws {Save-MasterRecord -Context $normalSession -Type Groups -Id G02 -Data @{groupId='G02';name='일반 실행 차단';active=$true} -ExpectedVersion 0 -RequestId (New-Request)} '관리자'
        Assert-Throws {New-User -Context $normalSession -Name '일반 실행 차단' -RequestId (New-Request)} '관리자'
        Assert-Throws {Set-PeriodStatus -Context $normalSession -Period $month -Status Closed -ExpectedVersion 1 -RequestId (New-Request) -Reason '일반 실행 차단'} '관리자'
        Assert-Throws {Save-Period -Context $normalSession -Period $month -Data (Get-Period $normalSession $month) -ExpectedVersion 1 -RequestId (New-Request)} '관리자'
        Assert-Throws {Save-Delivery -Context $normalSession -Case $peerCase -SourceId S03 -ExpectedVersion 0 -RequestId (New-Request) -Status Published -Documents @((New-Document P3))} '권한'
        $null=Save-Delivery -Context $adminSession -Case $peerCase -SourceId S03 -ExpectedVersion 0 -RequestId (New-Request) -Status Published -Documents @((New-Document P3))
        Assert-True (@(Get-Deliveries -Context $normalSession -Case $peerCase).Count -eq 0) 'authoritative refresh does not restore admin read access to an unassigned AI result'
        Assert-True ((Get-Context -Root $modeRoot -UserId U001).user.roles -contains 'admin') 'general-user role filtering never changes the stored administrator roles'

        $backup=Get-Context -Root $modeRoot -UserId $onlyAdmin.userId -AccessMode Admin
        $former=@(Get-Users $backup | Where-Object {$_.userId -eq 'U001'})[0]
        $former.roles=@('examiner','ai')
        $null=Save-MasterRecord -Context $backup -Type Users -Id U001 -Data $former -ExpectedVersion $former.version -RequestId (New-Request)
        Assert-Throws {Get-Periods -Context $adminSession} '총괄관리자'
        Assert-Throws {Get-Cases -Context $adminSession -Period $month -Scope All} '총괄관리자'
        Assert-True ((Get-Context -Root $modeRoot -AccessMode Admin).user.userId -eq $onlyAdmin.userId) 'admin auto-login skips accounts whose admin role was revoked'
        $current=@(Get-Users $backup | Where-Object {$_.userId -eq $onlyAdmin.userId})[0]
        $current.active=$false
        Assert-Throws {Save-MasterRecord -Context $backup -Type Users -Id $current.userId -Data $current -ExpectedVersion $current.version -RequestId (New-Request)} '본인의 관리자 권한'
        # Model an external provisioning change while keeping the existing
        # protection against an administrator disabling their own account.
        $null=Write-Revision -Path $current.path -ExpectedVersion $current.version -RequestId (New-Request) -ActorId setup -Kind Users -TargetId $current.userId -Data ($current | Select-Object * -ExcludeProperty path,version)
        Assert-Throws {Get-Periods -Context $backup} '활성 사용자'
        Assert-Throws {Get-Context -Root $modeRoot -AccessMode Admin} '활성 총괄관리자'
        Assert-True (@(Get-LoginUsers -Root $modeRoot -AccessMode Admin).Count -eq 0) 'no active administrator is never replaced with a regular user'
        Assert-True ((Get-Context -Root $modeRoot -UserId U002 -AccessMode User).user.roles -notcontains 'admin') 'ordinary login remains available when no admin is active'
        $peerPath=Join-Path $modeRoot 'Master/Users/U002'
        $peerRevision=Read-Revision -Path $peerPath
        $peerRevision.data.roles=@('admin')
        $null=Write-Revision -Path $peerPath -ExpectedVersion $peerRevision.version -RequestId (New-Request) -ActorId setup -Kind Users -TargetId U002 -Data $peerRevision.data
        Assert-Throws {Get-Cases -Context $peer -Period $month -Scope Mine} '일반 사용자'
        Assert-Throws {Get-Context -Root $modeRoot -UserId U002 -AccessMode User} '일반 사용자'
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
