Set-StrictMode -Version 2.0

$script:WJsonHasDateKind = (Get-Command ConvertFrom-Json).Parameters.ContainsKey('DateKind')

function ConvertFrom-WJson {
    param([string]$Json)
    if ($script:WJsonHasDateKind) { return ConvertFrom-Json -InputObject $Json -DateKind String }
    return ConvertFrom-Json -InputObject $Json
}

function ConvertTo-WCanonical {
    param($Value)
    if ($null -eq $Value) { return 'null' }
    if ($Value -is [System.Collections.IDictionary]) {
        $parts = @($Value.Keys | Sort-Object | ForEach-Object { (ConvertTo-Json -InputObject ([string]$_) -Compress) + ':' + (ConvertTo-WCanonical $Value[$_]) })
        return '{' + ($parts -join ',') + '}'
    }
    if ($Value -is [pscustomobject]) {
        $parts = @($Value.PSObject.Properties.Name | Sort-Object | ForEach-Object { (ConvertTo-Json -InputObject ([string]$_) -Compress) + ':' + (ConvertTo-WCanonical $Value.$_) })
        return '{' + ($parts -join ',') + '}'
    }
    if ($Value -is [array] -or $Value -is [System.Collections.IList]) {
        $parts = @($Value | ForEach-Object { ConvertTo-WCanonical $_ })
        return '[' + ($parts -join ',') + ']'
    }
    return ConvertTo-Json -InputObject $Value -Depth 60 -Compress
}

function Get-WFingerprint {
    param($Value)
    $canonical = ConvertTo-WCanonical (ConvertFrom-WJson (ConvertTo-Json -InputObject $Value -Depth 60 -Compress))
    $hash = [System.Security.Cryptography.SHA256]::Create()
    try { return ([System.BitConverter]::ToString($hash.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($canonical)))).Replace('-','').ToLowerInvariant() }
    finally { $hash.Dispose() }
}

function Assert-WRequest {
    param($Envelope, [string]$Fingerprint, [string]$ActorId)
    if ([string](Get-WValue $Envelope.data 'requestFingerprint') -cne $Fingerprint -or $Envelope.actorId -cne $ActorId) {
        throw '동일 저장 요청 ID를 다른 입력이나 작성자에 재사용할 수 없습니다.'
    }
}

if (-not (Get-Command Read-Revision -ErrorAction SilentlyContinue)) {
    Import-Module (Join-Path $PSScriptRoot 'Storage.psm1') -ErrorAction Stop
}

function Get-WValue {
    param($Object, [string]$Name, $Default = $null)
    if ($null -eq $Object) { return $Default }
    if ($Object -is [System.Collections.IDictionary]) {
        if ($Object.Contains($Name)) { return $Object[$Name] }
    } elseif ($null -ne $Object.PSObject.Properties[$Name]) { return $Object.$Name }
    return $Default
}

function Set-WValue {
    param($Object, [string]$Name, $Value)
    if ($Object -is [System.Collections.IDictionary]) { $Object[$Name] = $Value }
    else { $Object | Add-Member -NotePropertyName $Name -NotePropertyValue $Value -Force }
}

function Convert-WOptionalText {
    param($Value, [string]$Label, [int]$MaximumLength = 100)
    if ($null -eq $Value) { return '' }
    if ($Value -isnot [string]) { throw "$Label 값은 문자열이어야 합니다." }
    $text = $Value.Trim()
    if ($text.Length -gt $MaximumLength) { throw "$Label 값은 최대 $MaximumLength 자까지 입력하십시오." }
    return $text
}

function Assert-WId {
    param([string]$Value, [string]$Label = 'ID')
    if ($Value -notmatch '^[A-Za-z0-9][A-Za-z0-9_-]{0,100}$') { throw "$Label 값이 올바르지 않습니다." }
}

function Assert-WPeriod {
    param([string]$Period)
    if ($Period -notmatch '^\d{4}-(0[1-9]|1[0-2])$') { throw '운영월은 YYYY-MM 형식이어야 합니다.' }
}

function Get-WPeriodLock {
    param($Context, [string]$Period)
    Assert-WPeriod $Period
    return Join-Path $Context.root ('Locks\{0}\lifecycle.lock' -f $Period)
}

function Invoke-WPeriodReadLock {
    param($Context, [string]$Period, [scriptblock]$Action)
    # Shared file handles permit independent case saves. Month close/reopen uses
    # FileShare.None on the same file and waits for those short saves to finish.
    $periodGatePath = Get-WPeriodLock $Context $Period
    $null = [System.IO.Directory]::CreateDirectory([System.IO.Path]::GetDirectoryName($periodGatePath))
    $timer = [System.Diagnostics.Stopwatch]::StartNew()
    $handle = $null
    while ($null -eq $handle) {
        try { $handle = New-Object System.IO.FileStream($periodGatePath, [System.IO.FileMode]::OpenOrCreate, [System.IO.FileAccess]::Read, [System.IO.FileShare]::Read) }
        catch [System.IO.IOException] {
            if ($timer.ElapsedMilliseconds -ge 5000) { throw '운영월 상태 변경 작업이 진행 중입니다. 잠시 후 다시 저장하십시오.' }
            Start-Sleep -Milliseconds 40
        }
    }
    try { & $Action }
    finally { $handle.Dispose(); $timer.Stop() }
}

function Get-WMasterPath {
    param([string]$Root, [string]$Type, [string]$Id)
    Assert-WId $Id
    return Join-Path (Join-Path (Join-Path $Root 'Master') $Type) $Id
}

function Convert-WRecord {
    param($Envelope, [string]$Path)
    if ($null -eq $Envelope) { return $null }
    $record = ConvertFrom-WJson (ConvertTo-Json -InputObject (Get-WValue $Envelope 'data') -Depth 60)
    # New optional fields are projected at read time; immutable older JSON stays intact.
    if ((Get-WValue $Envelope 'kind') -eq 'Users') {
        foreach ($field in @('bureau','division')) { Set-WValue $record $field ([string](Get-WValue $record $field '')) }
    }
    if ((Get-WValue $Envelope 'kind') -eq 'Submission') {
        foreach ($field in @('cpc','examinerBureau','examinerDivision')) { Set-WValue $record $field ([string](Get-WValue $record $field '')) }
    }
    Set-WValue $record 'version' ([int]$Envelope.version)
    Set-WValue $record 'path' $Path
    return $record
}

function Get-WMasterRecords {
    param([string]$Root, [string]$Type)
    $parent = Join-Path (Join-Path $Root 'Master') $Type
    if (-not [System.IO.Directory]::Exists($parent)) { return }
    foreach ($folder in Get-ChildItem -LiteralPath $parent -Directory | Sort-Object Name) {
        $revision = Read-Revision -Path $folder.FullName
        if ($null -ne $revision) { Convert-WRecord $revision $folder.FullName }
    }
}

function Convert-WAccessUser {
    param($User, [ValidateSet('User','Admin','Unrestricted')][string]$AccessMode = 'Unrestricted')
    # Session roles come from both the current user record and the launcher mode.
    # Clone the record so a normal login never edits the stored administrator roles.
    $copy = ConvertFrom-WJson (ConvertTo-Json -InputObject $User -Depth 60)
    if ($AccessMode -eq 'User') {
        if (-not ((Test-WRole $copy 'examiner') -or (Test-WRole $copy 'ai'))) { throw '일반 사용자로 로그인하려면 심사관 또는 AI 담당자 역할이 필요합니다.' }
        Set-WValue $copy 'roles' @((Get-WValue $copy 'roles' @()) | Where-Object { $_ -ne 'admin' })
    } elseif ($AccessMode -eq 'Admin' -and -not (Test-WRole $copy 'admin')) {
        throw '활성 총괄관리자 계정을 찾을 수 없습니다. 관리자 계정과 역할을 확인하십시오.'
    }
    return $copy
}

function Assert-WContext {
    param($Context)
    $root = [System.IO.Path]::GetFullPath([string](Get-WValue $Context 'root'))
    $userId = [string](Get-WValue (Get-WValue $Context 'user') 'userId')
    Assert-WId $userId '사용자 ID'
    $record = Read-Revision -Path (Get-WMasterPath $root 'Users' $userId)
    if ($null -eq $record -or -not [bool](Get-WValue $record.data 'active' $true)) { throw '활성 사용자 등록을 확인할 수 없습니다.' }
    # Refresh activity and roles on every operation without tying users to PCs.
    # A general-user launcher cannot recover admin rights from a refreshed record.
    $accessMode = [string](Get-WValue $Context 'accessMode' 'Unrestricted')
    return Convert-WAccessUser -User $record.data -AccessMode $accessMode
}

function Test-WRole {
    param($User, [string]$Role)
    return @((Get-WValue $User 'roles' @())) -contains $Role
}

function Assert-WAdmin {
    param($Context)
    $user = Assert-WContext $Context
    if (-not (Test-WRole $user 'admin')) { throw '총괄관리자 권한이 필요합니다.' }
    return $user
}

function Get-Users {
    param([Parameter(Mandatory)]$Context)
    $user = Assert-WContext $Context
    $all = @(Get-WMasterRecords $Context.root 'Users')
    if (Test-WRole $user 'admin') { return $all }
    if (Test-WRole $user 'ai') { return $all | Select-Object userId, name, bureau, division, active, version }
    return $all | Where-Object { $_.userId -eq $user.userId }
}

function Get-Groups {
    param([Parameter(Mandatory)]$Context)
    $null = Assert-WContext $Context
    Get-WMasterRecords $Context.root 'Groups'
}

function Get-Sources {
    param([Parameter(Mandatory)]$Context)
    $null = Assert-WContext $Context
    Get-WMasterRecords $Context.root 'Sources'
}

function Get-Period {
    param([Parameter(Mandatory)]$Context, [Parameter(Mandatory)][string]$Period)
    $null = Assert-WContext $Context
    Assert-WPeriod $Period
    $path = Join-Path (Join-Path $Context.root 'Periods') $Period
    $envelope = Read-Revision -Path $path
    if ($null -eq $envelope) { throw "$Period 운영 설정이 없습니다." }
    return Convert-WRecord $envelope $path
}

function Get-Periods {
    param([Parameter(Mandatory)]$Context)
    $null = Assert-WContext $Context
    $path = Join-Path $Context.root 'Periods'
    if (-not [System.IO.Directory]::Exists($path)) { return }
    foreach ($folder in Get-ChildItem -LiteralPath $path -Directory | Sort-Object Name -Descending) {
        if ($folder.Name -match '^\d{4}-\d{2}$') {
            $revision = Read-Revision -Path $folder.FullName
            if ($null -ne $revision) { Convert-WRecord $revision $folder.FullName }
        }
    }
}

function Get-LoginUsers {
    param([Parameter(Mandatory)][string]$Root, [ValidateSet('User','Admin','All')][string]$AccessMode = 'All')
    $resolved = [System.IO.Path]::GetFullPath($Root)
    $entries = @(
        foreach ($record in @(Get-WMasterRecords $resolved 'Users')) {
            if (-not [bool](Get-WValue $record 'active' $true)) { continue }
            if ($AccessMode -eq 'User' -and -not ((Test-WRole $record 'examiner') -or (Test-WRole $record 'ai'))) { continue }
            if ($AccessMode -eq 'Admin' -and -not (Test-WRole $record 'admin')) { continue }
            $division = [string](Get-WValue $record 'division' '')
            if ([string]::IsNullOrWhiteSpace($division)) { $division = '미지정' }
            [pscustomobject]@{ userId = $record.userId; name = $record.name; bureau = [string](Get-WValue $record 'bureau' '');
                division = $division; displayName = ($division + ' - ' + $record.name); roles = @((Get-WValue $record 'roles' @())); active = $true }
        }
    )
    foreach ($duplicates in @($entries | Group-Object displayName | Where-Object { $_.Count -gt 1 })) {
        foreach ($entry in @($duplicates.Group)) {
            $suffix = $entry.bureau
            $sameBureau = @($duplicates.Group | Where-Object { $_.bureau -eq $entry.bureau })
            if (-not $suffix -or $sameBureau.Count -gt 1) {
                # Use a unique ID suffix, extending it for unusual imported IDs.
                $length = [Math]::Min(8, $entry.userId.Length)
                do {
                    $shortId = $entry.userId.Substring($entry.userId.Length - $length)
                    $collisions = @($sameBureau | Where-Object { $_.userId.EndsWith($shortId, [StringComparison]::OrdinalIgnoreCase) })
                    if ($collisions.Count -le 1 -or $length -ge $entry.userId.Length) { break }
                    $length++
                } while ($true)
                if ($suffix) { $suffix += ' / ' }
                $suffix += $shortId
            }
            $entry.displayName += ' (' + $suffix + ')'
        }
    }
    $entries | Sort-Object division, name, userId
}

function Get-Context {
    param([Parameter(Mandatory)][string]$Root, [string]$UserId,
        [ValidateSet('User','Admin','Unrestricted')][string]$AccessMode = 'Unrestricted')
    $resolved = [System.IO.Path]::GetFullPath($Root)
    $demo = [System.IO.File]::Exists((Join-Path $resolved '.demo-mode.json'))
    $users = @(Get-WMasterRecords $resolved 'Users' | Where-Object { [bool](Get-WValue $_ 'active' $true) })
    if ($UserId) { $matches = @($users | Where-Object { $_.userId -eq $UserId -and (Get-WValue $_ 'active' $true) }) }
    elseif ($AccessMode -eq 'Admin') { $matches = @($users | Where-Object { Test-WRole $_ 'admin' } | Sort-Object userId | Select-Object -First 1) }
    elseif ($AccessMode -eq 'User') { $matches = @($users | Where-Object { (Test-WRole $_ 'examiner') -or (Test-WRole $_ 'ai') } | Sort-Object userId | Select-Object -First 1) }
    else {
        # Backwards-compatible default for existing setup and maintenance scripts.
        $matches = @($users | Sort-Object @{Expression={ if (Test-WRole $_ 'admin') { 0 } else { 1 } }}, userId | Select-Object -First 1)
    }
    if ($matches.Count -ne 1) {
        if ($AccessMode -eq 'Admin') { throw '활성 총괄관리자 계정을 찾을 수 없습니다. 관리자 계정과 역할을 확인하십시오.' }
        throw '선택한 활성 사용자를 찾을 수 없습니다. 관리자에게 사용자 등록을 요청하십시오.'
    }
    $sessionUser = Convert-WAccessUser -User $matches[0] -AccessMode $AccessMode
    $context = [pscustomobject]@{ root = $resolved; user = $sessionUser; accessMode = $AccessMode; demoMode = $demo; demoImpersonation = [bool]$UserId; periods = @() }
    $context.periods = @(Get-Periods -Context $context)
    return $context
}

function Get-WMember {
    param($Period, [string]$UserId)
    return @((Get-WValue $Period 'members' @()) | Where-Object { $_.userId -eq $UserId }) | Select-Object -First 1
}

function Get-WAssignedSources {
    param($Period, [string]$UserId, [string]$GroupId)
    @((Get-WValue $Period 'assignments' @()) | Where-Object { $_.userId -eq $UserId -and $_.groupId -eq $GroupId } | ForEach-Object { $_.sourceId })
}

function Assert-WOpen {
    param($Period, [ValidateSet('Submission','Delivery','Evaluation')][string]$Activity)
    if ([string](Get-WValue $Period 'status') -ne 'Open') { throw '마감된 운영월입니다. 관리자가 재개방해야 수정할 수 있습니다.' }
    $field = $Activity.ToLowerInvariant() + 'Deadline'
    $deadline = [string](Get-WValue $Period $field '')
    if ($deadline) {
        $limit = [datetimeoffset]::Parse($deadline, [System.Globalization.CultureInfo]::InvariantCulture)
        if ([datetimeoffset]::UtcNow -gt $limit) { throw "$Activity 입력 기한이 지났습니다." }
    }
}

function Get-WCasePath {
    param($Context, $Case)
    Assert-WPeriod ([string]$Case.period)
    foreach ($name in @('groupId','userId','caseId')) { Assert-WId ([string]$Case.$name) $name }
    $path = Join-Path $Context.root ('Data\{0}\{1}\Submissions\{2}\{3}' -f $Case.period, $Case.groupId, $Case.userId, $Case.caseId)
    return $path
}

function Get-WCase {
    param($Context, $Case)
    $path = Get-WCasePath $Context $Case
    $envelope = Read-Revision -Path $path
    if ($null -eq $envelope) { throw '제출 건을 찾을 수 없습니다.' }
    $actual = Convert-WRecord $envelope $path
    if (-not (Test-WSelectionPublished $Context $actual)) { throw '아직 확정되지 않은 대상 건입니다.' }
    foreach ($name in @('caseId','period','groupId','userId')) {
        if ([string]$actual.$name -ne [string]$Case.$name) { throw '제출 건 참조정보가 일치하지 않습니다.' }
    }
    return $actual
}

function Assert-WCaseRead {
    param($Context, $Case)
    $user = Assert-WContext $Context
    if (Test-WRole $user 'admin') { return $user }
    if ($Case.userId -eq $user.userId -and (Test-WRole $user 'examiner')) { return $user }
    $period = Get-Period $Context $Case.period
    if ((Test-WRole $user 'ai') -and @(Get-WAssignedSources $period $user.userId $Case.groupId).Count -gt 0) { return $user }
    throw '이 제출 건에 접근할 권한이 없습니다.'
}

function Get-WDeliveryPath {
    param($Context, $Case, [string]$SourceId)
    Assert-WId $SourceId '조사 방식 ID'
    return Join-Path $Context.root ('Data\{0}\{1}\Deliveries\{2}\{3}\{4}' -f $Case.period, $Case.groupId, $SourceId, $Case.userId, $Case.caseId)
}

function Get-WCaseLock {
    param($Context, $Case)
    return Join-Path $Context.root ('Locks\{0}\{1}\{2}.lock' -f $Case.period, $Case.groupId, $Case.caseId)
}

function Assert-WDeliveryWrite {
    param($Context, $Case, [string]$SourceId, $Period, [switch]$ForRead)
    $user = Assert-WContext $Context
    $source = @(Get-WMasterRecords $Context.root 'Sources' | Where-Object { $_.sourceId -eq $SourceId }) | Select-Object -First 1
    if ($null -eq $source -or (-not $ForRead -and -not [bool](Get-WValue $source 'active' $true))) { throw '활성 조사 방식이 아닙니다.' }
    if ($source.type -eq 'General') { throw '선행기술조사는 품질 순위 비교 전용이며 검색 결과 입력 대상이 아닙니다.' }
    if (@((Get-WValue $Period 'sourceIds' @())) -notcontains $SourceId) { throw '이 운영월의 비교 조사 방식이 아닙니다.' }
    if (-not (Get-WMember $Period $Case.userId)) { throw '제출 심사관이 해당 월 참여자로 등록되어 있지 않습니다.' }
    if (Test-WRole $user 'admin') { return $user }
    if ((Test-WRole $user 'ai') -and @(Get-WAssignedSources $Period $user.userId $Case.groupId) -contains $SourceId) { return $user }
    throw '이 조사 방식의 검색 결과를 작성할 권한이 없습니다.'
}

function Get-Cases {
    param([Parameter(Mandatory)]$Context, [Parameter(Mandatory)][string]$Period, [ValidateSet('Mine','Assigned','All')][string]$Scope = 'Mine')
    $user = Assert-WContext $Context
    $settings = Get-Period $Context $Period
    if ($Scope -eq 'All' -and -not (Test-WRole $user 'admin')) { throw '전체 목록은 총괄관리자만 조회할 수 있습니다.' }
    if ($Scope -eq 'Assigned' -and -not ((Test-WRole $user 'ai') -or (Test-WRole $user 'admin'))) { throw '담당 그룹 목록 조회 권한이 없습니다.' }
    $base = Join-Path $Context.root ('Data\{0}' -f $Period)
    if (-not [System.IO.Directory]::Exists($base)) { return }
    $names = @{}
    foreach ($entry in @(Get-WMasterRecords $Context.root 'Users')) { $names[$entry.userId] = $entry.name }
    foreach ($group in Get-ChildItem -LiteralPath $base -Directory) {
        if ($Scope -eq 'Assigned' -and -not (Test-WRole $user 'admin') -and @(Get-WAssignedSources $settings $user.userId $group.Name).Count -eq 0) { continue }
        $submissions = Join-Path $group.FullName 'Submissions'
        if (-not [System.IO.Directory]::Exists($submissions)) { continue }
        foreach ($owner in Get-ChildItem -LiteralPath $submissions -Directory) {
            if ($Scope -eq 'Mine' -and $owner.Name -ne $user.userId) { continue }
            $selectionSnapshot = Read-Revision -Path (Get-WSelectionPath $Context $Period $owner.Name)
            foreach ($folder in Get-ChildItem -LiteralPath $owner.FullName -Directory) {
                $envelope = Read-Revision -Path $folder.FullName
                if ($null -ne $envelope) {
                    $record = Convert-WRecord $envelope $folder.FullName
                    if ($record.period -ne $Period -or $record.groupId -ne $group.Name -or $record.userId -ne $owner.Name -or $record.caseId -ne $folder.Name) { throw "제출정보 참조가 손상되었습니다: $($folder.FullName)" }
                    if (-not (Test-WSelectionPublished $Context $record -Selection $selectionSnapshot -UseSnapshot)) { continue }
                    $name = [string](Get-WValue $record 'examinerName' '')
                    if (-not $name -and $names.ContainsKey($record.userId)) { $name = $names[$record.userId] }
                    Set-WValue $record 'examinerName' $name
                    $record
                }
            }
        }
    }
}

function Convert-WApplicationNumber {
    param([string]$Value)
    $text = $Value.Trim() -replace '\s',''
    if ($text -match '^(\d{2})((?:19|20)\d{2})(\d{7})$') { return '{0}-{1}-{2}' -f $Matches[1], $Matches[2], $Matches[3] }
    if ($text -notmatch '^\d{2}-(?:19|20)\d{2}-\d{7}$') { throw '출원번호는 10-2026-0000001 형식 또는 숫자 13자리로 입력하십시오.' }
    return $text
}

function Get-WSelectionPath {
    param($Context, [string]$Period, [string]$UserId)
    Assert-WPeriod $Period
    Assert-WId $UserId
    return Join-Path $Context.root ('Data\{0}\Selections\{1}' -f $Period, $UserId)
}

function Test-WSelectionPublished {
    param($Context, $Case, $Selection=$null, [switch]$UseSnapshot)
    $request = [string](Get-WValue $Case 'selectionRequestId' '')
    if (-not $request) { return $true } # Already submitted by earlier versions.
    if (-not $UseSnapshot) {$Selection = Read-Revision -Path (Get-WSelectionPath $Context $Case.period $Case.userId)}
    return $null -ne $selection -and $selection.data.status -eq 'Confirmed' -and $selection.data.groupId -eq $Case.groupId -and $selection.requestId -eq $request -and $Case.caseId -in @($selection.data.caseIds)
}

function Test-WOtherSubmission {
    param($Context, [string]$Period, [string]$UserId, [string]$Number)
    $base=Join-Path $Context.root ('Data\'+$Period)
    if (-not [IO.Directory]::Exists($base)) {return $false}
    foreach ($group in Get-ChildItem -LiteralPath $base -Directory) {
        $sub=Join-Path $group.FullName Submissions
        if (-not [IO.Directory]::Exists($sub)) {continue}
        foreach ($owner in Get-ChildItem -LiteralPath $sub -Directory) {
            if ($owner.Name -eq $UserId) {continue}
            try {foreach ($folder in Get-ChildItem -LiteralPath $owner.FullName -Directory -ErrorAction Stop) {
                $other=Read-Revision -Path $folder.FullName
                if ($other -and $other.data.applicationNumber -eq $Number -and $other.data.status -ne 'Cancelled' -and (Test-WSelectionPublished $Context $other.data)) {return $true}
            }} catch [UnauthorizedAccessException] {}
        }
    }
    return $false
}

function Get-TargetSelection {
    param([Parameter(Mandatory)]$Context, [Parameter(Mandatory)][string]$Period)
    $user = Assert-WContext $Context
    if (-not (Test-WRole $user 'examiner')) { throw '심사관 역할이 필요합니다.' }
    $null = Get-Period $Context $Period
    $path = Get-WSelectionPath $Context $Period $user.userId
    $saved = Read-Revision -Path $path
    if ($saved) { return Convert-WRecord $saved $path }
    $existing = @(Get-Cases $Context $Period Mine | Where-Object { $_.status -ne 'Cancelled' })
    return [pscustomobject]@{period=$Period;userId=$user.userId;status=$(if ($existing.Count) {'Confirmed'} else {'NotStarted'});rows=@();caseIds=@($existing | ForEach-Object {$_.caseId});version=0;path=$path}
}

function Save-TargetSelection {
    param([Parameter(Mandatory)]$Context, [Parameter(Mandatory)][string]$Period,
        [Parameter(Mandatory)][int]$ExpectedVersion, [Parameter(Mandatory)][string]$RequestId,
        [Parameter(Mandatory)][ValidateSet('Draft','Confirmed')][string]$Status, [object[]]$Rows=@())
    $user = Assert-WContext $Context
    if (-not (Test-WRole $user 'examiner')) { throw '심사관 역할이 필요합니다.' }
    Assert-WPeriod $Period
    Assert-WId $RequestId '저장 요청 ID'
    # Drafts may contain unfinished/invalid numbers. Confirmation validates the entire list before publishing any case.
    $clean = @(foreach ($row in $Rows) {
        [pscustomobject]@{applicationNumber=(Convert-WOptionalText (Get-WValue $row 'applicationNumber' '') '출원번호' 100);cpc=(Convert-WOptionalText (Get-WValue $row 'cpc' '') 'CPC' 500)}
    })
    $fingerprint = Get-WFingerprint @{operation='SaveTargetSelection';period=$Period;userId=$user.userId;expectedVersion=$ExpectedVersion;status=$Status;rows=$clean}
    $lock = Join-Path $Context.root ('Locks\{0}\registration-{1}.lock' -f $Period, $user.userId)
    return Invoke-WPeriodReadLock $Context $Period {
      Invoke-FileLock -Path $lock -Action {
        $user = Assert-WContext $Context
        if (-not (Test-WRole $user 'examiner')) { throw '심사관 역할이 필요합니다.' }
        $path = Get-WSelectionPath $Context $Period $user.userId
        $previous = Read-Revision -Path $path
        $oldRequest = @(Get-Revisions -Path $path | Where-Object {$_.requestId -eq $RequestId}) | Select-Object -First 1
        if ($oldRequest) { Assert-WRequest $oldRequest $fingerprint $user.userId; return Convert-WRecord $oldRequest $path }
        $currentVersion = if ($previous) {[int]$previous.version} else {0}
        if ($currentVersion -ne $ExpectedVersion) { throw '다른 창에서 대상 건이 변경되었습니다. 새로 고침 후 다시 확인하십시오.' }
        if ($previous -and $previous.data.status -eq 'Confirmed') { throw '이미 대상 건이 확정된 운영월입니다.' }
        $settings = Get-Period $Context $Period
        Assert-WOpen $settings Submission
        $member = Get-WMember $settings $user.userId
        if (-not $member) { throw '해당 월 참여 심사관이 아닙니다.' }
        $existing = @(Get-Cases $Context $Period Mine | Where-Object {$_.status -ne 'Cancelled'})
        if ($existing.Count) { throw '이미 대상 건이 확정된 운영월입니다.' }
        $caseIds = @()
        if ($Status -eq 'Confirmed') {
            if (-not $clean.Count) { throw '확정할 대상 건을 한 건 이상 입력하십시오.' }
            $limit = [int](Get-WValue $settings 'submissionLimit' 0)
            if ($limit -gt 0 -and $clean.Count -gt $limit) { throw '월별 제출 건수 상한을 초과했습니다.' }
            $seen = @{}
            foreach ($row in $clean) {
                $row.applicationNumber = Convert-WApplicationNumber $row.applicationNumber
                if ($seen.ContainsKey($row.applicationNumber)) { throw '대상 건 목록에 중복된 출원번호가 있습니다.' }
                $seen[$row.applicationNumber] = $true
            }
            $index = 0
            foreach ($row in $clean) {
                $caseId = 'C' + (Get-WFingerprint @{userId=$user.userId;period=$Period;request=$RequestId;index=$index}).Substring(0,32)
                $index++
                $data = [pscustomobject]@{caseId=$caseId;period=$Period;groupId=$member.groupId;userId=$user.userId;examinerName=$user.name;
                    examinerBureau=[string](Get-WValue $user 'bureau' '');examinerDivision=[string](Get-WValue $user 'division' '');
                    applicationNumber=$row.applicationNumber;cpc=$row.cpc;status='Submitted';submittedAt=[datetimeoffset]::UtcNow.ToString('o');duplicateWarning=$false;
                    selectionRequestId=$RequestId;requestFingerprint=$fingerprint}
                $casePath = Get-WCasePath $Context $data
                $data.duplicateWarning=Test-WOtherSubmission $Context $Period $user.userId $row.applicationNumber
                $prepared = Read-Revision -Path $casePath
                if ($prepared) { Assert-WRequest $prepared $fingerprint $user.userId }
                else { $null = Write-Revision -Path $casePath -ExpectedVersion 0 -RequestId $RequestId -ActorId $user.userId -Data $data -Kind Submission -TargetId $caseId }
                $caseIds += $caseId
            }
        }
        # This one immutable revision publishes the entire set. Interrupted preparation remains invisible and retryable.
        $data = [pscustomobject]@{period=$Period;userId=$user.userId;groupId=$member.groupId;status=$Status;rows=$clean;caseIds=$caseIds;requestFingerprint=$fingerprint}
        $saved = Write-Revision -Path $path -ExpectedVersion $ExpectedVersion -RequestId $RequestId -ActorId $user.userId -Data $data -Kind TargetSelection -TargetId $user.userId
        return Convert-WRecord $saved $path
      }
    }
}

function New-Submission {
    param([Parameter(Mandatory)]$Context, [Parameter(Mandatory)][string]$Period, [Parameter(Mandatory)][string]$ApplicationNumber,
        [Parameter(Mandatory)][string]$RequestId, [string]$Reason = '', [string]$Cpc = '')
    $user = Assert-WContext $Context
    if (-not (Test-WRole $user 'examiner')) { throw '심사관 역할이 필요합니다.' }
    Assert-WPeriod $Period
    Assert-WId $RequestId '저장 요청 ID'
    $number = Convert-WApplicationNumber $ApplicationNumber
    $cpcText = Convert-WOptionalText $Cpc 'CPC' 500
    $requestData = @{operation='NewSubmission';period=$Period;applicationNumber=$number;reason=$Reason;userId=$user.userId}
    # Omit blank CPC from the fingerprint so pre-CPC requests can still be retried.
    if ($cpcText) { $requestData.cpc = $cpcText }
    $fingerprint = Get-WFingerprint $requestData
    $lock = Join-Path $Context.root ('Locks\{0}\registration-{1}.lock' -f $Period, $user.userId)
    return Invoke-WPeriodReadLock $Context $Period {
      Invoke-FileLock -Path $lock -Action {
        $user = Assert-WContext $Context
        $settings = Get-Period $Context $Period
        $member = Get-WMember $settings $user.userId
        if ($null -eq $member) { throw '해당 월 참여 심사관이 아닙니다.' }
        $existing = @(Get-Cases $Context $Period 'Mine')
        foreach ($item in $existing) {
            foreach ($rev in @(Get-Revisions -Path $item.path)) {
                if ($rev.requestId -eq $RequestId) {
                    Assert-WRequest $rev $fingerprint $user.userId
                    return Convert-WRecord $rev $item.path
                }
            }
        }
        if (Read-Revision -Path (Get-WSelectionPath $Context $Period $user.userId)) { throw '월별 대상 건 선정 화면에서 임시저장 또는 확정하십시오.' }
        Assert-WOpen $settings 'Submission'
        if (@($existing | Where-Object { $_.applicationNumber -eq $number -and $_.status -ne 'Cancelled' }).Count -gt 0) { throw '본인이 해당 월에 이미 제출한 출원번호입니다.' }
        $limit = [int](Get-WValue $settings 'submissionLimit' 0)
        if ($limit -gt 0 -and @($existing | Where-Object { $_.status -ne 'Cancelled' }).Count -ge $limit) { throw '월별 제출 건수 상한에 도달했습니다.' }
        $caseId = 'C' + [guid]::NewGuid().ToString('N')
        $data = [pscustomobject]@{ caseId = $caseId; period = $Period; groupId = $member.groupId; userId = $user.userId; examinerName = $user.name;
            examinerBureau = [string](Get-WValue $user 'bureau' ''); examinerDivision = [string](Get-WValue $user 'division' '');
            applicationNumber = $number; cpc = $cpcText; status = 'Submitted'; submittedAt = [datetimeoffset]::UtcNow.ToString('o'); duplicateWarning = $false; requestFingerprint=$fingerprint }
        # Cross-examiner duplicates are advisory; the per-user lock prevents duplicate registration by two own windows.
        $base = Join-Path $Context.root ('Data\{0}' -f $Period)
        if ([System.IO.Directory]::Exists($base)) {
            foreach ($group in Get-ChildItem -LiteralPath $base -Directory) {
                $sub = Join-Path $group.FullName 'Submissions'
                if (-not [System.IO.Directory]::Exists($sub)) { continue }
                foreach ($owner in Get-ChildItem -LiteralPath $sub -Directory) {
                    if ($owner.Name -eq $user.userId) { continue }
                    try {
                        foreach ($folder in Get-ChildItem -LiteralPath $owner.FullName -Directory -ErrorAction Stop) {
                            $other = Read-Revision -Path $folder.FullName
                            if ($null -ne $other -and $other.data.applicationNumber -eq $number -and $other.data.status -ne 'Cancelled' -and (Test-WSelectionPublished $Context $other.data)) { $data.duplicateWarning = $true }
                        }
                    } catch [System.UnauthorizedAccessException] { }
                }
            }
        }
        $path = Get-WCasePath $Context $data
        $saved = Write-Revision -Path $path -ExpectedVersion 0 -RequestId $RequestId -ActorId $user.userId -Reason $Reason -Data $data -Kind 'Submission' -TargetId $caseId
        return Convert-WRecord $saved $path
      }
    }
}

function Cancel-Submission {
    param([Parameter(Mandatory)]$Context, [Parameter(Mandatory)]$Case, [Parameter(Mandatory)][int]$ExpectedVersion,
        [Parameter(Mandatory)][string]$RequestId, [string]$Reason = '')
    $null = Assert-WContext $Context
    $actual = Get-WCase $Context $Case
    $fingerprint = Get-WFingerprint @{operation='CancelSubmission';caseId=$actual.caseId;expectedVersion=$ExpectedVersion;reason=$Reason}
    return Invoke-WPeriodReadLock $Context $actual.period {
      Invoke-FileLock -Path (Get-WCaseLock $Context $actual) -Action {
        $user = Assert-WContext $Context
        $actual = Get-WCase $Context $Case
        if ($actual.userId -ne $user.userId -or -not (Test-WRole $user 'examiner')) { throw '본인의 제출 건만 취소할 수 있습니다.' }
        $path = Get-WCasePath $Context $actual
        $oldRequest = @(Get-Revisions -Path $path | Where-Object { $_.requestId -eq $RequestId }) | Select-Object -First 1
        if ($oldRequest) { Assert-WRequest $oldRequest $fingerprint $user.userId; return Convert-WRecord $oldRequest $path }
        Assert-WOpen (Get-Period $Context $actual.period) 'Submission'
        if ($actual.status -eq 'Cancelled') { throw '이미 취소된 제출 건입니다.' }
        $data = [pscustomobject]@{ caseId = $actual.caseId; period = $actual.period; groupId = $actual.groupId; userId = $actual.userId; examinerName = $actual.examinerName;
            examinerBureau = $actual.examinerBureau; examinerDivision = $actual.examinerDivision;
            applicationNumber = $actual.applicationNumber; cpc = $actual.cpc; status = 'Cancelled'; submittedAt = $actual.submittedAt; cancelledAt = [datetimeoffset]::UtcNow.ToString('o'); cancellationReason = $Reason; requestFingerprint=$fingerprint }
        $saved = Write-Revision -Path $path -ExpectedVersion $ExpectedVersion -RequestId $RequestId -ActorId $user.userId -Reason $Reason -Data $data -Kind 'Submission' -TargetId $actual.caseId
        return Convert-WRecord $saved $path
      }
    }
}

function Get-WPublished {
    param($Context, $Case, [string]$SourceId)
    $path = Get-WDeliveryPath $Context $Case $SourceId
    $published = @(Get-Revisions -Path $path | Where-Object { (Get-WValue $_.data 'status') -eq 'Published' } | Sort-Object { [int]$_.version } -Descending) | Select-Object -First 1
    if ($null -ne $published) {
        Set-WValue $published 'path' $path
        Set-WValue $published 'sourceId' $SourceId
        Set-WValue $published 'sourceName' (Get-WValue $published.data 'sourceName' $SourceId)
    }
    return $published
}

function Get-Deliveries {
    param([Parameter(Mandatory)]$Context, [Parameter(Mandatory)]$Case)
    $actual = Get-WCase $Context $Case
    $user = Assert-WCaseRead $Context $actual
    $settings = Get-Period $Context $actual.period
    $isOwner = $actual.userId -eq $user.userId -and (Test-WRole $user 'examiner')
    $assigned = @(Get-WAssignedSources $settings $user.userId $actual.groupId)
    $aiIds = @(Get-WMasterRecords $Context.root 'Sources' | Where-Object {$_.type -eq 'AI'} | ForEach-Object {$_.sourceId})
    foreach ($sourceId in @((Get-WValue $settings 'sourceIds' @()))) {
        if ($sourceId -notin $aiIds) { continue }
        if (-not $isOwner -and -not (Test-WRole $user 'admin') -and $assigned -notcontains $sourceId) { continue }
        $published = Get-WPublished $Context $actual $sourceId
        if ($null -ne $published) { $published }
    }
}

function Get-DeliveryForEdit {
    param([Parameter(Mandatory)]$Context, [Parameter(Mandatory)]$Case, [Parameter(Mandatory)][string]$SourceId)
    $actual = Get-WCase $Context $Case
    $period = Get-Period $Context $actual.period
    $null = Assert-WDeliveryWrite $Context $actual $SourceId $period -ForRead
    $path = Get-WDeliveryPath $Context $actual $SourceId
    $record = Read-Revision -Path $path
    if ($record) { Set-WValue $record 'path' $path }
    return $record
}

function Get-DeliveryHistory {
    param([Parameter(Mandatory)]$Context, [Parameter(Mandatory)]$Case, [Parameter(Mandatory)][string]$SourceId)
    $null = Get-DeliveryForEdit $Context $Case $SourceId
    return Get-Revisions -Path (Get-WDeliveryPath $Context $Case $SourceId)
}

function Get-WAttachmentBase {
    param($Context, $Case, [string]$SourceId)
    Assert-WId $SourceId
    return Join-Path $Context.root ('Data\{0}\{1}\Attachments\{2}\{3}\{4}' -f $Case.period, $Case.groupId, $SourceId, $Case.userId, $Case.caseId)
}

function Resolve-DeliveryAttachment {
    param([Parameter(Mandatory)]$Context, [Parameter(Mandatory)]$Case, [Parameter(Mandatory)][string]$SourceId, [Parameter(Mandatory)]$Attachment)
    $actual = Get-WCase $Context $Case
    $null = Assert-WCaseRead $Context $actual
    $relative = [string](Get-WValue $Attachment 'relativePath')
    if ([string]::IsNullOrWhiteSpace($relative) -or [System.IO.Path]::IsPathRooted($relative)) { throw '첨부파일 상대 경로가 올바르지 않습니다.' }
    $base = [System.IO.Path]::GetFullPath((Get-WAttachmentBase $Context $actual $SourceId)).TrimEnd('\') + '\'
    $full = [System.IO.Path]::GetFullPath((Join-Path $Context.root $relative))
    if (-not $full.StartsWith($base, [System.StringComparison]::OrdinalIgnoreCase)) { throw '첨부파일이 해당 검색 결과의 저장 영역을 벗어났습니다.' }
    if (-not [System.IO.File]::Exists($full)) { throw '첨부파일을 찾을 수 없습니다.' }
    return $full
}

function Import-DeliveryAttachment {
    param([Parameter(Mandatory)]$Context, [Parameter(Mandatory)]$Case, [Parameter(Mandatory)][string]$SourceId, [Parameter(Mandatory)][string]$Path)
    $actual = Get-WCase $Context $Case
    $attachmentInputPath = $Path
    return Invoke-WPeriodReadLock $Context $actual.period {
      Invoke-FileLock -Path (Get-WCaseLock $Context $actual) -Action {
    $actual = Get-WCase $Context $Case
    $settings = Get-Period $Context $actual.period
    $null = Assert-WDeliveryWrite $Context $actual $SourceId $settings
    Assert-WOpen $settings 'Delivery'
    if ($actual.status -eq 'Cancelled') { throw '취소된 제출 건에 첨부할 수 없습니다.' }
    $source = Get-Item -LiteralPath $attachmentInputPath -ErrorAction Stop
    if ($source.PSIsContainer) { throw '파일만 첨부할 수 있습니다.' }
    $base = Get-WAttachmentBase $Context $actual $SourceId
    $null = [System.IO.Directory]::CreateDirectory($base)
    $id = [guid]::NewGuid().ToString('N')
    $extension = [System.IO.Path]::GetExtension($source.Name)
    $destination = Join-Path $base ($id + $extension)
    [System.IO.File]::Copy($source.FullName, $destination, $false)
    $rootPrefix = [System.IO.Path]::GetFullPath($Context.root).TrimEnd('\') + '\'
    $relative = $destination.Substring($rootPrefix.Length)
    return [pscustomobject]@{ attachmentId = $id; fileName = $source.Name; relativePath = $relative; size = $source.Length; sha256 = (Get-FileHash -LiteralPath $destination -Algorithm SHA256).Hash.ToLowerInvariant() }
      }
    }
}

function Convert-WDocuments {
    param($Context, $Case, [string]$SourceId, [object[]]$Documents, $Existing)
    if (@($Documents).Count -gt 7) { throw '조사 방식별 선행기술은 최대 7건입니다.' }
    $ids = @{}
    foreach ($item in @($Documents)) {
        if ($null -eq $item) { continue }
        $id = [string](Get-WValue $item 'documentId' '')
        if (-not $id) { $id = 'D' + [guid]::NewGuid().ToString('N') }
        else { Assert-WId $id '문헌 ID' }
        if ($ids.ContainsKey($id)) { throw '문헌 ID가 중복되었습니다.' }
        $ids[$id] = $true
        $oldDocument = @((Get-WValue (Get-WValue $Existing 'data') 'documents' @()) | Where-Object { $_.documentId -eq $id }) | Select-Object -First 1
        $type = [string](Get-WValue $item 'type' (Get-WValue $item 'documentType' (Get-WValue $oldDocument 'type' 'Patent')))
        $identifier = [string](Get-WValue $item 'identifier' '')
        $title = [string](Get-WValue $item 'title' (Get-WValue $oldDocument 'title' ''))
        if ([string]::IsNullOrWhiteSpace($identifier)) { throw '선행기술 문헌 출원번호 또는 식별정보를 입력하십시오.' }
        $attachments = @()
        foreach ($attachment in @((Get-WValue $item 'attachments' @()))) {
            $null = Resolve-DeliveryAttachment $Context $Case $SourceId $attachment
            $attachments += [pscustomobject]@{ attachmentId = [string](Get-WValue $attachment 'attachmentId'); fileName = [string](Get-WValue $attachment 'fileName');
                relativePath = [string](Get-WValue $attachment 'relativePath'); size = [long](Get-WValue $attachment 'size' 0); sha256 = [string](Get-WValue $attachment 'sha256') }
        }
        [pscustomobject]@{ documentId = $id; type = $type; identifier = $identifier.Trim(); title = $title.Trim(); description = [string](Get-WValue $item 'description' (Get-WValue $oldDocument 'description' '')); attachments = @($attachments) }
    }
}

function Save-Delivery {
    param([Parameter(Mandatory)]$Context, [Parameter(Mandatory)]$Case, [Parameter(Mandatory)][string]$SourceId,
        [Parameter(Mandatory)][int]$ExpectedVersion, [Parameter(Mandatory)][string]$RequestId,
        [Parameter(Mandatory)][ValidateSet('Draft','Published')][string]$Status, [object[]]$Documents = @(), [bool]$NoResults = $false,
        [string]$NoResultsReason = '', [string]$Reason = '', [string]$ModelName = '', [string]$ModelVersion = '')
    Assert-WId $RequestId '저장 요청 ID'
    $actual = Get-WCase $Context $Case
    $null = Assert-WDeliveryWrite $Context $actual $SourceId (Get-Period $Context $actual.period)
    $fingerprint = Get-WFingerprint @{operation='SaveDelivery';caseId=$actual.caseId;sourceId=$SourceId;expectedVersion=$ExpectedVersion;status=$Status;documents=@($Documents);noResults=$NoResults;noResultsReason=$NoResultsReason;reason=$Reason;modelName=$ModelName;modelVersion=$ModelVersion}
    return Invoke-WPeriodReadLock $Context $actual.period {
      Invoke-FileLock -Path (Get-WCaseLock $Context $actual) -Action {
        $actual = Get-WCase $Context $Case
        $settings = Get-Period $Context $actual.period
        $user = Assert-WDeliveryWrite $Context $actual $SourceId $settings
        $path = Get-WDeliveryPath $Context $actual $SourceId
        $oldRequest = @(Get-Revisions -Path $path | Where-Object { $_.requestId -eq $RequestId }) | Select-Object -First 1
        if ($oldRequest) { Assert-WRequest $oldRequest $fingerprint $user.userId; return $oldRequest }
        Assert-WOpen $settings 'Delivery'
        if ($actual.status -eq 'Cancelled') { throw '취소된 제출 건의 검색 결과를 수정할 수 없습니다.' }
        $current = Read-Revision -Path $path
        $version = 0
        if ($null -ne $current) { $version = [int]$current.version }
        if ($version -ne $ExpectedVersion) { throw '다른 창에서 검색 결과가 변경되었습니다. 최신 내용을 다시 불러오십시오.' }
        $docs = @(Convert-WDocuments $Context $actual $SourceId $Documents $current)
        if ($NoResults -and $docs.Count -gt 0) { throw '결과 없음과 문헌 목록을 동시에 등록할 수 없습니다.' }
        if ($NoResults -and [string]::IsNullOrWhiteSpace($NoResultsReason)) { throw '결과 없음 사유를 입력하십시오.' }
        if ($Status -eq 'Published' -and -not $NoResults -and $docs.Count -eq 0) { throw '저장할 선행기술 문헌 출원번호를 한 건 이상 입력하십시오.' }
        $source = @(Get-WMasterRecords $Context.root 'Sources' | Where-Object { $_.sourceId -eq $SourceId }) | Select-Object -First 1
        $data = [pscustomobject]@{ caseId = $actual.caseId; period = $actual.period; groupId = $actual.groupId; userId = $actual.userId;
            sourceId = $SourceId; sourceName = $source.name; status = $Status; documents = @($docs); noResults = $NoResults; noResultsReason = $NoResultsReason;
            authorId = $user.userId; authorName = $user.name; modelName = $ModelName; modelVersion = $ModelVersion; requestFingerprint=$fingerprint }
        return Write-Revision -Path $path -ExpectedVersion $ExpectedVersion -RequestId $RequestId -ActorId $user.userId -Reason $Reason -Data $data -Kind 'Delivery' -TargetId ($actual.caseId + '-' + $SourceId)
      }
    }
}

function Get-WEvaluationPath {
    param($Context, $Case)
    return Join-Path $Context.root ('Data\{0}\{1}\Evaluations\{2}\{3}' -f $Case.period, $Case.groupId, $Case.userId, $Case.caseId)
}

function Assert-WEvaluationOwner {
    param($Context, $Case)
    $user = Assert-WContext $Context
    if ($Case.userId -ne $user.userId -or -not (Test-WRole $user 'examiner')) { throw '평가 작성은 해당 제출 심사관만 할 수 있습니다.' }
    $period = Get-Period $Context $Case.period
    if (-not (Get-WMember $period $user.userId)) { throw '해당 월 참여 심사관이 아닙니다.' }
    return $user
}

function Get-Evaluation {
    param([Parameter(Mandatory)]$Context, [Parameter(Mandatory)]$Case)
    $user = Assert-WContext $Context
    $actual = Get-WCase $Context $Case
    if (-not (Test-WRole $user 'admin') -and ($actual.userId -ne $user.userId -or -not (Test-WRole $user 'examiner'))) { throw '평가자료 조회 권한이 없습니다.' }
    $path = Get-WEvaluationPath $Context $actual
    $record = Read-Revision -Path $path
    if ($null -eq $record) { return $null }
    return Add-WEvaluationState $Context $actual $record $path
}

function Get-WComparisonSources {
    param($Context, $Settings)
    Get-WMasterRecords $Context.root 'Sources' | Where-Object {
        $_.sourceId -in @($Settings.sourceIds) -or ($_.type -eq 'General' -and [bool](Get-WValue $_ 'active' $true))
    }
}

function Add-WEvaluationState {
    param($Context, $Case, $Record, [string]$Path)
    $settings = Get-Period $Context $Case.period
    $sources = @(Get-WComparisonSources $Context $settings)
    Set-WValue $Record 'rankOnlySources' @($sources | Where-Object {$_.type -eq 'General'} | Select-Object sourceId,name)
    $stale = @()
    foreach ($sourceId in @($sources | Where-Object {$_.type -eq 'AI'} | ForEach-Object {$_.sourceId})) {
        $published = Get-WPublished $Context $Case $sourceId
        $currentVersion = 0
        if ($published) { $currentVersion = [int]$published.version }
        $ratedVersion = [int](Get-WValue (Get-WValue $Record.data 'deliveryVersions') $sourceId 0)
        if ($currentVersion -ne $ratedVersion) { $stale += $sourceId }
    }
    Set-WValue $Record 'path' $Path
    Set-WValue $Record 'needsReconfirmation' ($stale.Count -gt 0)
    Set-WValue $Record 'staleSourceIds' @($stale)
    return $Record
}

function Convert-WRatings {
    param([object[]]$Ratings, $Published)
    $used = @{}
    foreach ($rating in @($Ratings)) {
        $sourceId = [string](Get-WValue $rating 'sourceId')
        $documentId = [string](Get-WValue $rating 'documentId')
        if (-not $Published.ContainsKey($sourceId)) { throw '검색 결과가 입력되지 않은 조사 방식의 문헌을 평가할 수 없습니다.' }
        $documents = @((Get-WValue $Published[$sourceId].data 'documents' @()))
        if (@($documents | Where-Object { $_.documentId -eq $documentId }).Count -ne 1) { throw '현재 검색 결과에 없는 문헌 평가입니다.' }
        $key = $sourceId + '/' + $documentId
        if ($used.ContainsKey($key)) { throw '동일 문헌 평가가 중복되었습니다.' }
        $used[$key] = $true
        $score = Get-WValue $rating 'score'
        if ($null -ne $score -and [string]$score -ne '') {
            if ([string]$score -notmatch '^[0-5]$') { throw '점수는 0~5의 정수 또는 미평가여야 합니다.' }
            $score = [int]$score
        } else { $score = $null }
        $cited = Get-WValue $rating 'cited'
        if ($null -ne $cited -and $cited -isnot [bool]) { throw '인용 여부는 선택하지 않음, true, false 중 하나여야 합니다.' }
        [pscustomobject]@{ sourceId = $sourceId; documentId = $documentId; score = $score; cited = $cited; note = [string](Get-WValue $rating 'note' '') }
    }
}

function Convert-WRankings {
    param([object[]]$Rankings, [string[]]$SourceIds, $Published, [string[]]$RankOnlySourceIds=@())
    $bySource = @{}
    $byRank = @{}
    foreach ($ranking in @($Rankings)) {
        $sourceId = [string](Get-WValue $ranking 'sourceId')
        if ($SourceIds -notcontains $sourceId) { throw '월별 비교 조사 방식에 없는 순위입니다.' }
        if ($bySource.ContainsKey($sourceId)) { throw '조사 방식의 품질 순위가 중복되었습니다.' }
        $status = [string](Get-WValue $ranking 'status')
        if ($status -notin @('Ranked','Outside','Excluded','NoResults','NotDelivered','Incomparable')) { throw '품질 순위 상태가 올바르지 않습니다.' }
        $rank = Get-WValue $ranking 'rank'
        if ($status -eq 'Ranked') {
            if ([string]$rank -notmatch '^[1-3]$') { throw '품질 순위는 1, 2, 3위로 입력하십시오.' }
            $rank = [int]$rank
            if ($byRank.ContainsKey([string]$rank)) { throw '동순위는 허용하지 않습니다.' }
            $byRank[[string]$rank] = $true
        } else {
            if ($null -ne $rank -and [string]$rank -ne '') { throw '순위가 있는 경우 Ranked 상태로 저장하십시오.' }
            $rank = $null
        }
        $rankOnly = $sourceId -in $RankOnlySourceIds
        $hasDelivery = $rankOnly -or $Published.ContainsKey($sourceId)
        $noResults = -not $rankOnly -and $hasDelivery -and [bool](Get-WValue $Published[$sourceId].data 'noResults' $false)
        if (-not $hasDelivery -and $status -ne 'NotDelivered') { throw '검색 결과 입력 전인 조사 방식은 NotDelivered 상태여야 합니다.' }
        if ($hasDelivery -and -not $noResults -and $status -in @('NotDelivered','NoResults')) { throw '검색 결과 입력 상태와 품질 순위 상태가 일치하지 않습니다.' }
        if ($noResults -and $status -notin @('NoResults','Excluded')) { throw '결과 없음 조사 방식에는 순위를 부여할 수 없습니다.' }
        $bySource[$sourceId] = [pscustomobject]@{ sourceId = $sourceId; rank = $rank; status = $status }
    }
    # Unentered results remain explicitly unranked. This permits partial evaluation saves.
    foreach ($sourceId in $SourceIds) {
        if (-not $bySource.ContainsKey($sourceId)) {
            $status = 'NotDelivered'
            if ($sourceId -in $RankOnlySourceIds) { $status = 'Outside' }
            elseif ($Published.ContainsKey($sourceId)) {
                if ([bool](Get-WValue $Published[$sourceId].data 'noResults' $false)) { $status = 'NoResults' } else { $status = 'Outside' }
            }
            $bySource[$sourceId] = [pscustomobject]@{ sourceId = $sourceId; rank = $null; status = $status }
        }
    }
    $comparable = @($bySource.Values | Where-Object { $_.status -in @('Ranked','Outside','Incomparable') })
    if ($comparable.Count -eq 1) {
        if ($comparable[0].status -eq 'Ranked') { throw '비교 가능한 결과가 한 개이면 비교 불가로 입력하십시오.' }
        $comparable[0].status = 'Incomparable'
    } elseif (@($comparable | Where-Object { $_.status -eq 'Incomparable' }).Count -gt 0) { throw '비교 가능한 결과가 두 개 이상이면 비교 불가를 선택할 수 없습니다.' }
    if ($byRank.Count -gt 0) {
        for ($n = 1; $n -le $byRank.Count; $n++) { if (-not $byRank.ContainsKey([string]$n)) { throw '품질 순위는 1위부터 연속해서 입력하십시오.' } }
        if ($byRank.Count -gt [Math]::Min(3, $comparable.Count)) { throw '비교 가능한 결과 수보다 많은 순위를 입력할 수 없습니다.' }
    }
    foreach ($sourceId in $SourceIds) { $bySource[$sourceId] }
}

function Save-Evaluation {
    param([Parameter(Mandatory)]$Context, [Parameter(Mandatory)]$Case, [Parameter(Mandatory)][int]$ExpectedVersion,
        [Parameter(Mandatory)][string]$RequestId, [Parameter(Mandatory)]$DeliveryVersions,
        [object[]]$DocumentRatings = @(), [object[]]$Rankings = @(), [string]$Reason = '')
    Assert-WId $RequestId '저장 요청 ID'
    $actual = Get-WCase $Context $Case
    $null = Assert-WEvaluationOwner $Context $actual
    $fingerprint = Get-WFingerprint @{operation='SaveEvaluation';caseId=$actual.caseId;expectedVersion=$ExpectedVersion;deliveryVersions=$DeliveryVersions;documentRatings=@($DocumentRatings);rankings=@($Rankings);reason=$Reason}
    return Invoke-WPeriodReadLock $Context $actual.period {
      Invoke-FileLock -Path (Get-WCaseLock $Context $actual) -Action {
        $actual = Get-WCase $Context $Case
        $user = Assert-WEvaluationOwner $Context $actual
        $path = Get-WEvaluationPath $Context $actual
        $oldRequest = @(Get-Revisions -Path $path | Where-Object { $_.requestId -eq $RequestId }) | Select-Object -First 1
        if ($oldRequest) { Assert-WRequest $oldRequest $fingerprint $user.userId; return Add-WEvaluationState $Context $actual $oldRequest $path }
        $settings = Get-Period $Context $actual.period
        Assert-WOpen $settings 'Evaluation'
        if ($actual.status -eq 'Cancelled') { throw '취소된 제출 건은 평가할 수 없습니다.' }
        $sources = @(Get-WComparisonSources $Context $settings)
        $sourceIds = @($sources | ForEach-Object {$_.sourceId})
        $rankOnlyIds = @($sources | Where-Object {$_.type -eq 'General'} | ForEach-Object {$_.sourceId})
        $deliveryIds = @($sources | Where-Object {$_.type -eq 'AI'} | ForEach-Object {$_.sourceId})
        $published = @{}
        $versions = [ordered]@{}
        foreach ($sourceId in $deliveryIds) {
            $latest = Get-WPublished $Context $actual $sourceId
            $currentVersion = 0
            if ($latest) { $published[$sourceId] = $latest; $currentVersion = [int]$latest.version }
            $expected = [int](Get-WValue $DeliveryVersions $sourceId 0)
            if ($currentVersion -ne $expected) { throw "$sourceId 검색 결과가 변경되었습니다. 최신 검색 결과를 확인한 뒤 평가하십시오." }
            $versions[$sourceId] = $currentVersion
        }
        if ($published.Count -eq 0) { throw '평가할 검색 결과가 없습니다.' }
        $ratings = @(Convert-WRatings $DocumentRatings $published)
        $ranks = @(Convert-WRankings $Rankings $sourceIds $published $rankOnlyIds)
        $completeDocuments = $true
        $documentCount = 0
        foreach ($sourceId in $published.Keys) {
            foreach ($document in @((Get-WValue $published[$sourceId].data 'documents' @()))) {
                $documentCount++
                $rated = @($ratings | Where-Object { $_.sourceId -eq $sourceId -and $_.documentId -eq $document.documentId }) | Select-Object -First 1
                if ($null -eq $rated -or $null -eq $rated.score -or $null -eq $rated.cited) { $completeDocuments = $false }
            }
        }
        $comparable = @($ranks | Where-Object { $_.status -in @('Ranked','Outside','Incomparable') }).Count
        $rankCount = @($ranks | Where-Object { $_.status -eq 'Ranked' }).Count
        $completeRanks = ($comparable -le 1 -or $rankCount -eq [Math]::Min(3,$comparable))
        $allDelivered = $published.Count -eq $deliveryIds.Count
        $data = [pscustomobject]@{ caseId = $actual.caseId; period = $actual.period; groupId = $actual.groupId; userId = $actual.userId;
            deliveryVersions = [pscustomobject]$versions; documentRatings = @($ratings); rankings = @($ranks);
            complete = ($completeDocuments -and $completeRanks -and $allDelivered); documentCount = $documentCount; requestFingerprint=$fingerprint }
        $null = Write-Revision -Path $path -ExpectedVersion $ExpectedVersion -RequestId $RequestId -ActorId $user.userId -Reason $Reason -Data $data -Kind 'Evaluation' -TargetId $actual.caseId
        return Get-Evaluation $Context $actual
      }
    }
}

function Get-CaseStatus {
    param([Parameter(Mandatory)]$Context, [Parameter(Mandatory)]$Case)
    $actual = Get-WCase $Context $Case
    $user = Assert-WCaseRead $Context $actual
    $deliveries = @(Get-Deliveries $Context $actual)
    $period = Get-Period $Context $actual.period
    $sourceMap = @{}
    foreach ($source in @(Get-Sources $Context)) { $sourceMap[$source.sourceId] = $source }
    $assigned = @(Get-WAssignedSources $period $user.userId $actual.groupId)
    $ownCase = $actual.userId -eq $user.userId -and (Test-WRole $user 'examiner')
    $evaluation = $null
    if ($actual.userId -eq $user.userId -or (Test-WRole $user 'admin')) { $evaluation = Get-Evaluation $Context $actual }
    $states = @()
    foreach ($sourceId in @((Get-WValue $period 'sourceIds' @()))) {
        if (-not $sourceMap.ContainsKey($sourceId) -or $sourceMap[$sourceId].type -ne 'AI') { continue }
        if (-not $ownCase -and -not (Test-WRole $user 'admin') -and $assigned -notcontains $sourceId) { continue }
        $delivery = @($deliveries | Where-Object { $_.data.sourceId -eq $sourceId }) | Select-Object -First 1
        $state = 'NotStarted'; $count = 0; $publishedVersion = 0; $editingVersion = 0
        if ($delivery) { $state = 'Published'; $count = @($delivery.data.documents).Count; $publishedVersion = $delivery.version; if ($delivery.data.noResults) { $state = 'NoResults' } }
        $canEdit = (Test-WRole $user 'admin') -or $assigned -contains $sourceId
        if ($canEdit) {
            $edit = Get-DeliveryForEdit $Context $actual $sourceId
            if ($edit) {
                $editingVersion = $edit.version
                if ($edit.data.status -eq 'Draft') { if ($delivery) { $state = 'PublishedWithDraft' } else { $state = 'Draft' }; $count = @($edit.data.documents).Count }
            }
        }
        $states += [pscustomobject]@{ sourceId = $sourceId; status = $state; documentCount = $count; publishedVersion = $publishedVersion; editingVersion = $editingVersion }
    }
    $evaluationState = 'NotEvaluated'
    if ($evaluation) { $evaluationState = 'Partial'; if ($evaluation.data.complete) { $evaluationState = 'Complete' }; if ($evaluation.needsReconfirmation) { $evaluationState = 'NeedsReconfirmation' } }
    $aiStates = @($states | Where-Object {$sourceMap.ContainsKey($_.sourceId) -and $sourceMap[$_.sourceId].type -eq 'AI'})
    $aiCompleted = @($aiStates | Where-Object {$_.publishedVersion -gt 0}).Count
    $label = if ($aiStates.Count -gt 0 -and $aiCompleted -eq $aiStates.Count) {'입력 완료'} else {'대기 중'}
    return [pscustomobject]@{ caseId = $actual.caseId; submissionStatus = $actual.status; deliveries = @($states); evaluationStatus = $evaluationState; deliveryStatusLabel=$label; aiCompleted=$aiCompleted; aiTotal=$aiStates.Count }
}

function Set-PeriodStatus {
    param([Parameter(Mandatory)]$Context, [Parameter(Mandatory)][string]$Period, [Parameter(Mandatory)][ValidateSet('Open','Closed')][string]$Status,
        [Parameter(Mandatory)][int]$ExpectedVersion, [Parameter(Mandatory)][string]$RequestId, [string]$Reason = '')
    $user = Assert-WAdmin $Context
    Assert-WPeriod $Period
    Assert-WId $RequestId '저장 요청 ID'
    $path = Join-Path (Join-Path $Context.root 'Periods') $Period
    $fingerprint = Get-WFingerprint @{operation='SetPeriodStatus';period=$Period;status=$Status;expectedVersion=$ExpectedVersion;reason=$Reason}
    return Invoke-FileLock -Path (Get-WPeriodLock $Context $Period) -Action {
        $path = Join-Path (Join-Path $Context.root 'Periods') $Period
        $user = Assert-WAdmin $Context
        $oldRequest = @(Get-Revisions -Path $path | Where-Object { $_.requestId -eq $RequestId }) | Select-Object -First 1
        if ($oldRequest) { Assert-WRequest $oldRequest $fingerprint $user.userId; return Convert-WRecord $oldRequest $path }
        $settings = Get-Period $Context $Period
        $data = $settings | Select-Object * -ExcludeProperty version,path
        Set-WValue $data 'status' $Status
        Set-WValue $data 'statusChangedAt' ([datetimeoffset]::UtcNow.ToString('o'))
        Set-WValue $data 'statusReason' $Reason
        Set-WValue $data 'requestFingerprint' $fingerprint
        $saved = Write-Revision -Path $path -ExpectedVersion $ExpectedVersion -RequestId $RequestId -ActorId $user.userId -Reason $Reason -Data $data -Kind 'Period' -TargetId $Period
        return Convert-WRecord $saved $path
    }
}

function New-User {
    param([Parameter(Mandatory)]$Context, [Parameter(Mandatory)][string]$Name,
        [string]$Bureau = '', [string]$Division = '', [string[]]$Roles = @('examiner'),
        [Parameter(Mandatory)][string]$RequestId, [string]$Reason = '', [bool]$Active = $true)
    $user = Assert-WAdmin $Context
    Assert-WId $RequestId '저장 요청 ID'
    $nameText = Convert-WOptionalText $Name '이름'
    if (-not $nameText) { throw '이름을 입력하십시오.' }
    $bureauText = Convert-WOptionalText $Bureau '심사국'
    $divisionText = Convert-WOptionalText $Division '심사과'
    $roleValues = @($Roles | Sort-Object -Unique)
    foreach ($role in $roleValues) { if ($role -notin @('admin','examiner','ai')) { throw '사용자 역할이 올바르지 않습니다.' } }
    $fingerprint = Get-WFingerprint @{operation='NewUser';name=$nameText;bureau=$bureauText;division=$divisionText;roles=@($roleValues);active=$Active;reason=$Reason}
    return Invoke-FileLock -Path (Join-Path $Context.root 'Locks\master-users-registration.lock') -Action {
        $user = Assert-WAdmin $Context
        foreach ($existing in @(Get-WMasterRecords $Context.root 'Users')) {
            foreach ($revision in @(Get-Revisions -Path $existing.path)) {
                if ($revision.requestId -ceq $RequestId) {
                    Assert-WRequest $revision $fingerprint $user.userId
                    return Convert-WRecord $revision $existing.path
                }
            }
        }
        do {
            $userId = 'U' + [guid]::NewGuid().ToString('N')
            $path = Get-WMasterPath $Context.root 'Users' $userId
        } while ([System.IO.Directory]::Exists($path))
        $data = [pscustomobject]@{userId=$userId;sid=('APP-' + [guid]::NewGuid().ToString('N'));name=$nameText;bureau=$bureauText;
            division=$divisionText;roles=@($roleValues);active=$Active;requestFingerprint=$fingerprint}
        $saved = Write-Revision -Path $path -ExpectedVersion 0 -RequestId $RequestId -ActorId $user.userId -Reason $Reason -Data $data -Kind Users -TargetId $userId
        return Convert-WRecord $saved $path
    }
}

function Save-MasterRecord {
    param([Parameter(Mandatory)]$Context, [Parameter(Mandatory)][Alias('Kind')][ValidateSet('Users','Groups','Sources')][string]$Type,
        [Parameter(Mandatory)][string]$Id, [Parameter(Mandatory)]$Data, [Parameter(Mandatory)][int]$ExpectedVersion,
        [Parameter(Mandatory)][string]$RequestId, [string]$Reason = '')
    $user = Assert-WAdmin $Context
    Assert-WId $Id
    Assert-WId $RequestId '저장 요청 ID'
    $clean = ConvertFrom-WJson (ConvertTo-Json -InputObject $Data -Depth 60)
    $clean = $clean | Select-Object * -ExcludeProperty version,path,requestFingerprint
    $idField = @{ Users = 'userId'; Groups = 'groupId'; Sources = 'sourceId' }[$Type]
    if ([string](Get-WValue $clean $idField) -ne $Id) { throw '마스터 ID와 저장 경로가 일치하지 않습니다.' }
    if ([string]::IsNullOrWhiteSpace([string](Get-WValue $clean 'name'))) { throw '이름을 입력하십시오.' }
    if ($null -ne (Get-WValue $clean 'active') -and (Get-WValue $clean 'active') -isnot [bool]) { throw 'active는 true 또는 false여야 합니다.' }
    if ($Type -eq 'Users') {
        foreach ($field in @('bureau','division')) {
            $fieldValue = $null
            if ($null -ne $clean.PSObject.Properties[$field]) { $fieldValue = $clean.$field }
            $label = @{bureau='심사국';division='심사과'}[$field]
            Set-WValue $clean $field (Convert-WOptionalText -Value $fieldValue -Label $label)
        }
        $roles = @((Get-WValue $clean 'roles' @()))
        foreach ($role in $roles) { if ($role -notin @('admin','examiner','ai')) { throw '사용자 역할이 올바르지 않습니다.' } }
        if ($Id -eq $user.userId -and (-not [bool](Get-WValue $clean 'active' $true) -or $roles -notcontains 'admin')) { throw '현재 관리자는 본인의 관리자 권한을 해제하거나 비활성화할 수 없습니다.' }
    }
    if ($Type -eq 'Sources' -and [string](Get-WValue $clean 'type') -notin @('General','AI')) { throw '조사 방식 유형은 General 또는 AI여야 합니다.' }
    # SID is managed metadata. Excluding it makes retries of missing-SID creates
    # deterministic and ensures an edit cannot change an existing legacy SID.
    $intentData = $clean
    if ($Type -eq 'Users') { $intentData = $clean | Select-Object * -ExcludeProperty sid }
    $fingerprint = Get-WFingerprint @{operation='SaveMasterRecord';type=$Type;id=$Id;data=$intentData;expectedVersion=$ExpectedVersion;reason=$Reason}
    Set-WValue $clean 'requestFingerprint' $fingerprint
    $path = Get-WMasterPath $Context.root $Type $Id
    $saveAction = {
        $user = Assert-WAdmin $Context
        $oldRequest = @(Get-Revisions -Path $path | Where-Object { $_.requestId -ceq $RequestId }) | Select-Object -First 1
        if ($oldRequest) { Assert-WRequest $oldRequest $fingerprint $user.userId; return Convert-WRecord $oldRequest $path }
        if ($Type -eq 'Users') {
            $previous = Read-Revision -Path $path
            if ($previous) { Set-WValue $clean 'sid' ([string](Get-WValue $previous.data 'sid' '')) }
            elseif ([string]::IsNullOrWhiteSpace([string](Get-WValue $clean 'sid' ''))) { Set-WValue $clean 'sid' ('APP-' + [guid]::NewGuid().ToString('N')) }
        }
        $saved = Write-Revision -Path $path -ExpectedVersion $ExpectedVersion -RequestId $RequestId -ActorId $user.userId -Reason $Reason -Data $clean -Kind $Type -TargetId $Id
        return Convert-WRecord $saved $path
    }
    if ($Type -eq 'Users') { return Invoke-FileLock -Path (Join-Path $Context.root 'Locks\master-users-registration.lock') -Action $saveAction }
    return & $saveAction
}

function Save-Period {
    param([Parameter(Mandatory)]$Context, [Parameter(Mandatory)][string]$Period, [Parameter(Mandatory)]$Data,
        [Parameter(Mandatory)][int]$ExpectedVersion, [Parameter(Mandatory)][string]$RequestId, [string]$Reason = '')
    $user = Assert-WAdmin $Context
    Assert-WPeriod $Period
    $clean = ConvertFrom-WJson (ConvertTo-Json -InputObject $Data -Depth 60)
    $clean = $clean | Select-Object * -ExcludeProperty version,path,requestFingerprint
    Set-WValue $clean 'requestFingerprint' (Get-WFingerprint @{operation='SavePeriod';period=$Period;data=$clean;expectedVersion=$ExpectedVersion;reason=$Reason})
    if ([string](Get-WValue $clean 'period') -ne $Period) { throw '운영월과 저장 경로가 일치하지 않습니다.' }
    if ([string](Get-WValue $clean 'status') -notin @('Open','Closed')) { throw '운영 상태는 Open 또는 Closed여야 합니다.' }
    $users = @{}; foreach ($item in @(Get-WMasterRecords $Context.root 'Users')) { $users[$item.userId] = $item }
    $groups = @{}; foreach ($item in @(Get-WMasterRecords $Context.root 'Groups')) { $groups[$item.groupId] = $item }
    $sources = @{}; foreach ($item in @(Get-WMasterRecords $Context.root 'Sources')) { $sources[$item.sourceId] = $item }
    $memberMap = @{}
    foreach ($member in @((Get-WValue $clean 'members' @()))) {
        if (-not $users.ContainsKey([string]$member.userId) -or -not $groups.ContainsKey([string]$member.groupId)) { throw '등록되지 않은 참여자 또는 그룹입니다.' }
        if ($memberMap.ContainsKey([string]$member.userId)) { throw '한 운영월의 심사관은 하나의 그룹에만 배정할 수 있습니다.' }
        if (-not (Test-WRole $users[$member.userId] 'examiner')) { throw '참여자는 심사관 역할이 있어야 합니다.' }
        $memberMap[$member.userId] = $member.groupId
    }
    $sourceIds = @((Get-WValue $clean 'sourceIds' @()))
    if ($sourceIds.Count -eq 0 -or @($sourceIds | Select-Object -Unique).Count -ne $sourceIds.Count) { throw '비교 조사 방식을 중복 없이 선택하십시오.' }
    foreach ($sourceId in $sourceIds) { if (-not $sources.ContainsKey([string]$sourceId)) { throw '등록되지 않은 조사 방식입니다.' } }
    $assignmentKeys = @{}
    foreach ($assignment in @((Get-WValue $clean 'assignments' @()))) {
        $id = [string]$assignment.userId
        $sourceId = [string]$assignment.sourceId
        if (-not $users.ContainsKey($id) -or -not (Test-WRole $users[$id] 'ai') -or -not $groups.ContainsKey([string]$assignment.groupId) -or $sourceIds -notcontains $sourceId) { throw 'AI 담당자 배정이 올바르지 않습니다.' }
        if ($sources[$sourceId].type -ne 'AI') { throw '일반 조사 방식에는 AI 담당자를 배정하지 않습니다.' }
        $key = [string]$assignment.groupId + '/' + $sourceId
        if ($assignmentKeys.ContainsKey($key)) { throw '그룹과 조사 방식별 담당자는 한 명으로 배정하십시오.' }
        $assignmentKeys[$key] = $true
    }
    foreach ($field in @('submissionDeadline','deliveryDeadline','evaluationDeadline')) {
        $value = [string](Get-WValue $clean $field '')
        if ($value) { $null = [datetimeoffset]::Parse($value, [System.Globalization.CultureInfo]::InvariantCulture) }
    }
    if ([int](Get-WValue $clean 'submissionLimit' 0) -lt 0) { throw '제출 상한은 0 이상이어야 합니다. 0은 제한 없음입니다.' }
    $path = Join-Path (Join-Path $Context.root 'Periods') $Period
    return Invoke-FileLock -Path (Get-WPeriodLock $Context $Period) -Action {
        $path = Join-Path (Join-Path $Context.root 'Periods') $Period
        $user = Assert-WAdmin $Context
        $previous = Read-Revision -Path $path
        if ($previous -and @(Get-Cases -Context $Context -Period $Period -Scope All).Count -gt 0) {
            foreach ($oldSourceId in @((Get-WValue $previous.data 'sourceIds' @()))) {
                if ($sourceIds -notcontains $oldSourceId) { throw '제출 건이 있는 운영월의 기존 비교 조사 방식을 제거할 수 없습니다.' }
            }
        }
        $saved = Write-Revision -Path $path -ExpectedVersion $ExpectedVersion -RequestId $RequestId -ActorId $user.userId -Reason $Reason -Data $clean -Kind 'Period' -TargetId $Period
        return Convert-WRecord $saved $path
    }
}

function Initialize-Demo {
    param([Parameter(Mandatory)][string]$Root)
    $rootPath = [System.IO.Path]::GetFullPath($Root)
    $marker = Join-Path $rootPath '.demo-mode.json'
    $null = [System.IO.Directory]::CreateDirectory($rootPath)
    $existingUsers = @(Get-WMasterRecords $rootPath 'Users')
    if ($existingUsers.Count -gt 0 -and -not [System.IO.File]::Exists($marker)) { throw '기존 운영 데이터에 데모 모드를 추가할 수 없습니다.' }
    if (-not [System.IO.File]::Exists($marker)) {
        [System.IO.File]::WriteAllText($marker, '{"demo":true}', (New-Object System.Text.UTF8Encoding($false)))
    }
    $records = @(
        @{ type='Users'; id='U001'; data=@{userId='U001';sid=('APP-'+[guid]::NewGuid().ToString('N'));name='데모 관리자';bureau='';division='';roles=@('admin','examiner','ai');active=$true} },
        @{ type='Users'; id='U002'; data=@{userId='U002';sid=('APP-'+[guid]::NewGuid().ToString('N'));name='데모 심사관';bureau='';division='';roles=@('examiner');active=$true} },
        @{ type='Users'; id='U003'; data=@{userId='U003';sid=('APP-'+[guid]::NewGuid().ToString('N'));name='데모 클로드 담당자';bureau='';division='';roles=@('examiner','ai');active=$true} },
        @{ type='Groups'; id='G01'; data=@{groupId='G01';name='시범 그룹';active=$true} },
        @{ type='Sources'; id='S01'; data=@{sourceId='S01';name='선행기술조사';type='General';active=$true} },
        @{ type='Sources'; id='S02'; data=@{sourceId='S02';name='ChatGPT';type='AI';active=$true} },
        @{ type='Sources'; id='S03'; data=@{sourceId='S03';name='클로드';type='AI';active=$true} }
    )
    foreach ($record in $records) {
        $path = Get-WMasterPath $rootPath $record.type $record.id
        if ($null -eq (Read-Revision -Path $path)) {
            $null = Write-Revision -Path $path -ExpectedVersion 0 -RequestId ([guid]::NewGuid().ToString('N')) -ActorId 'setup' -Reason '데모 초기 설정' -Data ([pscustomobject]$record.data) -Kind $record.type -TargetId $record.id
        }
    }
    $period = Get-Date -Format 'yyyy-MM'
    $path = Join-Path (Join-Path $rootPath 'Periods') $period
    if ($null -eq (Read-Revision -Path $path)) {
        $settings = [pscustomobject]@{period=$period;status='Open';members=@(@{userId='U001';groupId='G01'},@{userId='U002';groupId='G01'},@{userId='U003';groupId='G01'});
            sourceIds=@('S01','S02','S03');assignments=@(@{userId='U001';groupId='G01';sourceId='S02'},@{userId='U003';groupId='G01';sourceId='S03'});
            submissionDeadline='';deliveryDeadline='';evaluationDeadline='';submissionLimit=0}
        $null = Write-Revision -Path $path -ExpectedVersion 0 -RequestId ([guid]::NewGuid().ToString('N')) -ActorId 'setup' -Reason '데모 운영월 개시' -Data $settings -Kind 'Period' -TargetId $period
    }
    return Get-Context -Root $rootPath
}

. (Join-Path $PSScriptRoot 'Statistics.ps1')
Export-ModuleMember -Function Initialize-Demo,Get-Context,Get-LoginUsers,New-User,Get-Periods,Get-Users,Get-Groups,Get-Sources,Get-Period,Get-Cases,Get-TargetSelection,Save-TargetSelection,New-Submission,Cancel-Submission,Get-Deliveries,Get-DeliveryForEdit,Get-DeliveryHistory,Save-Delivery,Import-DeliveryAttachment,Resolve-DeliveryAttachment,Get-Evaluation,Save-Evaluation,Get-CaseStatus,Set-PeriodStatus,Save-MasterRecord,Save-Period,Get-AdminStatistics,Select-AdminStatistics
