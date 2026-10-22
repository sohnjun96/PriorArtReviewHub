[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$Root,
    [string]$AdminSid,
    [ValidatePattern('^[A-Za-z0-9][A-Za-z0-9_-]{0,100}$')][string]$AdminId = 'U001',
    [ValidateNotNullOrEmpty()][string]$AdminName = '총괄관리자',
    [ValidatePattern('^\d{4}-(0[1-9]|1[0-2])$')][string]$Period = (Get-Date -Format 'yyyy-MM')
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0
Import-Module (Join-Path (Split-Path $PSScriptRoot -Parent) 'App/Modules/Storage.psm1') -Force
if ([string]::IsNullOrWhiteSpace($AdminSid)) {
    $AdminSid = [Security.Principal.WindowsIdentity]::GetCurrent().User.Value
}
# Validate before creating anything. A supplied SID need not be the setup operator.
$sidObject = New-Object Security.Principal.SecurityIdentifier($AdminSid)
$AdminSid = $sidObject.Value
if ([string]::IsNullOrWhiteSpace($AdminName)) { throw '관리자 이름을 입력하세요.' }
$rootPath = [IO.Path]::GetFullPath($Root)
if ([IO.File]::Exists($rootPath)) { throw '저장소 경로가 파일입니다.' }
$initializationLock = Get-SafePath -Root $rootPath -Segments @('.initialize.lock')

Invoke-FileLock -Path $initializationLock -TimeoutMilliseconds 10000 -Action {
    # Only our persistent lock may exist. Never merge a bootstrap with demo data,
    # an earlier production store, deployment files or an incomplete prior setup.
    $existing = @(Get-ChildItem -LiteralPath $rootPath -Force | Where-Object { $_.FullName -cne $initializationLock })
    if ($existing.Count -gt 0) {
        throw '비어 있는 새 폴더에만 초기화할 수 있습니다. 기존 파일을 덮어쓰지 않았습니다.'
    }
    $records = @(
        @{ kind='Users'; id=$AdminId; data=[ordered]@{userId=$AdminId;sid=$AdminSid;name=$AdminName;roles=@('admin','examiner','ai');active=$true} },
        @{ kind='Groups'; id='G01'; data=[ordered]@{groupId='G01';name='운영 그룹';active=$true} },
        @{ kind='Sources'; id='S01'; data=[ordered]@{sourceId='S01';name='선행기술조사';type='General';active=$true} },
        @{ kind='Sources'; id='S02'; data=[ordered]@{sourceId='S02';name='ChatGPT';type='AI';active=$true} },
        @{ kind='Sources'; id='S03'; data=[ordered]@{sourceId='S03';name='클로드';type='AI';active=$true} }
    )
    foreach ($record in $records) {
        $path = Get-SafePath -Root $rootPath -Segments @('Master', $record.kind, $record.id)
        $null = Write-Revision -Path $path -ExpectedVersion 0 -RequestId ([guid]::NewGuid().ToString('N')) -ActorId $AdminId -Reason '운영 저장소 최초 설정' -Data $record.data -Kind $record.kind -TargetId $record.id
    }
    $settings = [ordered]@{
        period=$Period; status='Open'; members=@(@{userId=$AdminId;groupId='G01'});
        sourceIds=@('S01','S02','S03');
        assignments=@(@{userId=$AdminId;groupId='G01';sourceId='S02'}, @{userId=$AdminId;groupId='G01';sourceId='S03'});
        submissionDeadline=''; deliveryDeadline=''; evaluationDeadline=''; submissionLimit=0
    }
    $periodPath = Get-SafePath -Root $rootPath -Segments @('Periods', $Period)
    $null = Write-Revision -Path $periodPath -ExpectedVersion 0 -RequestId ([guid]::NewGuid().ToString('N')) -ActorId $AdminId -Reason '최초 운영월 개시' -Data $settings -Kind 'Period' -TargetId $Period
    [pscustomobject]@{ root=$rootPath; adminId=$AdminId; adminSid=$AdminSid; period=$Period; recordsCreated=6 }
}
