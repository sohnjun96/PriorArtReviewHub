param([switch]$KeepArtifacts)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0
$modulePath = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../App/Modules/Storage.psm1'))
Import-Module $modulePath -Force
$testRoot = Join-Path ([IO.Path]::GetTempPath()) ('AiSearch-StorageTests-' + [Guid]::NewGuid().ToString('N'))
[void][IO.Directory]::CreateDirectory($testRoot)
$script:passed = 0
$script:children = New-Object 'System.Collections.Generic.List[System.Diagnostics.Process]'

function Assert-True {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw "Assertion failed: $Message" }
}

function Assert-Throws {
    param([scriptblock]$Action, [string]$Pattern)
    $caught = $null
    try { & $Action | Out-Null }
    catch { $caught = $_ }
    if ($null -eq $caught) { throw "Expected an exception matching '$Pattern'." }
    if ($caught.Exception.Message -notmatch $Pattern) {
        throw "Expected '$Pattern', received '$($caught.Exception.Message)'."
    }
}

function Test-Case {
    param([string]$Name, [scriptblock]$Action)
    & $Action
    $script:passed++
    Write-Host "PASS $Name"
}

function Start-TestChild {
    param([string]$Code)
    $encoded = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($Code))
    $start = New-Object Diagnostics.ProcessStartInfo
    $start.FileName = [Diagnostics.Process]::GetCurrentProcess().MainModule.FileName
    $start.Arguments = '-NoProfile -NonInteractive -ExecutionPolicy Bypass -EncodedCommand ' + $encoded
    $start.UseShellExecute = $false
    $start.CreateNoWindow = $true
    $start.RedirectStandardOutput = $true
    $start.RedirectStandardError = $true
    $process = New-Object Diagnostics.Process
    $process.StartInfo = $start
    [void]$process.Start()
    $script:children.Add($process)
    return $process
}

function Wait-TestFile {
    param([string]$Path)
    $watch = [Diagnostics.Stopwatch]::StartNew()
    while (-not [IO.File]::Exists($Path)) {
        if ($watch.ElapsedMilliseconds -gt 10000) { throw "Timed out waiting for child marker '$Path'." }
        Start-Sleep -Milliseconds 20
    }
}

function Complete-TestChild {
    param([Diagnostics.Process]$Process, [int[]]$AllowedExitCodes = @(0))
    if (-not $Process.WaitForExit(10000)) { throw 'Child process did not finish.' }
    $stdout = $Process.StandardOutput.ReadToEnd()
    $stderr = $Process.StandardError.ReadToEnd()
    if ($AllowedExitCodes -notcontains $Process.ExitCode) {
        throw "Child failed with exit code $($Process.ExitCode). $stdout $stderr"
    }
    return $Process.ExitCode
}

function Quote-PSString {
    param([string]$Text)
    return "'" + $Text.Replace("'", "''") + "'"
}

$moduleLiteral = Quote-PSString $modulePath
try {
    Test-Case 'Missing record and immutable UTF-8 Korean revision round trip' {
        $path = Join-Path $testRoot 'roundtrip'
        Assert-True ($null -eq (Read-Revision $path)) 'new entity is null'
        Assert-True (@(Get-Revisions $path).Count -eq 0) 'new history is empty'
        $record = Write-Revision $path -ExpectedVersion 0 -RequestId 'request-1' -ActorId 'U023' -Kind 'Delivery' -TargetId 'C001' -Reason '최초 저장' -Data ([ordered]@{ title = '한글 문헌'; score = $null; docs = @(); number = '1020260012345' })
        Assert-True ($record.version -eq 1) 'first version is one'
        Assert-True ($record.data.title -ceq '한글 문헌') 'Korean text survives'
        Assert-True ($null -eq $record.data.score) 'null is distinct from zero'
        Assert-True ($record.data.docs -is [array] -and $record.data.docs.Count -eq 0) 'empty arrays survive'
        Assert-True ($record.data.number -is [string]) 'identifiers stay strings'
        $firstFile = Join-Path $path 'revisions/000001.json'
        $original = [IO.File]::ReadAllBytes($firstFile)
        Assert-True ($original[0] -eq 123) 'JSON is UTF-8 without BOM'
        $second = Write-Revision $path -ExpectedVersion 1 -RequestId 'request-2' -ActorId 'U023' -Data @{ title = '수정'; docs = @(@{id='D1'}, @{id='D2'}) }
        Assert-True ($second.version -eq 2) 'next version increments'
        Assert-True ([Convert]::ToBase64String($original) -ceq [Convert]::ToBase64String([IO.File]::ReadAllBytes($firstFile))) 'old revision unchanged'
        Assert-True ((Read-Revision $path).version -eq 2 -and @(Get-Revisions $path).Count -eq 2) 'latest read and history agree'
        foreach ($revision in @(Get-Revisions $path)) {
            Assert-True ($revision -is [pscustomobject] -and $null -ne $revision.PSObject.Properties['data']) 'history emits scalar revision envelopes'
        }
        Assert-True ([IO.File]::Exists((Join-Path $path 'write.lock'))) 'lock file remains'
    }

    Test-Case 'Idempotent request matching is semantic and stale saves are rejected' {
        $path = Join-Path $testRoot 'idempotent'
        $first = Write-Revision $path -ExpectedVersion 0 -RequestId 'stable' -ActorId 'U1' -Kind 'Test' -Data ([ordered]@{ b = @(@{x=1;y=2}); a = 'value' })
        $retry = Write-Revision $path -ExpectedVersion 0 -RequestId 'stable' -ActorId 'U1' -Kind 'Test' -Data ([ordered]@{ a = 'value'; b = @([ordered]@{y=2;x=1}) })
        Assert-True ($retry.version -eq 1 -and $retry.writtenAt -ceq $first.writtenAt) 'retry returns original revision'
        Assert-True (@(Get-Revisions $path).Count -eq 1) 'retry adds no revision'
        Assert-Throws { Write-Revision $path -ExpectedVersion 1 -RequestId 'stable' -ActorId 'U1' -Kind 'Test' -Data @{a='other'} } 'different content'
        Assert-Throws { Write-Revision $path -ExpectedVersion 1 -RequestId 'stable' -ActorId 'U2' -Kind 'Test' -Data ([ordered]@{ a = 'value'; b = @(@{x=1;y=2}) }) } 'different content'
        Assert-Throws { Write-Revision $path -ExpectedVersion 0 -RequestId 'new-request' -ActorId 'U1' -Data @{} } 'Version conflict'
        Assert-True (@(Get-Revisions $path).Count -eq 1) 'rejections preserve history'
        $numbers = Join-Path $testRoot 'numeric-semantic'
        $null = Write-Revision $numbers -ExpectedVersion 0 -RequestId 'number' -ActorId 'U1' -Data @{n=1}
        $numericRetry = Write-Revision $numbers -ExpectedVersion 0 -RequestId 'number' -ActorId 'U1' -Data @{n=1.0}
        Assert-True ($numericRetry.version -eq 1) 'equivalent JSON numbers match'
        Assert-Throws { Write-Revision $numbers -ExpectedVersion 0 -RequestId 'number' -ActorId 'U1' -Data @{n='1'} } 'different content'
    }

    Test-Case 'Validation prevents commit and temporary files are ignored' {
        $path = Join-Path $testRoot 'validation'
        Assert-Throws { Write-Revision $path -ExpectedVersion 0 -RequestId 'v1' -ActorId 'U1' -Data @{n=8} -Validate { param($value) if ($value.n -gt 7) { throw 'Seven documents maximum' } } } 'Seven documents'
        Assert-Throws { Write-Revision $path -ExpectedVersion 0 -RequestId 'v2' -ActorId 'U1' -Data @{} -Validate { $false } } 'validation rejected'
        Assert-True ($null -eq (Read-Revision $path)) 'validation failure adds nothing'
        $null = Write-Revision $path -ExpectedVersion 0 -RequestId 'v3' -ActorId 'U1' -Data @{n=7} -Validate { param($value) $value.n -le 7 }
        [IO.File]::WriteAllText((Join-Path $path 'revisions/.pending-crashed.tmp'), '{broken')
        Assert-True ((Read-Revision $path).version -eq 1) 'uncommitted temp ignored'
        Assert-Throws { Invoke-FileLock (Join-Path $path 'action.lock') { throw 'Action failure' } } 'Action failure'
        Assert-True ((Invoke-FileLock (Join-Path $path 'action.lock') { 'released' }) -eq 'released') 'failed action releases lock'
    }

    Test-Case 'Corrupt committed record, mismatched filename and sequence gaps fail loudly' {
        $corrupt = Join-Path $testRoot 'corrupt'
        $null = Write-Revision $corrupt -ExpectedVersion 0 -RequestId 'c1' -ActorId 'U1' -Data @{}
        $null = Write-Revision $corrupt -ExpectedVersion 1 -RequestId 'c2' -ActorId 'U1' -Data @{}
        [IO.File]::WriteAllText((Join-Path $corrupt 'revisions/000002.json'), '{broken')
        Assert-Throws { Read-Revision $corrupt } 'Cannot read committed revision'
        Assert-Throws { Write-Revision $corrupt -ExpectedVersion 2 -RequestId 'c3' -ActorId 'U1' -Data @{} } 'Cannot read committed revision'
        $gap = Join-Path $testRoot 'gap'
        $null = Write-Revision $gap -ExpectedVersion 0 -RequestId 'g1' -ActorId 'U1' -Data @{}
        [IO.File]::Move((Join-Path $gap 'revisions/000001.json'), (Join-Path $gap 'revisions/000002.json'))
        Assert-Throws { Read-Revision $gap } 'sequence is incomplete'
        $mismatch = Join-Path $testRoot 'mismatch'
        $null = Write-Revision $mismatch -ExpectedVersion 0 -RequestId 'm1' -ActorId 'U1' -Data @{}
        $file = Join-Path $mismatch 'revisions/000001.json'
        $broken = [regex]::Replace([IO.File]::ReadAllText($file), '("version"\s*:\s*)1', '${1}9')
        [IO.File]::WriteAllText($file, $broken)
        Assert-Throws { Read-Revision $mismatch } 'does not match filename'
    }

    Test-Case 'Root-contained path builder rejects traversal and Windows path aliases' {
        $safe = Get-SafePath -Root $testRoot -Segments @('Data', '2026-10', 'C001')
        Assert-True ($safe -eq (Join-Path $testRoot 'Data/2026-10/C001')) 'safe segments combine'
        foreach ($unsafe in @('..', '.', '../outside', 'bad/name', 'bad\name', 'C:escape', 'trailing.', 'trailing ', 'CON', 'NUL.json')) {
            Assert-Throws { Get-SafePath -Root $testRoot -Segments @($unsafe) } 'Unsafe path'
        }
    }

    Test-Case 'Separate-process exclusive lock, timeout and normal release' {
        $lock = Join-Path $testRoot 'contention.lock'
        $marker = Join-Path $testRoot 'contention-ready'
        $code = '$ErrorActionPreference = ''Stop''; Import-Module ' + $moduleLiteral + '; Invoke-FileLock -Path ' + (Quote-PSString $lock) + ' -Action { [IO.File]::WriteAllText(' + (Quote-PSString $marker) + ', ''ready''); Start-Sleep -Milliseconds 800 }'
        $child = Start-TestChild $code
        Wait-TestFile $marker
        Assert-Throws { Invoke-FileLock $lock { throw 'Should never enter' } -TimeoutMilliseconds 100 } 'Timed out acquiring'
        $null = Complete-TestChild $child
        Assert-True ((Invoke-FileLock $lock { 42 } -TimeoutMilliseconds 1000) -eq 42) 'lock available after other process exits'
        Assert-True ([IO.File]::Exists($lock)) 'shared lock path is never deleted'
    }

    Test-Case 'Two separate writer processes cannot both commit version one' {
        $path = Join-Path $testRoot 'race'
        $go = Join-Path $testRoot 'race-go'
        $workers = @()
        foreach ($id in @('a','b')) {
            $ready = Join-Path $testRoot ('race-ready-' + $id)
            $code = '$ErrorActionPreference = ''Stop''; Import-Module ' + $moduleLiteral + '; [IO.File]::WriteAllText(' + (Quote-PSString $ready) + ', ''ready''); while (-not [IO.File]::Exists(' + (Quote-PSString $go) + ')) { Start-Sleep -Milliseconds 20 }; try { $null = Write-Revision -Path ' + (Quote-PSString $path) + ' -ExpectedVersion 0 -RequestId ' + (Quote-PSString $id) + ' -ActorId ''U1'' -Data @{worker=' + (Quote-PSString $id) + '}; exit 0 } catch { if ($_.Exception.Message -match ''Version conflict'') { exit 11 }; [Console]::Error.WriteLine($_); exit 99 }'
            $workers += Start-TestChild $code
            Wait-TestFile $ready
        }
        [IO.File]::WriteAllText($go, 'go')
        $codes = @($workers | ForEach-Object { Complete-TestChild $_ -AllowedExitCodes @(0,11) })
        Assert-True (@($codes | Where-Object { $_ -eq 0 }).Count -eq 1) 'one writer commits'
        Assert-True (@($codes | Where-Object { $_ -eq 11 }).Count -eq 1) 'other writer receives stale version conflict'
        Assert-True (@(Get-Revisions $path).Count -eq 1) 'only one committed file exists'
    }

    Test-Case 'Terminated process releases the OS file lock' {
        $lock = Join-Path $testRoot 'terminated.lock'
        $marker = Join-Path $testRoot 'terminated-ready'
        $code = '$ErrorActionPreference = ''Stop''; Import-Module ' + $moduleLiteral + '; Invoke-FileLock -Path ' + (Quote-PSString $lock) + ' -Action { [IO.File]::WriteAllText(' + (Quote-PSString $marker) + ', ''ready''); Start-Sleep -Seconds 30 }'
        $child = Start-TestChild $code
        Wait-TestFile $marker
        $child.Kill()
        [void]$child.WaitForExit(5000)
        Assert-True ((Invoke-FileLock $lock { 'recovered' } -TimeoutMilliseconds 2000) -eq 'recovered') 'lock recovers after abrupt termination'
    }
    Write-Host ("Storage tests: {0} passed (PowerShell {1})." -f $script:passed, $PSVersionTable.PSVersion)
}
finally {
    foreach ($child in $script:children) {
        try { if (-not $child.HasExited) { $child.Kill(); [void]$child.WaitForExit(5000) } } catch { }
        $child.Dispose()
    }
    if ($KeepArtifacts) { Write-Host "Artifacts: $testRoot" }
    else {
        # Restrict recursive cleanup to the exact unique directory created above.
        $resolved = [IO.Path]::GetFullPath($testRoot)
        $tempPrefix = [IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\') + '\'
        if (-not $resolved.StartsWith($tempPrefix, [StringComparison]::OrdinalIgnoreCase) -or
            [IO.Path]::GetFileName($resolved) -notmatch '^AiSearch-StorageTests-[a-f0-9]{32}$') {
            throw 'Refusing cleanup outside the generated test directory.'
        }
        Remove-Item -LiteralPath $resolved -Recurse -Force
    }
}
