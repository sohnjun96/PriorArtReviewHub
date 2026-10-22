param([switch]$KeepArtifacts)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0
$testRoot = Join-Path ([IO.Path]::GetTempPath()) ('AiReview-ExecutionPolicyTests-' + [guid]::NewGuid().ToString('N'))
$windowsPowerShell = (Get-Command powershell.exe -ErrorAction Stop).Source
$parentPreference = [Environment]::GetEnvironmentVariable('PSExecutionPolicyPreference', 'Process')
$parentPolicies = @(Get-ExecutionPolicy -List | ForEach-Object { '{0}={1}' -f $_.Scope, $_.ExecutionPolicy }) -join ';'
$script:passed = 0

function Assert-Policy {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw "Assertion failed: $Message" }
}

function Test-PolicyCase {
    param([string]$Name, [scriptblock]$Action)
    & $Action
    $script:passed++
    Write-Host "PASS $Name"
}

function Invoke-PolicyChild {
    param([string]$Arguments)
    $start = New-Object Diagnostics.ProcessStartInfo
    $start.FileName = $windowsPowerShell
    $start.Arguments = '-NoLogo -NoProfile -NonInteractive ' + $Arguments
    $start.UseShellExecute = $false
    $start.CreateNoWindow = $true
    $start.RedirectStandardOutput = $true
    $start.RedirectStandardError = $true
    # Only this child receives the restrictive policy. No parent or registry settings change.
    $start.EnvironmentVariables['PSExecutionPolicyPreference'] = 'Restricted'
    # A PowerShell 7 host can otherwise direct Windows PowerShell to incompatible core modules.
    $start.EnvironmentVariables.Remove('PSModulePath')
    $process = New-Object Diagnostics.Process
    $process.StartInfo = $start
    try {
        [void]$process.Start()
        if (-not $process.WaitForExit(15000)) {
            $process.Kill()
            throw 'Windows PowerShell policy child timed out.'
        }
        [pscustomobject]@{
            exitCode = $process.ExitCode
            output = $process.StandardOutput.ReadToEnd()
            error = $process.StandardError.ReadToEnd()
        }
    }
    finally { $process.Dispose() }
}

function Assert-PolicyFixtureSuccess {
    param($Result)
    Assert-Policy ($Result.exitCode -eq 0) ('fixture runs: ' + $Result.error)
    Assert-Policy ($Result.output.Trim() -ceq 'RemoteSigned|fixture-module-ok') 'script and imported module run under RemoteSigned'
}

function Test-FixtureZone {
    param([string]$Path)
    return $null -ne (Get-Item -LiteralPath $Path -Stream Zone.Identifier -ErrorAction SilentlyContinue)
}

try {
    [void][IO.Directory]::CreateDirectory($testRoot)
    $fixtureScript = Join-Path $testRoot 'Fixture.ps1'
    $fixtureModule = Join-Path $testRoot 'Fixture.psm1'
    $utf8Bom = New-Object Text.UTF8Encoding($true)
    [IO.File]::WriteAllText($fixtureScript, @'
$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot 'Fixture.psm1') -Force -ErrorAction Stop
'{0}|{1}' -f (Get-ExecutionPolicy), (Get-PolicyFixtureValue)
'@, $utf8Bom)
    [IO.File]::WriteAllText($fixtureModule, @'
function Get-PolicyFixtureValue { 'fixture-module-ok' }
Export-ModuleMember -Function Get-PolicyFixtureValue
'@, $utf8Bom)

    $probeCode = '[pscustomobject]@{ machine=[string](Get-ExecutionPolicy -Scope MachinePolicy); user=[string](Get-ExecutionPolicy -Scope UserPolicy); effective=[string](Get-ExecutionPolicy) } | ConvertTo-Json -Compress'
    $encodedProbe = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($probeCode))
    $probeResult = Invoke-PolicyChild -Arguments ('-EncodedCommand ' + $encodedProbe)
    Assert-Policy ($probeResult.exitCode -eq 0) ('policy probe runs: ' + $probeResult.error)
    $probe = $probeResult.output | ConvertFrom-Json

    if ($probe.machine -ne 'Undefined' -or $probe.user -ne 'Undefined') {
        Write-Host ('SKIP child policy matrix: Group Policy has priority (MachinePolicy={0}, UserPolicy={1}).' -f $probe.machine, $probe.user)
    }
    else {
        Test-PolicyCase 'Windows PowerShell child inherits Restricted without changing parent policy' {
            Assert-Policy ($probe.effective -ceq 'Restricted') 'child process fixture policy is Restricted'
        }
        Test-PolicyCase 'Restricted blocks the local script before module import' {
            $result = Invoke-PolicyChild -Arguments ('-STA -File "' + $fixtureScript + '"')
            Assert-Policy ($result.exitCode -ne 0) 'script is rejected without launcher policy option'
            Assert-Policy ($result.output -notmatch 'fixture-module-ok') 'rejected script does not run'
        }
        Test-PolicyCase 'Launcher RemoteSigned permits local script and module despite inherited Restricted' {
            $result = Invoke-PolicyChild -Arguments ('-STA -ExecutionPolicy RemoteSigned -File "' + $fixtureScript + '"')
            Assert-PolicyFixtureSuccess -Result $result
        }

        $supportsZoneStream = $false
        try {
            Set-Content -LiteralPath $fixtureScript -Stream Zone.Identifier -Value "[ZoneTransfer]`r`nZoneId=3" -Encoding ASCII
            Set-Content -LiteralPath $fixtureModule -Stream Zone.Identifier -Value "[ZoneTransfer]`r`nZoneId=3" -Encoding ASCII
            $supportsZoneStream = (Test-FixtureZone -Path $fixtureScript) -and (Test-FixtureZone -Path $fixtureModule)
        }
        catch { $supportsZoneStream = $false }
        if ($supportsZoneStream) {
            Test-PolicyCase 'RemoteSigned keeps downloaded unsigned script blocked' {
                $result = Invoke-PolicyChild -Arguments ('-STA -ExecutionPolicy RemoteSigned -File "' + $fixtureScript + '"')
                Assert-Policy ($result.exitCode -ne 0) 'internet-zone script is still rejected'
                Assert-Policy ($result.output -notmatch 'fixture-module-ok') 'blocked internet-zone script does not run'
            }
            Test-PolicyCase 'Unblocking entry script alone leaves downloaded module blocked' {
                Unblock-File -LiteralPath $fixtureScript
                Assert-Policy (-not (Test-FixtureZone -Path $fixtureScript)) 'only entry fixture stream is removed'
                Assert-Policy (Test-FixtureZone -Path $fixtureModule) 'module fixture keeps internet-zone stream'
                $result = Invoke-PolicyChild -Arguments ('-STA -ExecutionPolicy RemoteSigned -File "' + $fixtureScript + '"')
                Assert-Policy ($result.exitCode -ne 0) 'downloaded module is still rejected'
                Assert-Policy ($result.output -notmatch 'fixture-module-ok') 'blocked module does not run'
            }
            Test-PolicyCase 'Targeted fixture unblock restores script and module execution' {
                Unblock-File -LiteralPath $fixtureModule
                Assert-Policy (-not (Test-FixtureZone -Path $fixtureModule)) 'module fixture stream is removed'
                $result = Invoke-PolicyChild -Arguments ('-STA -ExecutionPolicy RemoteSigned -File "' + $fixtureScript + '"')
                Assert-PolicyFixtureSuccess -Result $result
            }
        }
        elseif (-not $supportsZoneStream) {
            Write-Host 'SKIP downloaded-file matrix: Zone.Identifier streams are unavailable.'
        }
    }

    Test-PolicyCase 'Parent execution policy and process environment remain unchanged' {
        $afterPreference = [Environment]::GetEnvironmentVariable('PSExecutionPolicyPreference', 'Process')
        $afterPolicies = @(Get-ExecutionPolicy -List | ForEach-Object { '{0}={1}' -f $_.Scope, $_.ExecutionPolicy }) -join ';'
        Assert-Policy ([object]::Equals($parentPreference, $afterPreference)) 'parent environment is unchanged'
        Assert-Policy ($parentPolicies -ceq $afterPolicies) 'all parent policy scopes are unchanged'
    }
    Write-Host ('Execution policy tests: {0} passed (host PowerShell {1}; Windows PowerShell child).' -f $script:passed, $PSVersionTable.PSVersion)
}
finally {
    if ($KeepArtifacts) { Write-Host "Artifacts: $testRoot" }
    elseif ([IO.Directory]::Exists($testRoot)) {
        $resolved = [IO.Path]::GetFullPath($testRoot)
        $tempPrefix = [IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\') + '\'
        if (-not $resolved.StartsWith($tempPrefix, [StringComparison]::OrdinalIgnoreCase) -or
            [IO.Path]::GetFileName($resolved) -notmatch '^AiReview-ExecutionPolicyTests-[a-f0-9]{32}$') { throw 'Unsafe policy test cleanup boundary.' }
        Remove-Item -LiteralPath $resolved -Recurse -Force
    }
}
