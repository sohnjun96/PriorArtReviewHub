param([switch]$KeepArtifacts)
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0
$projectPath = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$testRoot = Join-Path ([IO.Path]::GetTempPath()) ('AiReview-DeploymentTests-' + [guid]::NewGuid().ToString('N'))
$storePath = Join-Path $testRoot 'Store'
$deploymentPath = Join-Path $testRoot 'Program'
$script:passed = 0

function Assert-Deployment {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw "Assertion failed: $Message" }
}
function Assert-DeploymentThrows {
    param([scriptblock]$Action, [string]$Pattern)
    $caught = $null
    try { & $Action | Out-Null } catch { $caught = $_ }
    if ($null -eq $caught -or $caught.Exception.Message -notmatch $Pattern) { throw "Expected rejection matching '$Pattern'." }
}
function Test-DeploymentCase {
    param([string]$Name, [scriptblock]$Action)
    & $Action
    $script:passed++
    Write-Host "PASS $Name"
}

try {
    Test-DeploymentCase 'Production bootstrap creates automatic IDs for selection login' {
        $summary = & (Join-Path $projectPath 'Tools/Initialize-Store.ps1') -Root $storePath -AdminName '운영 초기화 검증'
        $script:initialAdminId = $summary.adminId
        Import-Module (Join-Path $projectPath 'App/Modules/Workflow.psm1') -Force -DisableNameChecking
        $context = Get-Context -Root $storePath -UserId $summary.adminId
        $loginUsers = @(Get-LoginUsers -Root $storePath)
        Assert-Deployment (-not $context.demoMode) 'production context'
        Assert-Deployment (-not [IO.File]::Exists((Join-Path $storePath '.demo-mode.json'))) 'no demo marker'
        Assert-Deployment ($summary.adminId -match '^U[a-f0-9]{32}$' -and $context.user.userId -eq $summary.adminId) 'automatic initial user ID'
        Assert-Deployment ($context.user.sid -match '^APP-[a-f0-9]{32}$' -and $summary.adminSid -eq $context.user.sid) 'automatic internal identity'
        Assert-Deployment ($loginUsers.Count -eq 1 -and $loginUsers[0].userId -eq $summary.adminId) 'initial admin appears in login choices'
        Assert-Deployment (@(Get-Users -Context $context).Count -eq 1) 'one initial user'
        Assert-Deployment (@(Get-Groups -Context $context).Count -eq 1 -and @(Get-Sources -Context $context).Count -eq 3) 'one group and three sources'
        Assert-Deployment (@($context.user.roles).Count -eq 3 -and @($context.user.roles) -contains 'admin' -and @($context.user.roles) -contains 'examiner' -and @($context.user.roles) -contains 'ai') 'all initial roles'
        $period = Get-Period -Context $context -Period $summary.period
        Assert-Deployment (@($period.members).Count -eq 1 -and @($period.assignments).Count -eq 2) 'single member and two assigned AI methods'
        Assert-Deployment ($period.members[0].userId -eq $summary.adminId -and @($period.assignments | Where-Object { $_.userId -ne $summary.adminId }).Count -eq 0) 'generated ID used in monthly assignments'
        Assert-Deployment ($summary.recordsCreated -eq 6) 'six initial records'
    }
    Test-DeploymentCase 'Production users select their identity on the same PC' {
        $adminContext = Get-Context -Root $storePath -UserId $script:initialAdminId
        $examiner = New-User -Context $adminContext -Name '동시 사용자' -Bureau '특허심사국' -Division '특허심사과' -Roles @('examiner') -RequestId ([guid]::NewGuid().ToString('N'))
        $examinerContext = Get-Context -Root $storePath -UserId $examiner.userId
        $adminAgain = Get-Context -Root $storePath -UserId $script:initialAdminId
        $loginUsers = @(Get-LoginUsers -Root $storePath)
        $examinerChoice = @($loginUsers | Where-Object { $_.userId -eq $examiner.userId })
        Assert-Deployment ($examiner.userId -match '^U[a-f0-9]{32}$' -and $examiner.sid -match '^APP-[a-f0-9]{32}$') 'new user identifiers are automatic'
        Assert-Deployment ($examinerContext.user.userId -eq $examiner.userId -and $adminAgain.user.userId -eq $script:initialAdminId) 'same Windows process selects either production user'
        Assert-Deployment ($examinerContext.user.sid -ne $adminAgain.user.sid) 'selected users have distinct internal identities'
        Assert-Deployment ($loginUsers.Count -eq 2 -and $examinerChoice.Count -eq 1 -and $examinerChoice[0].displayName -ceq '특허심사과 - 동시 사용자') 'login list uses division and name'
        Assert-Deployment (@(Get-Users -Context $examinerContext).Count -eq 1) 'examiner keeps their role-specific user view'
    }
    Test-DeploymentCase 'Repeated bootstrap and nonempty targets reject without overwriting' {
        $firstRecord = Join-Path $storePath ('Master/Users/{0}/revisions/000001.json' -f $script:initialAdminId)
        $before = [Convert]::ToBase64String([IO.File]::ReadAllBytes($firstRecord))
        Assert-DeploymentThrows { & (Join-Path $projectPath 'Tools/Initialize-Store.ps1') -Root $storePath } '비어 있는 새 폴더'
        Assert-Deployment ([Convert]::ToBase64String([IO.File]::ReadAllBytes($firstRecord)) -ceq $before) 'initial revision preserved'
        $nonempty = Join-Path $testRoot 'Nonempty'
        [void][IO.Directory]::CreateDirectory($nonempty)
        [IO.File]::WriteAllText((Join-Path $nonempty 'keep.txt'), 'existing')
        Assert-DeploymentThrows { & (Join-Path $projectPath 'Tools/Initialize-Store.ps1') -Root $nonempty } '비어 있는 새 폴더'
        Assert-Deployment ([IO.File]::ReadAllText((Join-Path $nonempty 'keep.txt')) -ceq 'existing') 'unrelated file preserved'
    }
    Test-DeploymentCase 'Release contains exact modules, XAML and version-specific launchers' {
        $release = & (Join-Path $projectPath 'Tools/Publish-Release.ps1') -Destination $deploymentPath -DataRoot $storePath
        Assert-Deployment ($release.version -eq '0.1.10') 'default release version'
        foreach ($relative in @('Start.ps1', 'Modules/Storage.psm1', 'Modules/Workflow.psm1', 'Modules/Export.psm1', 'Views/Main.xaml', 'Views/Login.xaml', 'Ui/Admin.ps1', 'Ui/Layout.ps1', 'Views/Admin.xaml')) {
            $sourceFile = Join-Path (Join-Path $projectPath 'App') $relative
            $copiedFile = Join-Path $release.release $relative
            Assert-Deployment ([IO.File]::Exists($copiedFile)) "published file $relative"
            Assert-Deployment ((Get-FileHash -LiteralPath $sourceFile -Algorithm SHA256).Hash -ceq (Get-FileHash -LiteralPath $copiedFile -Algorithm SHA256).Hash) "identical copy $relative"
        }
        $launchText = [IO.File]::ReadAllText((Join-Path $deploymentPath 'Launch.ps1'))
        Assert-Deployment ($launchText.Contains('App\releases\0.1.10\Start.ps1')) 'launcher pins exact release'
        Assert-Deployment ($launchText.Contains($storePath.Replace("'", "''"))) 'launcher pins production data root'
        Assert-Deployment ($launchText.Contains('param([switch]$AdminMode)') -and $launchText.Contains('-AdminMode:$AdminMode')) 'shared PowerShell launcher forwards its switch'
        Assert-Deployment ($release.launcher -ceq (Join-Path $deploymentPath '실행.cmd') -and $release.adminLauncher -ceq (Join-Path $deploymentPath '관리자용.cmd') -and $release.userLauncher -ceq (Join-Path $deploymentPath '일반사용자용.cmd')) 'launcher metadata includes legacy, administrator and user paths'
        $commandText = [IO.File]::ReadAllText($release.userLauncher)
        $adminCommandText = [IO.File]::ReadAllText($release.adminLauncher)
        $legacyCommandText = [IO.File]::ReadAllText($release.launcher)
        Assert-Deployment ($commandText -ceq $legacyCommandText) 'legacy launcher starts ordinary user mode'
        Assert-Deployment ($commandText -notmatch '(?i)-AdminMode' -and $adminCommandText.Replace(' -AdminMode', '') -ceq $commandText) 'CMD role differs only by administrator switch'
        foreach ($entry in @($release.launcher, $release.adminLauncher, $release.userLauncher)) {
            $entryText = [IO.File]::ReadAllText($entry)
            Assert-Deployment ($entryText.Contains('-STA') -and $entryText.Contains('Launch.ps1') -and $entryText.Contains('setlocal DisableDelayedExpansion') -and $entryText.Contains('if errorlevel 1 pause') -and $entryText -match '(?i)-ExecutionPolicy\s+RemoteSigned' -and $entryText -notmatch '(?i)-ExecutionPolicy\s+Bypass') 'STA launchers use process RemoteSigned and retain error display'
            Assert-Deployment (@([IO.File]::ReadAllBytes($entry) | Where-Object { $_ -gt 127 }).Count -eq 0) 'CMD contents are ASCII for every Windows code page'
        }
        $bom = [IO.File]::ReadAllBytes((Join-Path $deploymentPath 'Launch.ps1'))
        Assert-Deployment ($bom.Length -gt 3 -and $bom[0] -eq 239 -and $bom[1] -eq 187 -and $bom[2] -eq 191) 'PowerShell launcher UTF-8 BOM'
    }
    Test-DeploymentCase 'Source demo launchers choose administrator or ordinary user mode' {
        $userCommand = [IO.File]::ReadAllText((Join-Path $projectPath '일반사용자용.cmd')).Replace("`r`n", "`n").Trim()
        $adminCommand = [IO.File]::ReadAllText((Join-Path $projectPath '관리자용.cmd')).Replace("`r`n", "`n").Trim()
        $legacyCommand = [IO.File]::ReadAllText((Join-Path $projectPath '실행.cmd')).Replace("`r`n", "`n").Trim()
        Assert-Deployment ($userCommand -ceq $legacyCommand -and $adminCommand.Replace(' -AdminMode', '') -ceq $userCommand) 'all demo launchers use the same store with only the administrator switch differing'
        Assert-Deployment ($userCommand.Contains('%~dp0App\Start.ps1') -and $userCommand -match '(?i)-Demo' -and $userCommand -notmatch '(?i)-AdminMode') 'source user launcher enters demo selection mode'
        foreach ($entryName in @('관리자용.cmd', '일반사용자용.cmd', '실행.cmd', '운영실행.cmd.example')) {
            $entryBytes = [IO.File]::ReadAllBytes((Join-Path $projectPath $entryName))
            Assert-Deployment (@($entryBytes | Where-Object { $_ -gt 127 }).Count -eq 0) "ASCII source launcher $entryName"
        }
        $productionExample = [IO.File]::ReadAllText((Join-Path $projectPath '운영실행.cmd.example'))
        Assert-Deployment ($productionExample.Contains('-DataRoot "\\FILESERVER\AIReview"') -and $productionExample.Contains('append -AdminMode')) 'production example documents the same administrator flag'
    }
    Test-DeploymentCase 'Shared launcher forwards mode and the identical store at runtime' {
        $probeRoot = Join-Path $testRoot 'LauncherProbe'
        $probeRelease = Join-Path $probeRoot 'App/releases/0.1.10'
        [void][IO.Directory]::CreateDirectory($probeRelease)
        [IO.File]::Copy((Join-Path $deploymentPath 'Launch.ps1'), (Join-Path $probeRoot 'Launch.ps1'))
        $probeScript = 'param([switch]$AdminMode,[string]$DataRoot)' + "`r`n" + '[pscustomobject]@{ adminMode=[bool]$AdminMode; dataRoot=$DataRoot } | ConvertTo-Json -Compress' + "`r`n"
        [IO.File]::WriteAllText((Join-Path $probeRelease 'Start.ps1'), $probeScript, (New-Object Text.UTF8Encoding($true)))
        $userResult = & (Join-Path $probeRoot 'Launch.ps1') | ConvertFrom-Json
        $adminResult = & (Join-Path $probeRoot 'Launch.ps1') -AdminMode | ConvertFrom-Json
        Assert-Deployment (-not $userResult.adminMode -and $adminResult.adminMode) 'administrator switch reaches Start.ps1'
        Assert-Deployment ($userResult.dataRoot -ceq $storePath -and $adminResult.dataRoot -ceq $storePath) 'both entry points share the pinned production data root'
    }
    Test-DeploymentCase 'Generated launcher parses on Windows PowerShell 5.1' {
        $launchLiteral = "'" + (Join-Path $deploymentPath 'Launch.ps1').Replace("'", "''") + "'"
        $code = '$t=$null; $e=$null; $null=[Management.Automation.Language.Parser]::ParseFile(' + $launchLiteral + ',[ref]$t,[ref]$e); if (@($e).Count) { [Console]::Error.WriteLine(($e -join ''; '')); exit 1 }; exit 0'
        $start = New-Object Diagnostics.ProcessStartInfo
        $start.FileName = (Get-Command powershell.exe -ErrorAction Stop).Source
        $start.Arguments = '-NoProfile -NonInteractive -EncodedCommand ' + [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($code))
        $start.UseShellExecute = $false
        $start.CreateNoWindow = $true
        $start.RedirectStandardError = $true
        $process = New-Object Diagnostics.Process
        $process.StartInfo = $start
        try {
            [void]$process.Start()
            if (-not $process.WaitForExit(10000)) { $process.Kill(); throw 'PowerShell parser child timed out.' }
            $errorText = $process.StandardError.ReadToEnd()
            Assert-Deployment ($process.ExitCode -eq 0) "Windows PowerShell launcher parser: $errorText"
        } finally { $process.Dispose() }
    }
    Test-DeploymentCase 'Existing deployment and destination inside source are rejected' {
        Assert-DeploymentThrows { & (Join-Path $projectPath 'Tools/Publish-Release.ps1') -Destination $deploymentPath -DataRoot $storePath } '덮어쓰지 않습니다'
        $unsafeDestination = Join-Path (Join-Path $projectPath 'App') 'MustNeverBeCreated'
        Assert-DeploymentThrows { & (Join-Path $projectPath 'Tools/Publish-Release.ps1') -Destination $unsafeDestination -DataRoot $storePath } '원본 App 폴더 안'
        Assert-Deployment (-not (Test-Path -LiteralPath $unsafeDestination)) 'unsafe target never created'
    }
    Test-DeploymentCase 'Every reserved launcher rejects an existing file before release creation' {
        foreach ($entryName in @('실행.cmd', '관리자용.cmd', '일반사용자용.cmd', 'Launch.ps1')) {
            $reservedTarget = Join-Path $testRoot ('Reserved-' + [guid]::NewGuid().ToString('N'))
            [void][IO.Directory]::CreateDirectory($reservedTarget)
            $existingEntry = Join-Path $reservedTarget $entryName
            [IO.File]::WriteAllText($existingEntry, 'keep-existing-launcher')
            Assert-DeploymentThrows { & (Join-Path $projectPath 'Tools/Publish-Release.ps1') -Destination $reservedTarget -DataRoot $storePath } '덮어쓰지 않습니다'
            Assert-Deployment ([IO.File]::ReadAllText($existingEntry) -ceq 'keep-existing-launcher') "preserved reserved launcher $entryName"
            Assert-Deployment (-not (Test-Path -LiteralPath (Join-Path $reservedTarget 'App/releases/0.1.10'))) 'release is not written when a reserved launcher exists'
        }
    }
    Write-Host ('Deployment tests: {0} passed (PowerShell {1}).' -f $script:passed, $PSVersionTable.PSVersion)
}
finally {
    if ($KeepArtifacts) { Write-Host "Artifacts: $testRoot" }
    elseif ([IO.Directory]::Exists($testRoot)) {
        $resolved = [IO.Path]::GetFullPath($testRoot)
        $tempPrefix = [IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\') + '\'
        if (-not $resolved.StartsWith($tempPrefix, [StringComparison]::OrdinalIgnoreCase) -or
            [IO.Path]::GetFileName($resolved) -notmatch '^AiReview-DeploymentTests-[a-f0-9]{32}$') { throw 'Unsafe test cleanup boundary.' }
        Remove-Item -LiteralPath $resolved -Recurse -Force
    }
}
