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
        Assert-Deployment ($release.version -eq '0.1.3') 'default release version'
        foreach ($relative in @('Start.ps1', 'Modules/Storage.psm1', 'Modules/Workflow.psm1', 'Modules/Export.psm1', 'Views/Main.xaml', 'Views/Login.xaml')) {
            $sourceFile = Join-Path (Join-Path $projectPath 'App') $relative
            $copiedFile = Join-Path $release.release $relative
            Assert-Deployment ([IO.File]::Exists($copiedFile)) "published file $relative"
            Assert-Deployment ((Get-FileHash -LiteralPath $sourceFile -Algorithm SHA256).Hash -ceq (Get-FileHash -LiteralPath $copiedFile -Algorithm SHA256).Hash) "identical copy $relative"
        }
        $launchText = [IO.File]::ReadAllText((Join-Path $deploymentPath 'Launch.ps1'))
        Assert-Deployment ($launchText.Contains('App\releases\0.1.3\Start.ps1')) 'launcher pins exact release'
        Assert-Deployment ($launchText.Contains($storePath.Replace("'", "''"))) 'launcher pins production data root'
        $commandText = [IO.File]::ReadAllText((Join-Path $deploymentPath '실행.cmd'))
        Assert-Deployment ($commandText.Contains('-STA') -and $commandText.Contains('Launch.ps1') -and $commandText -match '(?i)-ExecutionPolicy\s+RemoteSigned' -and $commandText -notmatch '(?i)-ExecutionPolicy\s+Bypass') 'STA launcher uses process RemoteSigned without Bypass'
        $bom = [IO.File]::ReadAllBytes((Join-Path $deploymentPath 'Launch.ps1'))
        Assert-Deployment ($bom.Length -gt 3 -and $bom[0] -eq 239 -and $bom[1] -eq 187 -and $bom[2] -eq 191) 'PowerShell launcher UTF-8 BOM'
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
