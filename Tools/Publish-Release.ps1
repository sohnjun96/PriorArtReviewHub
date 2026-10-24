[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$Destination,
    [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$DataRoot,
    [ValidatePattern('^\d+\.\d+\.\d+(?:-[A-Za-z0-9.-]+)?$')][string]$Version = '0.1.18'
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0
$projectPath = Split-Path $PSScriptRoot -Parent
$sourcePath = [IO.Path]::GetFullPath((Join-Path $projectPath 'App')).TrimEnd('\', '/')
Import-Module (Join-Path $sourcePath 'Modules/Storage.psm1') -Force
$destinationPath = [IO.Path]::GetFullPath($Destination)
$dataPath = [IO.Path]::GetFullPath($DataRoot)
if ($destinationPath -match '[\r\n\x00]' -or $dataPath -match '[\r\n\x00]') { throw '경로에 줄바꿈 또는 제어문자를 사용할 수 없습니다.' }
if ([IO.File]::Exists($destinationPath)) { throw '배포 대상 경로가 파일입니다.' }
if ($destinationPath.Equals($sourcePath, [StringComparison]::OrdinalIgnoreCase) -or
    $destinationPath.StartsWith($sourcePath + '\', [StringComparison]::OrdinalIgnoreCase)) {
    throw '원본 App 폴더 안으로 배포할 수 없습니다.'
}
$releasePath = Get-SafePath -Root $destinationPath -Segments @('App', 'releases', $Version)
$entryPath = Get-SafePath -Root $destinationPath -Segments @('실행.cmd')
$adminEntryPath = Get-SafePath -Root $destinationPath -Segments @('관리자용.cmd')
$userEntryPath = Get-SafePath -Root $destinationPath -Segments @('일반사용자용.cmd')
$launchPath = Get-SafePath -Root $destinationPath -Segments @('Launch.ps1')
$publishLock = Get-SafePath -Root $destinationPath -Segments @('.publish.lock')
$utf8Bom = New-Object Text.UTF8Encoding($true)

function Write-NewDeploymentFile {
    param([string]$Path, [string]$Text, [Text.Encoding]$Encoding)
    $stream = New-Object IO.FileStream($Path, [IO.FileMode]::CreateNew, [IO.FileAccess]::Write, [IO.FileShare]::None)
    try {
        $preamble = $Encoding.GetPreamble()
        if ($preamble.Length) { $stream.Write($preamble, 0, $preamble.Length) }
        $bytes = $Encoding.GetBytes($Text)
        $stream.Write($bytes, 0, $bytes.Length)
        $stream.Flush($true)
    } finally { $stream.Dispose() }
}

Invoke-FileLock -Path $publishLock -TimeoutMilliseconds 10000 -Action {
    foreach ($reserved in @($releasePath, $entryPath, $adminEntryPath, $userEntryPath, $launchPath)) {
        if (Test-Path -LiteralPath $reserved) { throw "기존 배포 파일을 덮어쓰지 않습니다: $reserved" }
    }
    $sourceEntries = @(Get-ChildItem -LiteralPath $sourcePath -Force -Recurse)
    foreach ($item in $sourceEntries) {
        if (($item.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) { throw "원본에 연결 폴더/파일이 있습니다: $($item.FullName)" }
    }
    $stagingPath = Get-SafePath -Root $destinationPath -Segments @('App', 'releases', ('.pending-release-' + [guid]::NewGuid().ToString('N')))
    [void][IO.Directory]::CreateDirectory($stagingPath)
    try {
        foreach ($item in $sourceEntries) {
            $relative = $item.FullName.Substring($sourcePath.Length + 1)
            $copyPath = Get-SafePath -Root $stagingPath -Segments @($relative -split '[\\/]')
            if ($item.PSIsContainer) { [void][IO.Directory]::CreateDirectory($copyPath) }
            else {
                [void][IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($copyPath))
                [IO.File]::Copy($item.FullName, $copyPath, $false)
            }
        }
        if (-not [IO.File]::Exists((Join-Path $stagingPath 'Start.ps1'))) { throw '배포 원본에 Start.ps1이 없습니다.' }
        # Both absolute paths above were built inside the explicit destination.
        # Directory.Move refuses an existing version directory.
        [IO.Directory]::Move($stagingPath, $releasePath)
        $escapedDataPath = $dataPath.Replace("'", "''")
        $launchText = 'param([switch]$AdminMode)' + "`r`n" +
            '$ErrorActionPreference = ''Stop''' + "`r`n" +
            '& (Join-Path $PSScriptRoot ''App\releases\' + $Version + '\Start.ps1'') -AdminMode:$AdminMode -DataRoot ''' + $escapedDataPath + "'`r`n"
        Write-NewDeploymentFile -Path $launchPath -Text $launchText -Encoding $utf8Bom
        $commandText = "@echo off`r`nsetlocal DisableDelayedExpansion`r`npowershell.exe -NoLogo -NoProfile -STA -ExecutionPolicy RemoteSigned -File ""%~dp0Launch.ps1""`r`nif errorlevel 1 pause`r`n"
        $adminCommandText = "@echo off`r`nsetlocal DisableDelayedExpansion`r`npowershell.exe -NoLogo -NoProfile -STA -ExecutionPolicy RemoteSigned -File ""%~dp0Launch.ps1"" -AdminMode`r`nif errorlevel 1 pause`r`n"
        Write-NewDeploymentFile -Path $entryPath -Text $commandText -Encoding ([Text.Encoding]::ASCII)
        Write-NewDeploymentFile -Path $userEntryPath -Text $commandText -Encoding ([Text.Encoding]::ASCII)
        Write-NewDeploymentFile -Path $adminEntryPath -Text $adminCommandText -Encoding ([Text.Encoding]::ASCII)
        [pscustomobject]@{ destination=$destinationPath; release=$releasePath; launcher=$entryPath; adminLauncher=$adminEntryPath; userLauncher=$userEntryPath; dataRoot=$dataPath; version=$Version }
    }
    finally {
        if ([IO.Directory]::Exists($stagingPath)) {
            $stagingFullPath = [IO.Path]::GetFullPath($stagingPath)
            $releaseParent = [IO.Path]::GetDirectoryName($releasePath).TrimEnd('\') + '\'
            if (-not $stagingFullPath.StartsWith($releaseParent, [StringComparison]::OrdinalIgnoreCase) -or
                [IO.Path]::GetFileName($stagingFullPath) -notmatch '^\.pending-release-[a-f0-9]{32}$') {
                throw '임시 배포 폴더의 안전한 경계를 확인할 수 없습니다.'
            }
            Remove-Item -LiteralPath $stagingFullPath -Recurse -Force
        }
    }
}
