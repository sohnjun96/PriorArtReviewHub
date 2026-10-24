$ErrorActionPreference='Stop'
Set-StrictMode -Version 2.0
Add-Type -AssemblyName PresentationFramework
. (Join-Path $PSScriptRoot '../App/Ui/Export.ps1')
. (Join-Path $PSScriptRoot '../App/Ui/Statistics.ps1')
$testRoot=Join-Path ([IO.Path]::GetTempPath()) ('AiReview-ExportDialog-'+[guid]::NewGuid().ToString('N'))
[void][IO.Directory]::CreateDirectory($testRoot)
$script:dialogAccepted=$false;$script:chosenPath=Join-Path $testRoot '내가 정한 파일.xlsx'
$script:bundleAccepted=$false;$script:bundlePromptCount=0;$script:workCount=0;$script:dialogCount=0
$script:context=@{};$script:selectedCase=[pscustomobject]@{applicationNumber='10-2026-1234567'}
$script:scope='Mine'
function Assert-Dialog([bool]$Condition,[string]$Message) {if (-not $Condition) {throw "FAIL: $Message"}}
# Exercise the real SaveFileDialog configuration, replacing only its modal interaction.
function Show-ExportSaveDialog($Dialog) {
    $script:dialogCount++
    $script:observedDialog=$Dialog
    $script:defaultName=$Dialog.FileName
    $Dialog.FileName=$script:chosenPath
    return $script:dialogAccepted
}
function Confirm-ExportBundleOverwrite([string]$Path) {$script:bundlePromptCount++;$script:bundlePromptPath=$Path;return $script:bundleAccepted}
function Get-SelectedPeriod {return '2026-10'}
function Get-WorkspaceScope {return $script:scope}
function Set-Status([string]$Text) {$script:status=$Text}
function Invoke-UiWork($Work,$Arguments,$Completed) {$script:workCount++;$script:work=$Work;$script:workArguments=$Arguments}
try {
    $result=Select-ExcelExportTarget -FileName '월별 결과.xlsx'
    Assert-Dialog ($null -eq $result) 'Cancel returns no export target.'
    Assert-Dialog ($script:observedDialog -is [Microsoft.Win32.SaveFileDialog] -and $script:observedDialog.OverwritePrompt -and $script:observedDialog.AddExtension -and $script:observedDialog.CheckPathExists) 'Native save dialog checks path, adds extension, and confirms overwrite.'
    Assert-Dialog ($script:observedDialog.DefaultExt.TrimStart('.') -eq 'xlsx' -and $script:observedDialog.Filter -match '\*\.xlsx') 'Excel file type is selected.'
    foreach ($scopeValue in @('Mine','Assigned','All')) {
        $script:scope=$scopeValue
        Export-Current $false
        Export-Current $true
        Export-Current $true 'C001'
    }
    $script:statisticsFilter=[pscustomobject]@{groups=@();accounts=@();cases=@([pscustomobject]@{applicationNumber='10-2026-1234567'});sources=@()}
    $script:statisticsSnapshot=[pscustomobject]@{period='2026-10';capturedAt='2026-10-24'}
    $script:ui=@{StatsScopeText=[pscustomobject]@{Text='2026-10 · 그룹 1 · 선택 계정'}}
    Export-StatisticsView
    Assert-Dialog ($script:workCount -eq 0 -and @(Get-ChildItem -LiteralPath $testRoot -Force).Count -eq 0) 'Cancelling every export route creates no files and starts no background work.'
    $script:dialogAccepted=$true
    $result=Select-ExcelExportTarget -FileName '월별 결과.xlsx'
    Assert-Dialog ($result.Path -eq $script:chosenPath -and -not $result.OverwriteWorkbook -and -not $result.OverwriteBundle) 'The chosen Unicode path and filename are kept exactly.'
    $null=Select-ExcelExportTarget -FileName '월별 결과.xlsx'
    Assert-Dialog ($script:observedDialog.InitialDirectory -eq $testRoot) 'The next dialog starts in the last selected directory.'
    [IO.File]::WriteAllText($script:chosenPath,'original workbook')
    $bundlePath=[IO.Path]::ChangeExtension($script:chosenPath,'.zip')
    [IO.File]::WriteAllText($bundlePath,'original bundle')
    $result=Select-ExcelExportTarget -FileName '결과.xlsx' -IncludeAttachments $true
    Assert-Dialog ($null -eq $result -and $script:bundlePromptPath -eq $bundlePath) 'Declining companion ZIP overwrite cancels the whole export.'
    Assert-Dialog ([IO.File]::ReadAllText($script:chosenPath) -eq 'original workbook' -and [IO.File]::ReadAllText($bundlePath) -eq 'original bundle') 'Declining overwrite preserves both files.'
    $script:bundleAccepted=$true
    $result=Select-ExcelExportTarget -FileName '결과.xlsx' -IncludeAttachments $true
    Assert-Dialog ($result.OverwriteWorkbook -and $result.OverwriteBundle) 'Confirmed replacements are explicit for each file.'
    $script:scope='Assigned'
    Export-Current $true 'C001'
    Assert-Dialog ($script:workArguments[3].Path -eq $script:chosenPath -and $script:workArguments[4] -eq $true -and $script:workArguments[5] -eq 'C001' -and $script:workArguments[2] -eq 'Assigned') 'Case export forwards selected path, ZIP choice, scope and case.'
    function Export-ReviewSnapshot($Context,$Period,$Scope,$WorkbookPath,$OverwriteWorkbook,$OverwriteBundle,$IncludeAttachments,$CaseId) {
        Assert-Dialog ($WorkbookPath -eq $script:chosenPath -and $OverwriteWorkbook -and $OverwriteBundle -and $IncludeAttachments -and $CaseId -eq 'C001') 'Background snapshot export honors the selected target.'
    }
    $argsForWork=$script:workArguments;& $script:work @argsForWork
    $promptsBefore=$script:bundlePromptCount
    Export-StatisticsView
    Assert-Dialog ($script:bundlePromptCount -eq $promptsBefore -and $script:workArguments[0].Path -eq $script:chosenPath) 'Statistics opens the Excel chooser without a ZIP confirmation.'
    Assert-Dialog (@($script:workArguments[1]['출원별 현황']).Count -eq 1 -and $script:workArguments[1]['조회 기준'][0].scope -eq $ui.StatsScopeText.Text) 'Statistics retains the currently filtered data and scope.'
    function Export-Workbook($Path,$Sheets,$Columns,$Force) {Assert-Dialog ($Path -eq $script:chosenPath -and $Force) 'Statistics background export honors chosen path and confirmed overwrite.'}
    $argsForWork=$script:workArguments;& $script:work @argsForWork
    Write-Output 'PASS: Excel save dialog, all export routes, cancellation, exact filename, remembered folder, and overwrite confirmation.'
} finally {
    $resolved=[IO.Path]::GetFullPath($testRoot)
    $tempPrefix=[IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\')+'\'
    if ($resolved.StartsWith($tempPrefix,[StringComparison]::OrdinalIgnoreCase) -and [IO.Path]::GetFileName($resolved).StartsWith('AiReview-ExportDialog-')) {Remove-Item -LiteralPath $resolved -Recurse -Force}
}
