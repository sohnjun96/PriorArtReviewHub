$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0
if ([Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') { throw 'Run this test with powershell.exe -STA or pwsh.exe -STA.' }
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase
$appPath = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../App'))
$scriptPaths = @((Join-Path $appPath 'Start.ps1')) + @(Get-ChildItem -LiteralPath (Join-Path $appPath 'Ui') -Filter '*.ps1' -File | ForEach-Object { $_.FullName })
$sourceParts = @()
foreach ($scriptPath in $scriptPaths) {
    $part = [IO.File]::ReadAllText($scriptPath)
    $tokens = $null
    $parseErrors = $null
    $null = [Management.Automation.Language.Parser]::ParseInput($part, [ref]$tokens, [ref]$parseErrors)
    if (@($parseErrors).Count) { throw ('GUI PowerShell syntax errors in {0}: {1}' -f $scriptPath, ($parseErrors -join '; ')) }
    $sourceParts += $part
}
$source = $sourceParts -join [Environment]::NewLine
function Read-TestView([string]$Name) {
    $reader = [Xml.XmlReader]::Create((Join-Path $appPath ('Views/' + $Name + '.xaml')))
    try { return [Windows.Markup.XamlReader]::Load($reader) }
    finally { $reader.Close() }
}
function Find-TestControl([string]$Name) {
    $control = $window.FindName($Name)
    if ($null -eq $control) { $control = $adminView.FindName($Name) }
    if ($null -eq $control) { $control = $statisticsView.FindName($Name) }
    return $control
}
function Get-TestDescendants($Root) {
    if ($null -eq $Root) { return }
    $Root
    $childCount = [Windows.Media.VisualTreeHelper]::GetChildrenCount($Root)
    for ($index = 0; $index -lt $childCount; $index++) {
        Get-TestDescendants ([Windows.Media.VisualTreeHelper]::GetChild($Root, $index))
    }
}
function Assert-ChoiceColumn($Grid, [int]$ColumnIndex, [string]$Property, [string[]]$ExpectedValues, [bool]$RequireNoResultsGuard = $false) {
    $column = $Grid.Columns[$ColumnIndex]
    if ($column -isnot [Windows.Controls.DataGridTemplateColumn] -or $null -eq $column.CellTemplate) {
        throw "Column '$($column.Header)' must show a dropdown without first entering cell edit mode."
    }
    $content = $column.CellTemplate.LoadContent()
    $combo = @(Get-TestDescendants $content | Where-Object { $_ -is [Windows.Controls.ComboBox] }) | Select-Object -First 1
    if ($null -eq $combo -or $combo.IsEditable -or $combo.SelectedValuePath -cne 'Tag') {
        throw "Column '$($column.Header)' must have a fixed dropdown using item values."
    }
    $binding = [Windows.Data.BindingOperations]::GetBinding($combo, [Windows.Controls.Primitives.Selector]::SelectedValueProperty)
    if ($null -eq $binding -or $binding.Path.Path -cne $Property -or $binding.Mode -ne [Windows.Data.BindingMode]::TwoWay -or $binding.UpdateSourceTrigger -ne [Windows.Data.UpdateSourceTrigger]::PropertyChanged) {
        throw "Column '$($column.Header)' must write its selected value immediately into '$Property'."
    }
    $values = @($combo.Items | ForEach-Object { [string]$_.Tag })
    if (($values -join '|') -cne ($ExpectedValues -join '|')) {
        throw "Column '$($column.Header)' has unexpected choices: $($values -join ', ')."
    }
    # Always visible template controls need a separate read-only session guard.
    $hasReadOnlyGuard = $false
    $hasNoResultsGuard = $false
    $style = $combo.Style
    while ($null -ne $style) {
        foreach ($trigger in @($style.Triggers)) {
            if ($trigger -is [Windows.DataTrigger] -and $trigger.Binding -is [Windows.Data.Binding] -and $trigger.Binding.Path.Path -ceq 'IsReadOnly' -and $trigger.Binding.RelativeSource.AncestorType -eq [Windows.Controls.DataGrid] -and [string]$trigger.Value -ieq 'True') {
                foreach ($setter in @($trigger.Setters)) {
                    if ($setter -is [Windows.Setter] -and $setter.Property -eq [Windows.UIElement]::IsEnabledProperty -and [string]$setter.Value -ieq 'False') { $hasReadOnlyGuard = $true }
                }
            }
            if ($trigger -is [Windows.DataTrigger] -and $trigger.Binding -is [Windows.Data.Binding] -and $trigger.Binding.Path.Path -ceq 'noResults' -and [string]$trigger.Value -ieq 'True') {
                foreach ($setter in @($trigger.Setters)) {
                    if ($setter -is [Windows.Setter] -and $setter.Property -eq [Windows.UIElement]::IsEnabledProperty -and [string]$setter.Value -ieq 'False') { $hasNoResultsGuard = $true }
                }
            }
        }
        $style = $style.BasedOn
    }
    if (-not $hasReadOnlyGuard) { throw "Column '$($column.Header)' leaves its dropdown editable in a read-only grid." }
    if ($RequireNoResultsGuard -and -not $hasNoResultsGuard) { throw "Column '$($column.Header)' allows ranking an empty search result." }
}
function Assert-AdminChoiceColumn($Grid, [int]$ColumnIndex, [string]$Field, [string]$ValuePath) {
    $column = $Grid.Columns[$ColumnIndex]
    if ($column -isnot [Windows.Controls.DataGridTemplateColumn] -or $null -eq $column.CellTemplate) { throw "Administrator column '$($column.Header)' must show a dropdown immediately." }
    $content = $column.CellTemplate.LoadContent()
    $combo = @(Get-TestDescendants $content | Where-Object { $_ -is [Windows.Controls.ComboBox] }) | Select-Object -First 1
    if ($null -eq $combo -or $combo.IsEditable -or $combo.DisplayMemberPath -cne 'displayName' -or $combo.SelectedValuePath -cne $ValuePath) { throw "Administrator column '$($column.Header)' must select a named record instead of entering an ID." }
    $binding = [Windows.Data.BindingOperations]::GetBinding($combo, [Windows.Controls.Primitives.Selector]::SelectedValueProperty)
    if ($null -eq $binding -or $binding.Path.Path -cne $Field -or $binding.Mode -ne [Windows.Data.BindingMode]::TwoWay -or $binding.UpdateSourceTrigger -ne [Windows.Data.UpdateSourceTrigger]::PropertyChanged) { throw "Administrator column '$($column.Header)' must apply the selected record immediately." }
}
$window = Read-TestView Main
try {
    $hostControl = $window.FindName('AdminContentHost')
    if ($null -eq $hostControl -or $hostControl -isnot [Windows.Controls.ContentControl]) { throw 'Main view must host the administrator forms.' }
    $script:appDirectory = $appPath
    $script:ui = @{AdminContentHost=$hostControl}
    . (Join-Path $appPath 'Ui/Admin.ps1')
    Initialize-AdminUi
    $adminView = $script:adminView
    $script:ui.StatisticsHost=$window.FindName('StatisticsHost')
    . (Join-Path $appPath 'Ui/Statistics.ps1')
    Initialize-StatisticsUi
    $statisticsView=$script:statisticsView
    $controlReferences = @([regex]::Matches($source, '\$ui\.([A-Za-z][A-Za-z0-9_]*)') | ForEach-Object { $_.Groups[1].Value } | Sort-Object -Unique)
    foreach ($name in $controlReferences) {
        if ($null -eq (Find-TestControl $name)) { throw "GUI scripts refer to missing XAML control '$name'." }
    }
    $eventReferences = @([regex]::Matches($source, '\$ui\.([A-Za-z][A-Za-z0-9_]*)\.Add_([A-Za-z][A-Za-z0-9_]*)\('))
    foreach ($reference in $eventReferences) {
        $name = $reference.Groups[1].Value
        $eventName = $reference.Groups[2].Value
        $control = Find-TestControl $name
        if ($null -eq $control.GetType().GetEvent($eventName)) { throw "Control '$name' has no event '$eventName'." }
    }
    foreach ($gridName in @('DocumentsGrid', 'RankGrid', 'RatingsGrid', 'SubmissionRowsGrid', 'PeriodMembersGrid', 'PeriodAssignmentsGrid')) {
        $grid = Find-TestControl $gridName
        if ($grid.CanUserAddRows -or $grid.CanUserDeleteRows -or $grid.AutoGenerateColumns) { throw "Grid '$gridName' permits uncontrolled row/column creation." }
    }
    $headers = @($window.FindName('MainTabs').Items | ForEach-Object { $_.Header })
    if (($headers -join '|') -cne '내 심사 업무|AI 담당 업무|통계·현황|운영 관리') { throw 'Examiner work, assigned AI work and administrator work must have separate workspaces.' }
    $mainNames = @($window.FindName('MainTabs').Items | ForEach-Object { $_.Name })
    if (($mainNames -join '|') -cne 'PersonalWorkTab|DeliveryTab|StatisticsTab|AdminTab') { throw 'The role workspaces must remain independent top-level tabs.' }
    $personalTabs = $window.FindName('PersonalTabs')
    if ($null -eq $personalTabs -or (($personalTabs.Items | ForEach-Object { $_.Header }) -join '|') -cne '1. 대상 건 선정|2. AI 검색 결과 평가') { throw 'Target selection and result evaluation must be separate examiner steps.' }
    if ((($personalTabs.Items | ForEach-Object { $_.Name }) -join '|') -cne 'TargetSelectionTab|EvaluationTab') { throw 'The personal workspace must expose separate selection and evaluation tabs.' }
    if ($null -ne $window.FindName('TargetTabs') -or $null -ne $window.FindName('GoDeliveryButton')) { throw 'Target selection must not contain a mixed delivery lookup tab or a shortcut into assigned AI work.' }
    $evaluationTabs = $window.FindName('EvaluationTabs')
    if ($null -eq $evaluationTabs -or $evaluationTabs.Items.Count -ne 2 -or $evaluationTabs.SelectedIndex -ne 1) { throw 'The evaluation step must separate delivered results from evaluation entry and open evaluation entry by default.' }
    if ((($evaluationTabs.Items | ForEach-Object { $_.Header }) -join '|') -cne '결과 조회|평가 입력') { throw 'Delivered results and evaluation entry must have distinct sub-tabs.' }
    [xml]$mainXml = [IO.File]::ReadAllText((Join-Path $appPath 'Views/Main.xaml'))
    $xmlNamespaces = New-Object Xml.XmlNamespaceManager($mainXml.NameTable)
    $xmlNamespaces.AddNamespace('w', 'http://schemas.microsoft.com/winfx/2006/xaml/presentation')
    $xmlNamespaces.AddNamespace('x', 'http://schemas.microsoft.com/winfx/2006/xaml')
    foreach ($controlName in @('PublishedGrid', 'RatingsGrid', 'RankGrid')) {
        if ($null -eq $mainXml.SelectSingleNode("//w:TabItem[@x:Name='EvaluationTab']//*[@x:Name='$controlName']", $xmlNamespaces)) { throw "Control '$controlName' must belong to the evaluation step." }
    }
    if ($null -ne $mainXml.SelectSingleNode("//w:TabItem[@x:Name='TargetSelectionTab']//*[@x:Name='PublishedGrid' or @x:Name='RatingsGrid' or @x:Name='DocumentsGrid']", $xmlNamespaces)) { throw 'The selection step must only register and select target cases.' }
    if ($null -eq $mainXml.SelectSingleNode("//w:TabItem[@x:Name='DeliveryTab']//*[@x:Name='DocumentsGrid']", $xmlNamespaces)) { throw 'AI result entry must belong to the assigned AI workspace.' }
    $ratingsGrid = $window.FindName('RatingsGrid')
    $ratingHeaders = @($ratingsGrid.Columns | ForEach-Object { $_.Header })
    if (($ratingHeaders -join '|') -cne '조사 방식|문헌|점수(0~5점)|인용여부(O/X)|비고') { throw 'Evaluation grid must contain exactly the requested five columns.' }
    Assert-ChoiceColumn $ratingsGrid 2 scoreText @('', '0', '1', '2', '3', '4', '5')
    Assert-ChoiceColumn $ratingsGrid 3 citedText @('', 'O', 'X')
    Assert-ChoiceColumn ($window.FindName('RankGrid')) 1 choice @('', '대기 중', '1', '2', '3', '순위 외', '평가 제외', '비교 불가') $true
    $membersGrid = Find-TestControl PeriodMembersGrid
    $assignmentsGrid = Find-TestControl PeriodAssignmentsGrid
    if ($membersGrid.Columns.Count -ne 2 -or $assignmentsGrid.Columns.Count -ne 3) { throw 'Month settings must show named participants and AI assignments.' }
    Assert-AdminChoiceColumn $membersGrid 0 userId userId
    Assert-AdminChoiceColumn $membersGrid 1 groupId groupId
    Assert-AdminChoiceColumn $assignmentsGrid 0 userId userId
    Assert-AdminChoiceColumn $assignmentsGrid 1 groupId groupId
    Assert-AdminChoiceColumn $assignmentsGrid 2 sourceId sourceId
    foreach ($name in @('AdminUsersGrid','AdminGroupsGrid','GroupMembersGrid')) {
        $grid=Find-TestControl $name
        if ($grid -isnot [Windows.Controls.DataGrid] -or -not $grid.IsReadOnly -or $grid.CanUserAddRows) {throw "Administrator list '$name' must select records without editing rows directly."}
    }
    foreach ($name in @('AdminKindCombo', 'AdminBureauFilter', 'AdminDivisionFilter', 'AdminSourceCombo', 'AdminPeriodCombo', 'SourceTypeCombo', 'PeriodYearCombo', 'PeriodMonthCombo')) {
        $combo = Find-TestControl $name
        if ($combo -isnot [Windows.Controls.ComboBox] -or $combo.IsEditable) { throw "Administrator selector '$name' must use predefined dropdown choices." }
    }
    foreach ($name in @('UserBureauBox', 'UserDivisionBox')) {
        $combo = Find-TestControl $name
        if ($combo -isnot [Windows.Controls.ComboBox] -or -not $combo.IsEditable) { throw "Organization selector '$name' must support choosing an existing value or entering a new value." }
    }
    if ($null -ne (Find-TestControl 'UserSidBox')) { throw 'User registration must not request a Windows SID.' }
    if ($null -ne (Find-TestControl 'AdminJsonBox')) { throw 'Administrator settings must use forms instead of a JSON editor.' }
    $loginWindow = Read-TestView Login
    try {
        foreach ($controlName in @('LoginUsersList', 'LoginSearchBox', 'LoginButton', 'LoginRefreshButton', 'LoginStatusText')) {
            if (-not $loginWindow.FindName($controlName)) { throw "Login control is missing: $controlName" }
        }
    } finally { $loginWindow.Close() }
    Write-Host ('GUI structure passed: {0} referenced controls, {1} event bindings, separate role workspaces and examiner steps, administrator forms, immediate dropdown bindings, valid WPF XAML (PowerShell {2}).' -f $controlReferences.Count, $eventReferences.Count, $PSVersionTable.PSVersion)
} finally { $window.Close() }
