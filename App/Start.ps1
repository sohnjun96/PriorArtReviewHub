param(
    [string]$DataRoot,
    [switch]$Demo,
    [switch]$AdminMode,
    [switch]$SmokeTest,
    [switch]$ExerciseUiWrites,
    [switch]$ExerciseWorkspaceSwitch,
    [switch]$ExerciseStatistics,
    [ValidateSet('','Selection','Evaluation','AI','Statistics')][string]$SmokeWorkspace='',
    [ValidateSet('','Draft','Confirm')][string]$ExerciseTargetSelection='',
    [string]$ScreenshotPath,
    [string]$SmokeLoginUserId = 'U001',
    [string]$SmokeReportPath
)
$ErrorActionPreference = 'Stop'
if ($ExerciseUiWrites -and (-not $SmokeTest -or -not $Demo)) {throw '화면 쓰기 검증은 -Demo -SmokeTest에서만 실행할 수 있습니다.'}
if ($ExerciseWorkspaceSwitch -and (-not $SmokeTest -or -not $Demo)) {throw '업무 전환 검증은 -Demo -SmokeTest에서만 실행할 수 있습니다.'}
if ($ExerciseStatistics -and (-not $SmokeTest -or -not $Demo -or -not $AdminMode)) {throw '통계 검증은 -Demo -SmokeTest -AdminMode에서만 실행할 수 있습니다.'}
if ($ExerciseTargetSelection -and (-not $SmokeTest -or -not $Demo)) {throw '대상 건 검증은 -Demo -SmokeTest에서만 실행할 수 있습니다.'}
if ($SmokeWorkspace -and -not $SmokeTest) {throw '화면 지정 검증은 -SmokeTest에서만 실행할 수 있습니다.'}
if ([Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') { throw 'PowerShell을 -STA 옵션으로 실행하세요.' }
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase
$script:appDirectory = $PSScriptRoot
foreach ($module in @('Storage','Workflow','Export')) { Import-Module (Join-Path $PSScriptRoot "Modules\$module.psm1") -Force -DisableNameChecking }
if (-not $DataRoot) {
    if ($Demo) { $DataRoot = Join-Path (Split-Path $PSScriptRoot -Parent) 'DemoData' }
    else { throw '운영 모드는 -DataRoot 공유폴더 경로가 필요합니다. 시연하려면 실행.cmd를 이용하세요.' }
}
$DataRoot = [IO.Path]::GetFullPath($DataRoot)
if ($Demo) { Initialize-Demo -Root $DataRoot | Out-Null }
$script:launchMode=if ($AdminMode) {'Admin'} else {'User'}
$script:loginDialogShown=$false
function Write-UiImage($Window,[string]$Path) {
    if (-not $Path) {return}
    $Window.UpdateLayout()
    $bitmap=New-Object Windows.Media.Imaging.RenderTargetBitmap ([int]$Window.ActualWidth),([int]$Window.ActualHeight),96,96,([Windows.Media.PixelFormats]::Pbgra32)
    $bitmap.Render($Window)
    $encoder=New-Object Windows.Media.Imaging.PngBitmapEncoder
    $encoder.Frames.Add([Windows.Media.Imaging.BitmapFrame]::Create($bitmap))
    $stream=[IO.File]::Create($Path);try {$encoder.Save($stream)} finally {$stream.Dispose()}
}
function Show-LoginDialog($Owner=$null) {
    $script:loginDialogShown=$true
    $loginReader=[Xml.XmlReader]::Create((Join-Path $appDirectory 'Views\Login.xaml'))
    try {$loginWindow=[Windows.Markup.XamlReader]::Load($loginReader)} finally {$loginReader.Close()}
    if ($Owner) {$loginWindow.Owner=$Owner;$loginWindow.WindowStartupLocation='CenterOwner'}
    $login=@{window=$loginWindow;root=$DataRoot;context=$null;users=@();error='';smoke=[bool]$SmokeTest;smokeUserId=$SmokeLoginUserId;screenshot=$ScreenshotPath;listCommand=(Get-Command Get-LoginUsers);contextCommand=(Get-Command Get-Context);imageWriter=${function:Write-UiImage}}
    foreach ($controlName in @('LoginSearchBox','LoginUsersList','LoginStatusText','LoginRefreshButton','LoginButton')) {$login[$controlName]=$loginWindow.FindName($controlName)}
    $login.filter={
        $search=$login.LoginSearchBox.Text.Trim()
        $login.LoginUsersList.ItemsSource=@($login.users | Where-Object {-not $search -or ($_.displayName+' '+$_.bureau).IndexOf($search,[StringComparison]::OrdinalIgnoreCase) -ge 0})
        $login.LoginStatusText.Text="활성 사용자 $($login.users.Count)명 · 본인을 선택하고 로그인하세요."
    }.GetNewClosure()
    $login.reload={
        try {$login.users=@(& $login.listCommand -Root $login.root -AccessMode User);& $login.filter}
        catch {$login.LoginStatusText.Text=$_.Exception.Message}
    }.GetNewClosure()
    $login.LoginSearchBox.Add_TextChanged({& $login.filter}.GetNewClosure())
    $login.LoginRefreshButton.Add_Click({& $login.reload}.GetNewClosure())
    $login.LoginButton.Add_Click({
        if (-not $login.LoginUsersList.SelectedItem) {$login.LoginStatusText.Text='목록에서 본인을 선택하세요.';return}
        try {$login.context=& $login.contextCommand -Root $login.root -UserId $login.LoginUsersList.SelectedItem.userId -AccessMode User;$login.window.DialogResult=$true}
        catch {$login.LoginStatusText.Text=$_.Exception.Message;if ($login.smoke) {$login.error=$_.Exception.Message;$login.window.Close()}}
    }.GetNewClosure())
    $login.LoginUsersList.Add_MouseDoubleClick({if ($login.LoginUsersList.SelectedItem) {$login.LoginButton.RaiseEvent((New-Object Windows.RoutedEventArgs([Windows.Controls.Button]::ClickEvent)))}}.GetNewClosure())
    & $login.reload
    if ($SmokeTest) {
        $loginWindow.Add_ContentRendered({
            $selected=@($login.users | Where-Object {$_.userId -eq $login.smokeUserId}) | Select-Object -First 1
            if (-not $selected) {$login.error='화면 검증용 로그인 사용자를 찾을 수 없습니다.';$login.window.Close();return}
            $login.LoginUsersList.SelectedItem=$selected
            if ($login.screenshot) {
                $loginImage=Join-Path ([IO.Path]::GetDirectoryName($login.screenshot)) (([IO.Path]::GetFileNameWithoutExtension($login.screenshot))+'-Login.png')
                & $login.imageWriter $login.window $loginImage
            }
            $login.LoginButton.RaiseEvent((New-Object Windows.RoutedEventArgs([Windows.Controls.Button]::ClickEvent)))
        }.GetNewClosure())
    }
    [void]$loginWindow.ShowDialog()
    if ($login.error) {throw $login.error}
    return $login.context
}
$script:context=if ($AdminMode) {Get-Context -Root $DataRoot -AccessMode Admin} else {Show-LoginDialog}
if (-not $script:context) {return}
$reader = [Xml.XmlReader]::Create((Join-Path $PSScriptRoot 'Views\Main.xaml'))
try { $script:window = [Windows.Markup.XamlReader]::Load($reader) } finally { $reader.Close() }
# Fit the initial window to the desktop work area, including the taskbar.
$desktopArea=[Windows.SystemParameters]::WorkArea
$window.Width=[Math]::Min($window.Width,[Math]::Max($window.MinWidth,$desktopArea.Width))
$window.Height=[Math]::Min($window.Height,[Math]::Max($window.MinHeight,$desktopArea.Height))
$script:ui = @{}
$names = @('IdentityText','PeriodCombo','RefreshButton','ScopeCombo','SearchBox','CaseCountText','ExportButton','ZipButton','CasesGrid','CaseTitle','CaseSubtitle','MainTabs','SubmissionBox','SubmitButton','CancelSubmissionButton','SubmissionResult','DeliverySummaryText','PublishedGrid','DeliveryTab','SourceCombo','LoadDeliveryButton','DeliveryStateText','DocumentsGrid','AddDocumentButton','RemoveDocumentButton','DraftButton','PublishButton','EvaluationStateText','RankGrid','RatingsGrid','SaveEvaluationButton','EvaluationSavedText','AdminTab','AdminContentHost','ModeText','StatusText')
$names += 'CaseExportButton'
$names += @('StatisticsTab','StatisticsHost','CaseHeaderRow','CaseHeaderPanel')
$names += @('EvaluationScroll','PublishedScroll','DeliveryScroll')
$names += @('CaseSidebar','CaseSidebarColumn','CaseGapColumn','CaseSplitter','SelectionDraftButton','SelectionHelpText')
$names += @('CaseMetadataText','ApplicationNumberBox','CpcBox','AddSubmissionRowButton','AddPastedRowsButton','RemoveSubmissionRowButton','SubmissionRowsGrid','SwitchUserButton','PersonalWorkTab','PersonalTabs','TargetSelectionTab','EvaluationTab','EvaluationTabs','ResultNumbersBox','AddResultNumbersButton','CaseStateCombo','CasesEmptyText','GoEvaluationButton','DocumentCountText','CaseListTitleText','CaseListHintText','WorkflowHintText')
foreach ($name in $names) { $script:ui[$name] = $window.FindName($name) }
$script:busy = $false
$script:changing = $false
$script:cases = @()
$script:selectedCase = $null
$script:deliveryEdit = $null
$script:evaluation = $null
$script:loadedSource = $null
$script:loadedPeriod = $null
$script:published = @()
$script:allSources = @()
$script:documents = New-Object 'System.Collections.ObjectModel.ObservableCollection[object]'
$script:ratings = New-Object 'System.Collections.ObjectModel.ObservableCollection[object]'
$script:ranks = New-Object 'System.Collections.ObjectModel.ObservableCollection[object]'
$script:submissionRows = New-Object 'System.Collections.ObjectModel.ObservableCollection[object]'
$script:requests = @{}
$script:editDirty = $false
$script:evalDirty = $false
$script:adminDirty = $false
$script:adminVersion = 0
$script:adminIsNewUser = $false
$script:smokeFailure = ''
$script:lastPeriod = $null
$script:lastScope = $null
$script:lastCase = $null
$script:activeWorkspace=''
$script:workspaceCases=@{Personal='';AI='';Admin=''}
$script:workspaceFilters=@{}
$script:caseDetailLoaded=$false
$script:targetSelection=$null
$script:selectionLoadedKey=''
$script:selectionBaselineStamp='[]'

function Set-Status([string]$Message) { $ui.StatusText.Text = $Message }
function Update-ActionAvailability {
    $hasCase=$null -ne $script:selectedCase
    $writable=$hasCase -and $script:selectedCase.status -ne 'Cancelled' -and $script:loadedPeriod.status -ne 'Closed'
    $own=$writable -and $script:selectedCase.userId -eq $context.User.userId -and 'examiner' -in $context.User.roles
    $canEdit=$writable -and $script:loadedSource -and $ui.SourceCombo.SelectedItem
    $ui.CaseExportButton.IsEnabled=$hasCase -and -not $script:busy
    $ui.CancelSubmissionButton.IsEnabled=$own -and -not $script:busy
    $ui.GoEvaluationButton.IsEnabled=$hasCase -and $script:selectedCase.userId -eq $context.User.userId
    $ui.LoadDeliveryButton.IsEnabled=$hasCase -and $ui.SourceCombo.Items.Count -gt 0
    foreach ($name in @('ResultNumbersBox','AddResultNumbersButton','AddDocumentButton','RemoveDocumentButton','DraftButton','PublishButton')) {$ui[$name].IsEnabled=[bool]$canEdit}
    $ui.DocumentsGrid.IsReadOnly=-not $canEdit
    $ui.SaveEvaluationButton.IsEnabled=$own -and $script:published.Count -gt 0
    $ui.RatingsGrid.IsReadOnly=-not $own;$ui.RankGrid.IsReadOnly=-not $own
    $ui.SubmitButton.IsEnabled='examiner' -in $context.User.roles -and $script:loadedPeriod.status -ne 'Closed'
    $canSelect=$ui.SubmitButton.IsEnabled -and ($AdminMode -or ($script:targetSelection -and $script:targetSelection.status -ne 'Confirmed'))
    foreach ($name in @('SubmitButton','SelectionDraftButton','ApplicationNumberBox','CpcBox','SubmissionBox','AddSubmissionRowButton','AddPastedRowsButton','RemoveSubmissionRowButton')) {$ui[$name].IsEnabled=[bool]$canSelect}
    $ui.SubmissionRowsGrid.IsReadOnly=-not $canSelect
    $ui.DocumentCountText.Text="$($script:documents.Count) / 7건"
    $ui.ScopeCombo.IsEnabled=$false
}
function Show-Error($Problem) {
    $message = [string]$Problem
    if ($Problem.Exception) { $message = $Problem.Exception.Message }
    Set-Status "처리하지 못했습니다. 입력내용은 유지됩니다. $message"
    if ($SmokeTest) {$script:smokeFailure=$message;return}
    [Windows.MessageBox]::Show($window,$message,'처리 확인','OK','Warning') | Out-Null
}
function Invoke-UiAction([scriptblock]$Action) { try { & $Action } catch { Show-Error $_ } }
function Get-Request([string]$Key,$Payload) {
    $json = ConvertTo-Json -InputObject $Payload -Depth 80 -Compress
    if (-not $requests.ContainsKey($Key) -or $requests[$Key].json -cne $json) { $requests[$Key] = @{json=$json;id=[guid]::NewGuid().ToString('N')} }
    return $requests[$Key].id
}
function Get-EvaluationInputStamp {
    $inputState=[ordered]@{
        ratings=@($script:ratings | ForEach-Object {[ordered]@{source=[string]$_.sourceId;document=[string]$_.documentId;score=[string]$_.scoreText;cited=[string]$_.citedText;note=[string]$_.note}})
        ranks=@($script:ranks | ForEach-Object {[ordered]@{source=[string]$_.sourceId;choice=[string]$_.choice}})
    }
    return ConvertTo-Json -InputObject $inputState -Depth 10 -Compress
}
function Update-EvaluationDirty {
    if (-not $script:changing -and -not $script:busy -and $script:evalBaselineStamp) {$script:evalDirty=(Get-EvaluationInputStamp) -cne $script:evalBaselineStamp}
}
function Confirm-Discard {
    if (-not ($script:editDirty -or $script:evalDirty -or $script:adminDirty)) { return $true }
    if ($SmokeTest) {throw "화면 검증 중 예상하지 않은 미저장 변경: 결과=$script:editDirty 평가=$script:evalDirty 운영=$script:adminDirty"}
    return ([Windows.MessageBox]::Show($window,'저장하지 않은 입력을 버리고 이동할까요?','입력내용 확인','YesNo','Question') -eq 'Yes')
}
function Has-PendingSubmissions {
    $changed=if ($AdminMode) {$script:submissionRows.Count -gt 0} else {(Get-SelectionInputStamp) -cne $script:selectionBaselineStamp}
    return $changed -or -not [string]::IsNullOrWhiteSpace($ui.ApplicationNumberBox.Text) -or -not [string]::IsNullOrWhiteSpace($ui.CpcBox.Text) -or -not [string]::IsNullOrWhiteSpace($ui.SubmissionBox.Text)
}
function Get-SelectionInputStamp {
    return ConvertTo-Json -InputObject @($script:submissionRows | ForEach-Object {[ordered]@{applicationNumber=[string]$_.applicationNumber;cpc=[string]$_.cpc}}) -Compress
}
function Set-SelectionStage {
    $previous=$script:changing;$script:changing=$true
    $confirmed=$script:targetSelection -and $script:targetSelection.status -eq 'Confirmed'
    $ui.TargetSelectionTab.Visibility=if ($AdminMode -or -not $confirmed) {'Visible'} else {'Collapsed'}
    $ui.EvaluationTab.Visibility=if ($AdminMode -or $confirmed) {'Visible'} else {'Collapsed'}
    $ui.SelectionDraftButton.Visibility=if ($AdminMode) {'Collapsed'} else {'Visible'}
    $ui.CancelSubmissionButton.Visibility=if ($AdminMode) {'Visible'} else {'Collapsed'}
    $ui.SubmitButton.Content=if ($AdminMode) {'출원 등록'} else {'선정 확정'}
    $ui.SelectionHelpText.Text=if ($AdminMode) {'출원번호와 CPC를 입력한 뒤 출원 등록을 누르세요.'} else {'이번 운영월의 출원번호와 CPC를 입력하세요. 임시저장 후 이어서 수정할 수 있으며, 선정 확정 후 평가 단계로 이동합니다.'}
    foreach ($tabs in @($ui.PersonalTabs,$ui.EvaluationTabs)) {
        if ($AdminMode) {$tabs.ClearValue([Windows.Controls.Control]::TemplateProperty)} else {$tabs.Template=$window.Resources['ContentOnlyTabs']}
        $tabs.Margin=if ($AdminMode) {12} else {0}
    }
    if (-not $AdminMode -and 'ai' -notin $context.User.roles) {$ui.MainTabs.Template=$window.Resources['ContentOnlyTabs']} else {$ui.MainTabs.ClearValue([Windows.Controls.Control]::TemplateProperty)}
    if (-not $AdminMode) {
        $ui.PersonalTabs.SelectedItem=if ($confirmed) {$ui.EvaluationTab} else {$ui.TargetSelectionTab}
        $ui.EvaluationTabs.SelectedIndex=1
    }
    $script:changing=$previous
    Set-WorkspacePresentation
}
function Save-SelectionEditor([string]$Status) {
    Add-SubmissionInputRows
    [void]$ui.SubmissionRowsGrid.CommitEdit([Windows.Controls.DataGridEditingUnit]::Cell,$true)
    [void]$ui.SubmissionRowsGrid.CommitEdit([Windows.Controls.DataGridEditingUnit]::Row,$true)
    $period=Get-SelectedPeriod
    $rows=@($script:submissionRows | ForEach-Object {[pscustomobject]@{applicationNumber=[string]$_.applicationNumber;cpc=[string]$_.cpc}})
    $expected=if ($script:targetSelection) {[int]$script:targetSelection.version} else {0}
    $request=Get-Request "selection:$period" @{expected=$expected;status=$Status;rows=$rows}
    Set-Status '운영월별 대상 건 저장 중…'
    Invoke-UiWork -Work {
        param($ctx,$period,$expected,$request,$status,$rows)
        Save-TargetSelection -Context $ctx -Period $period -ExpectedVersion $expected -RequestId $request -Status $status -Rows $rows
    } -Arguments @($context,$period,$expected,$request,$Status,$rows) -Completed {
        param($result)
        $script:targetSelection=$result[0]
        $script:selectionBaselineStamp=Get-SelectionInputStamp
        $ui.SubmissionResult.Text=if ($script:targetSelection.status -eq 'Confirmed') {'선정 확정됨'} else {'임시저장 완료 · 다음 로그인에서도 이어서 수정할 수 있습니다.'}
        if ($script:targetSelection.status -eq 'Confirmed') {Set-SelectionStage;Load-Cases} else {Set-Status $ui.SubmissionResult.Text}
    }
}
function Save-TestScreenshot([string]$Suffix='') {
    if (-not $ScreenshotPath) {return}
    $path=$ScreenshotPath
    if ($Suffix) {$path=Join-Path ([IO.Path]::GetDirectoryName($ScreenshotPath)) (([IO.Path]::GetFileNameWithoutExtension($ScreenshotPath))+$Suffix+'.png')}
    Write-UiImage $window $path
}
function Invoke-UiWork([scriptblock]$Work,[object[]]$Arguments,[scriptblock]$Completed) {
    if ($script:busy) { throw '현재 작업이 끝난 뒤 다시 시도하세요.' }
    $script:busy = $true
    $ui.MainTabs.IsEnabled = $false; $ui.CasesGrid.IsEnabled = $false
    $ui.PeriodCombo.IsEnabled = $false; $ui.ScopeCombo.IsEnabled = $false
    $ui.RefreshButton.IsEnabled = $false; $ui.ExportButton.IsEnabled = $false; $ui.ZipButton.IsEnabled = $false
    $ui.CaseExportButton.IsEnabled=$false
    $ui.SwitchUserButton.IsEnabled=$false
    $ps = [PowerShell]::Create()
    [void]$ps.AddScript({param($app) foreach ($name in @('Storage','Workflow','Export')) { Import-Module (Join-Path $app "Modules\$name.psm1") -DisableNameChecking }})
    [void]$ps.AddArgument($appDirectory)
    [void]$ps.AddStatement().AddScript($Work.ToString())
    foreach ($argument in $Arguments) { [void]$ps.AddArgument($argument) }
    $handle = $ps.BeginInvoke()
    $timer = New-Object Windows.Threading.DispatcherTimer
    $timer.Interval = [TimeSpan]::FromMilliseconds(100)
    $state = @{ps=$ps;handle=$handle;timer=$timer;completed=$Completed}
    $timer.Tag = $state
    $timer.Add_Tick({
        param($sender,$eventArgs)
        $job = $sender.Tag
        if (-not $job.handle.IsCompleted) { return }
        $sender.Stop()
        try {
            $result = $job.ps.EndInvoke($job.handle)
            if ($job.ps.Streams.Error.Count -gt 0) { throw $job.ps.Streams.Error[0] }
        } catch {
            $job.ps.Dispose(); $sender.Tag=$null
            $script:busy = $false
            $ui.MainTabs.IsEnabled=$true; $ui.CasesGrid.IsEnabled=$true; $ui.PeriodCombo.IsEnabled=$true; $ui.ScopeCombo.IsEnabled=$true
            $ui.RefreshButton.IsEnabled=$true; $ui.ExportButton.IsEnabled=$true; $ui.ZipButton.IsEnabled=$true
            $ui.CaseExportButton.IsEnabled=$true
            $ui.SwitchUserButton.IsEnabled=$true
            Update-ActionAvailability
            Show-Error $_
            return
        }
        $job.ps.Dispose(); $sender.Tag=$null
        $script:busy=$false
        $ui.MainTabs.IsEnabled=$true; $ui.CasesGrid.IsEnabled=$true; $ui.PeriodCombo.IsEnabled=$true; $ui.ScopeCombo.IsEnabled=$true
        $ui.RefreshButton.IsEnabled=$true; $ui.ExportButton.IsEnabled=$true; $ui.ZipButton.IsEnabled=$true
        $ui.CaseExportButton.IsEnabled=$true
        $ui.SwitchUserButton.IsEnabled=$true
        try { & $job.completed $result } catch { Show-Error $_ }
        if (-not $script:busy) {Update-ActionAvailability}
    })
    $timer.Start()
}
function Get-SelectedPeriod { if ($ui.PeriodCombo.SelectedItem) { return [string]$ui.PeriodCombo.SelectedItem.period }; throw '운영월을 선택하세요.' }
function Require-Case { if (-not $script:selectedCase) { throw '왼쪽 목록에서 출원을 선택하세요.' }; return $script:selectedCase }
function Get-UiValue($Object,[string]$Name,$Default='') {
    if ($null -eq $Object) {return $Default}
    if ($Object -is [Collections.IDictionary]) {if ($Object.Contains($Name) -and $null -ne $Object[$Name]) {return $Object[$Name]};return $Default}
    $property=$Object.PSObject.Properties[$Name]
    if ($null -eq $property -or $null -eq $property.Value) {return $Default}
    return $property.Value
}
function Find-VisualControls($Parent,[type]$Type) {
    if ($null -eq $Parent) {return}
    if ($Parent -is $Type) {$Parent}
    for ($index=0;$index -lt [Windows.Media.VisualTreeHelper]::GetChildrenCount($Parent);$index++) {Find-VisualControls ([Windows.Media.VisualTreeHelper]::GetChild($Parent,$index)) $Type}
}
function Enable-GridSingleClick($Grid) {
    $Grid.Add_PreviewMouseDown({
        param($sender,$eventArgs)
        if ($sender.IsReadOnly -or $eventArgs.ChangedButton -ne [Windows.Input.MouseButton]::Left) {return}
        $node=$eventArgs.OriginalSource
        while ($node -and $node -isnot [Windows.Controls.DataGridCell]) {
            if ($node -is [Windows.Controls.ComboBox] -or $node -is [Windows.Controls.Primitives.ButtonBase] -or $node -is [Windows.Controls.TextBox]) {return}
            if ($node -is [Windows.Media.Visual] -or $node -is [Windows.Media.Media3D.Visual3D]) {$node=[Windows.Media.VisualTreeHelper]::GetParent($node)}
            elseif ($node -is [Windows.FrameworkContentElement]) {$node=$node.Parent}
            else {return}
        }
        if ($node -and -not $node.IsReadOnly -and -not $node.IsEditing) {
            $sender.SelectedItem=$node.DataContext
            $sender.CurrentCell=New-Object Windows.Controls.DataGridCellInfo($node.DataContext,$node.Column)
            [void]$node.Focus();[void]$sender.BeginEdit()
        }
    })
}
function Set-Identity {
    $affiliation=@((Get-UiValue $context.User 'bureau'),(Get-UiValue $context.User 'division')) | Where-Object {$_}
    $modeLabel=if ($AdminMode) {'관리자'} else {'일반 사용자'}
    $ui.IdentityText.Text=(@($modeLabel,$context.User.name,($affiliation -join ' / ')) | Where-Object {$_}) -join ' · '
}
function Get-CurrentWorkspace {
    if ($ui.MainTabs.SelectedItem -eq $ui.StatisticsTab) {return 'Statistics'}
    if ($ui.MainTabs.SelectedItem -eq $ui.DeliveryTab) {return 'AI'}
    if ($ui.MainTabs.SelectedItem -eq $ui.AdminTab) {return 'Admin'}
    return 'Personal'
}
function Get-WorkspaceScope {
    if ($script:activeWorkspace -eq 'Personal') {return 'Mine'}
    if ($script:activeWorkspace -eq 'AI' -and 'admin' -notin $context.User.roles) {return 'Assigned'}
    return 'All'
}
function Set-WorkspacePresentation {
    $scope=Get-WorkspaceScope
    $name=switch ($script:activeWorkspace) {'Personal' {'본인이 선정한 출원'};'AI' {if ($scope -eq 'All') {'전체 출원 · 관리자 대리 입력'} else {'AI 배정 그룹의 출원'}};'Admin' {'전체 출원 · 운영 관리'}}
    $previous=$script:changing;$script:changing=$true
    $ui.ScopeCombo.ItemsSource=@([pscustomobject]@{id=$scope;name=$name});$ui.ScopeCombo.SelectedIndex=0;$ui.ScopeCombo.IsEnabled=$false
    $script:changing=$previous
    $statistics=$script:activeWorkspace -eq 'Statistics'
    $ui.CaseHeaderPanel.Visibility=if ($statistics) {'Collapsed'} else {'Visible'}
    $ui.CaseHeaderRow.Height=if ($statistics) {0} else {78}
    $selectionOnly=$statistics -or (-not $AdminMode -and $script:activeWorkspace -eq 'Personal' -and (-not $script:targetSelection -or $script:targetSelection.status -ne 'Confirmed'))
    $ui.CaseSidebar.Visibility=if ($selectionOnly) {'Collapsed'} else {'Visible'}
    $ui.CaseSplitter.Visibility=$ui.CaseSidebar.Visibility
    $showExaminer=$script:activeWorkspace -in @('AI','Admin')
    $ui.CaseSidebarColumn.MinWidth=if ($selectionOnly) {0} elseif ($showExaminer) {375} else {290}
    $ui.CaseSidebarColumn.Width=if ($selectionOnly) {0} elseif ($showExaminer) {390} else {340}
    $ui.CaseGapColumn.Width=if ($selectionOnly) {0} else {14}
    $ui.CasesGrid.Columns[1].Visibility=if ($script:activeWorkspace -eq 'Personal') {'Collapsed'} else {'Visible'}
    switch ($script:activeWorkspace) {
        'Personal' {
            $selection=$ui.PersonalTabs.SelectedItem -eq $ui.TargetSelectionTab
            $ui.CaseListTitleText.Text=if ($selection) {'내 선정 대상 건'} else {'내 평가 대상 건'}
            $ui.CaseListHintText.Text=if ($selection) {'본인 출원을 등록·선정하는 단계입니다.'} else {'본인이 선정한 출원의 결과를 평가합니다.'}
            $ui.WorkflowHintText.Text=if ($selection) {'1. 대상 건 선정 · 임시저장한 목록을 검토한 뒤 대상 건을 확정하세요.'} else {'2. AI 검색 결과 평가 · 왼쪽 출원번호를 클릭하면 바로 평가할 수 있습니다. 입력 완료는 이번 달 AI 검색 결과가 모두 입력된 상태입니다.'}
            if ($selectionOnly) {$ui.CaseTitle.Text=(Get-UiValue $script:loadedPeriod 'period' '')+' 대상 건 선정';$ui.CaseMetadataText.Text='';$ui.CaseSubtitle.Text='출원번호 · CPC 입력 → 임시저장 → 선정 확정'}
            $ui.CasesEmptyText.Text=if ($selection) {"선정한 출원이 없습니다.`n출원번호와 CPC를 등록하세요."} else {"평가할 본인 출원이 없습니다.`n1. 대상 건 선정에서 출원을 등록하세요."}
        }
        'AI' {
            $ui.CaseListTitleText.Text='AI 검색 대상 건'
            $ui.CaseListHintText.Text='담당 그룹의 AI 검색 결과를 입력합니다.'
            $ui.WorkflowHintText.Text='AI 담당 업무 · 본인의 대상 건 선정·평가는 내 심사 업무에서 진행하세요.'
            $ui.CasesEmptyText.Text="담당 대상 건이 없습니다.`n선택한 월의 그룹·AI 배정을 확인하세요."
        }
        'Admin' {
            $ui.CaseListTitleText.Text='전체 운영 현황'
            $ui.CaseListHintText.Text='사용자·그룹·AI 배정과 월 운영을 관리합니다.'
            $ui.WorkflowHintText.Text='운영 관리 · 전체 출원과 월별 운영 설정을 확인합니다.'
            $ui.CasesEmptyText.Text='선택한 운영월에 등록된 출원이 없습니다.'
        }
    }
}
function Change-Workspace {
    $next=Get-CurrentWorkspace
    if ($next -eq $script:activeWorkspace) {Set-WorkspacePresentation;return}
    if (-not (Confirm-Discard)) {
        $script:changing=$true;$ui.MainTabs.SelectedItem=$script:lastMainTab;$script:changing=$false;return
    }
    if ($script:adminDirty) {Select-AdminRecord}
    if ($script:activeWorkspace) {
        $script:workspaceCases[$script:activeWorkspace]=if ($script:selectedCase) {$script:selectedCase.caseId} else {''}
        $script:workspaceFilters[$script:activeWorkspace]=@{search=$ui.SearchBox.Text;state=$ui.CaseStateCombo.SelectedIndex}
    }
    $script:activeWorkspace=$next;$script:lastMainTab=$ui.MainTabs.SelectedItem
    $script:changing=$true
    Clear-Detail
    $filter=$script:workspaceFilters[$next]
    $ui.SearchBox.Text=if ($filter) {$filter.search} else {''}
    $ui.CaseStateCombo.SelectedIndex=if ($filter) {[int]$filter.state} else {0}
    $script:changing=$false
    Set-WorkspacePresentation
    Load-Cases -SelectCaseId $script:workspaceCases[$next]
}
function Select-Workspace([ValidateSet('Selection','Evaluation','AI','Admin','Statistics')][string]$Name) {
    $tab=switch ($Name) {'Selection' {$ui.PersonalWorkTab};'Evaluation' {$ui.PersonalWorkTab};'AI' {$ui.DeliveryTab};'Admin' {$ui.AdminTab};'Statistics' {$ui.StatisticsTab}}
    if ($tab.Visibility -ne 'Visible') {throw '현재 역할에서 사용할 수 없는 업무입니다.'}
    if ($Name -in @('Selection','Evaluation')) {
        $step=if ($Name -eq 'Selection') {$ui.TargetSelectionTab} else {$ui.EvaluationTab}
        if ($step.Visibility -ne 'Visible') {$step=if ($script:targetSelection -and $script:targetSelection.status -eq 'Confirmed') {$ui.EvaluationTab} else {$ui.TargetSelectionTab}}
        if ($ui.MainTabs.SelectedItem -eq $ui.PersonalWorkTab) {$ui.PersonalTabs.SelectedItem=$step}
        else {$previous=$script:changing;$script:changing=$true;$ui.PersonalTabs.SelectedItem=$step;$script:changing=$previous}
        if ($Name -eq 'Evaluation') {$ui.EvaluationTabs.SelectedIndex=1}
    }
    $ui.MainTabs.SelectedItem=$tab
}
function Add-SubmissionInputRows([bool]$IncludeSingle=$true,[bool]$IncludePaste=$true) {
    $pending=@()
    if ($IncludeSingle -and ($ui.ApplicationNumberBox.Text.Trim() -or $ui.CpcBox.Text.Trim())) {
        if (-not $ui.ApplicationNumberBox.Text.Trim()) {throw 'CPC에 대응하는 출원번호를 입력하세요.'}
        $pending += [pscustomobject]@{applicationNumber=$ui.ApplicationNumberBox.Text.Trim();cpc=$ui.CpcBox.Text.Trim()}
    }
    if ($IncludePaste) {
        foreach ($line in @($ui.SubmissionBox.Text -split '[\r\n]+')) {
            if ([string]::IsNullOrWhiteSpace($line)) {continue}
            $columns=@($line -split "`t")
            if ($columns.Count -gt 2) {throw '붙여넣기는 출원번호와 CPC 두 열만 선택하세요.'}
            $number=$columns[0].Trim()
            if (-not $number) {throw '붙여넣기 행의 첫 열에 출원번호가 필요합니다.'}
            $cpc='';if ($columns.Count -eq 2) {$cpc=$columns[1].Trim()}
            $pending += [pscustomobject]@{applicationNumber=$number;cpc=$cpc}
        }
    }
    foreach ($row in $pending) {$script:submissionRows.Add([pscustomobject]@{rowKey=[guid]::NewGuid().ToString('N');applicationNumber=$row.applicationNumber;cpc=$row.cpc;message=''})}
    if ($IncludeSingle) {$ui.ApplicationNumberBox.Clear();$ui.CpcBox.Clear()}
    if ($IncludePaste) {$ui.SubmissionBox.Clear()}
}

function Update-Filter {
    $search = $ui.SearchBox.Text.Trim()
    $state=if ($ui.CaseStateCombo.SelectedItem) {[string]$ui.CaseStateCombo.SelectedItem.Tag} else {'Active'}
    $filtered = @($script:cases | Where-Object {
        $text=@($_.applicationNumber,$_.examinerName,(Get-UiValue $_ 'cpc'),(Get-UiValue $_ 'examinerBureau'),(Get-UiValue $_ 'examinerDivision')) -join "`n"
        $matches=-not $search -or $text.IndexOf($search,[StringComparison]::OrdinalIgnoreCase) -ge 0
        $matches -and ($state -eq 'All' -or ($state -eq 'Active' -and $_.status -ne 'Cancelled') -or ($state -eq 'Cancelled' -and $_.status -eq 'Cancelled') -or ($state -eq 'Duplicate' -and $_.duplicateWarning))
    })
    $previous=$script:changing;$script:changing=$true
    $ui.CasesGrid.ItemsSource=$filtered
    if ($script:selectedCase) {$ui.CasesGrid.SelectedItem=@($filtered | Where-Object {$_.caseId -eq $script:selectedCase.caseId}) | Select-Object -First 1}
    $script:changing=$previous
    $ui.CaseCountText.Text="표시 $($filtered.Count) / 전체 $($script:cases.Count)건"
    $ui.CasesEmptyText.Visibility=if ($filtered.Count) {'Collapsed'} else {'Visible'}
}
function Clear-Detail {
    $script:caseDetailLoaded=$false
    $script:evalBaselineStamp=$null
    $script:selectedCase=$null; $script:deliveryEdit=$null; $script:evaluation=$null; $script:loadedSource=$null
    $script:published=@(); $script:editDirty=$false; $script:evalDirty=$false
    $script:documents.Clear(); $script:ratings.Clear(); $script:ranks.Clear()
    $ui.ResultNumbersBox.Clear()
    $ui.PublishedGrid.ItemsSource=@(); $ui.CaseTitle.Text='대상 건을 등록하거나 목록에서 선택하세요'
    $ui.CaseSubtitle.Text='선택한 업무의 출원 목록에서 대상 건을 선택하세요.'
    $ui.CaseMetadataText.Text=''
    $ui.DeliveryStateText.Text=''; $ui.EvaluationSavedText.Text=''
    $ui.DeliverySummaryText.Text='왼쪽에서 출원을 선택하면 검색 결과가 표시됩니다.'
    $ui.EvaluationStateText.Text='왼쪽에서 본인 출원을 선택하면 바로 평가할 수 있습니다.'
    Update-ActionAvailability
}
function Load-Cases([string]$SelectCaseId='') {
    if ($script:activeWorkspace -eq 'Statistics') {Load-Statistics;return}
    if (-not $SelectCaseId -and $script:selectedCase) {$SelectCaseId=$script:selectedCase.caseId}
    $script:refreshCaseId=$SelectCaseId
    $period = Get-SelectedPeriod
    $scope=Get-WorkspaceScope
    $script:lastPeriod=$ui.PeriodCombo.SelectedItem
    $script:lastScope=$ui.ScopeCombo.SelectedItem
    $script:lastCase=$null
    $script:changing=$true
    Clear-Detail
    $script:cases=@();$ui.CasesGrid.ItemsSource=@();$ui.CaseCountText.Text='자료를 불러오는 중…';$ui.CasesEmptyText.Visibility='Collapsed'
    $script:changing=$false
    Set-Status '출원 및 월 운영정보 조회 중…'
    Invoke-UiWork -Work {
        param($ctx,$month,$scope)
        $cases=@(Get-Cases -Context $ctx -Period $month -Scope $scope)
        if ($scope -eq 'Mine') {foreach ($case in $cases) {
            $progress=Get-CaseStatus -Context $ctx -Case $case
            $case | Add-Member -NotePropertyName deliveryStatusLabel -NotePropertyValue $progress.deliveryStatusLabel
            $case | Add-Member -NotePropertyName deliveryHint -NotePropertyValue ("AI 결과 입력 $($progress.aiCompleted) / $($progress.aiTotal)")
        }}
        $selection=$null;if ('examiner' -in $ctx.User.roles) {$selection=Get-TargetSelection -Context $ctx -Period $month}
        [pscustomobject]@{cases=$cases;selection=$selection;period=(Get-Period -Context $ctx -Period $month);sources=@(Get-Sources -Context $ctx);groups=@(Get-Groups -Context $ctx)}
    } -Arguments @($context,$period,$scope) -Completed {
        param($result)
        $snapshot=$result[0]; $script:loadedPeriod=$snapshot.period
        $script:cases=@($snapshot.cases)
        foreach ($case in $script:cases) {
            $case | Add-Member -NotePropertyName statusLabel -NotePropertyValue $(if ($case.status -eq 'Cancelled') {'취소'} elseif ($script:activeWorkspace -eq 'Personal') {$case.deliveryStatusLabel} elseif ($case.duplicateWarning) {'중복 확인'} else {'선정 확정'}) -Force
            $case | Add-Member -NotePropertyName caseHint -NotePropertyValue ("CPC: "+(Get-UiValue $case 'cpc')+"`n"+(Get-UiValue $case 'examinerBureau')+' / '+(Get-UiValue $case 'examinerDivision')) -Force
        }
        $script:allSources=@($snapshot.sources)
        $script:allGroups=@($snapshot.groups)
        $script:changing=$true
        if (-not $AdminMode) {
            $key=$context.User.userId+':'+$script:loadedPeriod.period
            if ($key -ne $script:selectionLoadedKey -or (-not (Has-PendingSubmissions) -and $snapshot.selection.version -ne $script:targetSelection.version)) {
                $script:targetSelection=$snapshot.selection;$script:selectionLoadedKey=$key
                $script:submissionRows.Clear();$ui.ApplicationNumberBox.Clear();$ui.CpcBox.Clear();$ui.SubmissionBox.Clear()
                if ($script:targetSelection -and $script:targetSelection.status -eq 'Draft') {foreach ($row in $script:targetSelection.rows) {
                    $script:submissionRows.Add([pscustomobject]@{rowKey=[guid]::NewGuid().ToString('N');applicationNumber=$row.applicationNumber;cpc=$row.cpc;message=''})
                }}
                $script:selectionBaselineStamp=Get-SelectionInputStamp
                $ui.SubmissionResult.Text=if ($script:targetSelection -and $script:targetSelection.status -eq 'Draft') {'저장한 대상 건 초안을 불러왔습니다.'} else {''}
            } elseif (-not (Has-PendingSubmissions)) {$script:targetSelection=$snapshot.selection}
        }
        Set-SelectionStage
        $ui.SourceCombo.ItemsSource=@()
        Update-Filter
        $script:changing=$false
        Set-WorkspacePresentation
        Set-Status '조회 완료 · 목록에서 출원을 선택하세요.'
        $next=@($ui.CasesGrid.ItemsSource | Where-Object {$_.caseId -eq $script:refreshCaseId}) | Select-Object -First 1
        if (-not $next) {$next=@($ui.CasesGrid.ItemsSource) | Select-Object -First 1}
        if ($next) {$ui.CasesGrid.SelectedItem=$next}
    }
}
function Show-CaseDetail($Case) {
    if (-not $script:selectedCase -or $script:selectedCase.caseId -ne $Case.caseId) {
        foreach ($scrollName in @('EvaluationScroll','PublishedScroll','DeliveryScroll')) {$ui[$scrollName].ScrollToTop()}
    }
    Set-Status '검색 결과와 평가 자료를 불러오는 중…'
    $script:evalBaselineStamp=$null
    $script:caseDetailLoaded=$false
    $script:preferredSource=if ($ui.SourceCombo.SelectedItem) {[string]$ui.SourceCombo.SelectedItem.sourceId} else {''}
    $script:changing=$true
    $script:deliveryEdit=$null;$script:loadedSource=$null;$script:evaluation=$null;$script:published=@()
    $script:documents.Clear();$script:ratings.Clear();$script:ranks.Clear();$ui.ResultNumbersBox.Clear();$ui.PublishedGrid.ItemsSource=@()
    $ui.DeliveryStateText.Text='조회 중…'
    $script:editDirty=$false;$script:evalDirty=$false;$script:changing=$false
    $script:selectedCase=$Case
    $script:lastCase=$Case
    $ui.CaseTitle.Text=$Case.applicationNumber
    $cpc=[string](Get-UiValue $Case 'cpc');if (-not $cpc) {$cpc='미입력'}
    $affiliation=@((Get-UiValue $Case 'examinerBureau'),(Get-UiValue $Case 'examinerDivision')) | Where-Object {$_}
    $ui.CaseMetadataText.Text=(@("CPC: $cpc",($affiliation -join ' / ')) | Where-Object {$_}) -join ' · '
    $ui.CaseMetadataText.ToolTip=$ui.CaseMetadataText.Text
    $groupName=@($script:allGroups | Where-Object {$_.groupId -eq $Case.groupId}) | Select-Object -First 1
    $script:caseGroupName=if ($groupName) {$groupName.name} else {'미지정 그룹'}
    $ui.CaseSubtitle.Text="$($Case.examinerName) · $script:caseGroupName"
    if ($script:activeWorkspace -eq 'Personal' -and $ui.PersonalTabs.SelectedItem -eq $ui.TargetSelectionTab) {
        Set-Status '대상 건 조회 완료 · 결과 조회와 평가는 2. AI 검색 결과 평가에서 진행하세요.'
        Update-ActionAvailability
        return
    }
    Invoke-UiWork -Work {
        param($ctx,$case)
        $evaluation=$null
        if ($case.userId -eq $ctx.User.userId -or 'Admin' -in $ctx.User.roles) { $evaluation=Get-Evaluation -Context $ctx -Case $case }
        [pscustomobject]@{deliveries=@(Get-Deliveries -Context $ctx -Case $case);evaluation=$evaluation;progress=(Get-CaseStatus -Context $ctx -Case $case)}
    } -Arguments @($context,$Case) -Completed {
        param($result)
        $snapshot=$result[0]
        $script:changing=$true
        if ($script:activeWorkspace -eq 'Personal') {
            $script:selectedCase.statusLabel=if ($script:selectedCase.status -eq 'Cancelled') {'취소'} else {$snapshot.progress.deliveryStatusLabel}
            $ui.CasesGrid.Items.Refresh()
        }
        $script:deliveryEdit=$null; $script:loadedSource=$null; $script:documents.Clear(); $ui.DeliveryStateText.Text='편집 자료를 불러오세요'
        $script:published=@($snapshot.deliveries); $script:evaluation=$snapshot.evaluation
        $editable=@($script:allSources | Where-Object {
            $source=$_
            $source.sourceId -in $script:loadedPeriod.sourceIds -and (
                'admin' -in $context.User.roles -or
                ($script:activeWorkspace -eq 'AI' -and $source.type -eq 'AI' -and 'ai' -in $context.User.roles -and @($script:loadedPeriod.assignments | Where-Object {$_.userId -eq $context.User.userId -and $_.groupId -eq $script:selectedCase.groupId -and $_.sourceId -eq $source.sourceId}).Count -gt 0)
            )
        })
        $ui.SourceCombo.ItemsSource=$editable
        $ui.SourceCombo.SelectedItem=@($editable | Where-Object {$_.sourceId -eq $script:preferredSource}) | Select-Object -First 1
        if (-not $ui.SourceCombo.SelectedItem) {$ui.SourceCombo.SelectedIndex=0}
        $rows=@()
        foreach ($delivery in $script:published) {
            if ($delivery.data.noResults) { $rows += [pscustomobject]@{sourceName=$delivery.data.sourceName;version=$delivery.version;type='결과 없음';identifier=('결과 없음 · '+$delivery.data.noResultsReason);title=$delivery.data.noResultsReason;description=''} }
            foreach ($doc in @($delivery.data.documents)) { $rows += [pscustomobject]@{sourceName=$delivery.data.sourceName;version=$delivery.version;type=$doc.type;identifier=$doc.identifier;title=$doc.title;description=$doc.description} }
        }
        $ui.PublishedGrid.ItemsSource=$rows
        $ui.DeliverySummaryText.Text=if ($script:published.Count) {"조사 방식 $($script:published.Count)개 · 문헌 $(@($rows | Where-Object {$_.type -ne '결과 없음'}).Count)건 · 등록한 결과는 바로 반영됩니다."} else {'등록된 검색 결과가 없습니다. 담당자가 결과를 등록하면 여기에 표시됩니다.'}
        $states=@{NotStarted='미착수';Draft='작성 중';Published='입력 완료';NoResults='결과 없음';PublishedWithDraft='입력 완료 후 수정 중'}
        $sourceStatus=@($snapshot.progress.deliveries | ForEach-Object {$s=$_;$label=@($script:allSources | Where-Object {$_.sourceId -eq $s.sourceId}) | Select-Object -First 1;"$($label.name): $($states[$s.status]) $($s.documentCount)건"}) -join ' / '
        $ui.CaseSubtitle.Text="$($script:selectedCase.examinerName) · $script:caseGroupName · $sourceStatus"
        $ui.CaseSubtitle.ToolTip=$ui.CaseSubtitle.Text
        Initialize-EvaluationRows
        $script:caseDetailLoaded=$true
        if ($SmokeTest -and $script:expectedRefreshSource -and @($script:published | Where-Object {$_.data.sourceId -eq $script:expectedRefreshSource -and $_.version -eq $script:expectedRefreshVersion}).Count) {$script:automaticDeliveryRefreshVerified=$true}
        $script:editDirty=$false; $script:evalDirty=$false; $script:changing=$false
        Set-Status '조회 완료 · 내용을 확인하고 평가를 저장하세요.'
        if ($ui.MainTabs.SelectedItem -eq $ui.DeliveryTab -and $ui.SourceCombo.SelectedItem) {Load-DeliveryEditor}
    }
}
function Initialize-EvaluationRows {
    $script:ratings.Clear(); $script:ranks.Clear()
    $own=$script:selectedCase.userId -eq $context.User.userId -and $script:selectedCase.status -ne 'Cancelled' -and $script:loadedPeriod.status -ne 'Closed'
    $ui.SaveEvaluationButton.IsEnabled=$own
    $ui.RatingsGrid.IsReadOnly=-not $own; $ui.RankGrid.IsReadOnly=-not $own
    $stale=$false
    $old=$null; if ($script:evaluation) { $old=$script:evaluation.data }
    foreach ($delivery in $script:published) {
        if ($old) { $prop=$old.deliveryVersions.PSObject.Properties[$delivery.data.sourceId]; if (-not $prop -or [int]$prop.Value -ne $delivery.version) { $stale=$true } }
        $prior=@(); if ($old) { $prior=@($old.rankings | Where-Object {$_.sourceId -eq $delivery.data.sourceId}) }
        $choice=''; if ($delivery.data.noResults) {$choice='평가 제외'} elseif ($prior.Count) {
            switch ($prior[0].status) { 'Ranked' {$choice=[string]$prior[0].rank}; 'Outside' {$choice='순위 외'}; 'Excluded' {$choice='평가 제외'}; 'Incomparable' {$choice='비교 불가'} }
        }
        $script:ranks.Add([pscustomobject]@{sourceId=$delivery.data.sourceId;sourceName=$delivery.data.sourceName;deliveryVersion=$delivery.version;choice=$choice;noResults=[bool]$delivery.data.noResults;notDelivered=$false})
        foreach ($doc in @($delivery.data.documents)) {
            $priorRating=@(); if ($old) { $priorRating=@($old.documentRatings | Where-Object {$_.sourceId -eq $delivery.data.sourceId -and $_.documentId -eq $doc.documentId}) }
            $score='';$cited='';$note=''
            if ($priorRating.Count) { $rating=$priorRating[0];if ($null -ne $rating.score) {$score=[string]$rating.score};if ($null -ne $rating.cited) {$cited=if ($rating.cited) {'O'} else {'X'}};$note=$rating.note }
            $script:ratings.Add([pscustomobject]@{sourceId=$delivery.data.sourceId;sourceName=$delivery.data.sourceName;documentId=$doc.documentId;identifier=$doc.identifier;title=$doc.title;scoreText=$score;citedText=$cited;note=$note})
        }
    }
    # Show every comparison method, including the general prior-art search before it is delivered.
    foreach ($source in @($script:allSources | Where-Object {$_.sourceId -in $script:loadedPeriod.sourceIds -or $_.type -eq 'General'})) {
        if ($source.sourceId -notin @($script:ranks | ForEach-Object {$_.sourceId})) {
            $script:ranks.Add([pscustomobject]@{sourceId=$source.sourceId;sourceName=$source.name;deliveryVersion=0;choice='대기 중';noResults=$false;notDelivered=$true})
        }
    }
    $ordered=@($script:ranks | Sort-Object @{Expression={if ($_.sourceId -in @($script:allSources | Where-Object {$_.type -eq 'General'} | ForEach-Object {$_.sourceId})) {0} else {1}}},sourceId)
    $script:ranks.Clear();foreach ($rank in $ordered) {$script:ranks.Add($rank)}
    $ui.EvaluationSavedText.Text=if ($script:evaluation) {'저장된 평가'} else {'평가 저장 전'}
    $ui.EvaluationStateText.Text=if ($stale) {'검색 결과가 수정되었습니다. 문헌 평가와 전체 순위를 확인한 뒤 저장하세요.'} elseif ($own) {'선행기술조사와 AI 검색 결과를 비교해 순위를 선택하세요. 대기 중인 방식은 검색 결과가 입력되면 평가할 수 있습니다.'} else {'본인 출원의 평가만 입력할 수 있습니다.'}
    $script:evalBaselineStamp=Get-EvaluationInputStamp
}
function Load-DeliveryEditor {
    if ($script:activeWorkspace -ne 'AI' -or $ui.DeliveryTab.Visibility -ne 'Visible') {throw 'AI 담당 업무에서 배정된 검색 결과를 입력하세요.'}
    $case=Require-Case
    if (-not $ui.SourceCombo.SelectedItem) {throw '조사 방식을 선택하세요.'}
    $source=[string]$ui.SourceCombo.SelectedItem.sourceId
    Set-Status '저장된 검색 결과 조회 중…'
    Invoke-UiWork -Work {param($ctx,$case,$source) Get-DeliveryForEdit -Context $ctx -Case $case -SourceId $source} -Arguments @($context,$case,$source) -Completed {
        param($result)
        $script:changing=$true
        $script:deliveryEdit=if ($result.Count) {$result[0]} else {$null}
        $script:loadedSource=[string]$ui.SourceCombo.SelectedItem.sourceId
        $script:documents.Clear()
        $ui.ResultNumbersBox.Clear()
        if ($script:deliveryEdit) {
            foreach ($doc in @($script:deliveryEdit.data.documents)) {
                $script:documents.Add([pscustomobject]@{documentId=$doc.documentId;type=(Get-UiValue $doc 'type' '특허');identifier=$doc.identifier;title=(Get-UiValue $doc 'title');description=(Get-UiValue $doc 'description');attachments=@($doc.attachments)})
            }
        }
        $ui.DeliveryStateText.Text=if ($script:deliveryEdit) {$(if ($script:deliveryEdit.data.status -eq 'Published') {'입력 완료'} else {'임시저장'})} else {'신규 입력 · 저장된 결과 없음'}
        $script:editDirty=$false;$script:changing=$false
        Set-Status '편집 자료를 불러왔습니다.'
    }
}
function Add-ResultNumberRows {
    if ([string]::IsNullOrWhiteSpace($ui.ResultNumbersBox.Text)) {return}
    if (-not $script:loadedSource) {throw '왼쪽에서 대상 출원을 선택하고 조사 방식을 불러오세요.'}
    $numbers=@($ui.ResultNumbersBox.Text -split '[\r\n,;\t]+' | ForEach-Object {$_.Trim()} | Where-Object {$_} | Select-Object -Unique)
    $existing=@($script:documents | ForEach-Object {$_.identifier.Trim()})
    $numbers=@($numbers | Where-Object {$_ -notin $existing})
    if ($script:documents.Count+$numbers.Count -gt 7) {throw '조사 방식별 출원번호는 최대 7건입니다. 입력내용은 유지됩니다.'}
    foreach ($number in $numbers) {$script:documents.Add([pscustomobject]@{documentId=[guid]::NewGuid().ToString('N');type='특허';identifier=$number;title='';description='';attachments=@()})}
    $ui.ResultNumbersBox.Clear();$script:editDirty=$true
}
function Save-DeliveryEditor([string]$Status) {
    if ($script:activeWorkspace -ne 'AI' -or $ui.DeliveryTab.Visibility -ne 'Visible') {throw 'AI 담당 업무에서 배정된 검색 결과를 입력하세요.'}
    $case=Require-Case
    if ($script:evalDirty -and -not (Confirm-Discard)) {return}
    if (-not $script:loadedSource -or $script:loadedSource -ne $ui.SourceCombo.SelectedItem.sourceId) {throw '조사 방식의 편집 자료를 먼저 불러오세요.'}
    Add-ResultNumberRows
    [void]$ui.DocumentsGrid.CommitEdit([Windows.Controls.DataGridEditingUnit]::Cell,$true);[void]$ui.DocumentsGrid.CommitEdit([Windows.Controls.DataGridEditingUnit]::Row,$true)
    $docs=@($script:documents | ForEach-Object {[pscustomobject]@{documentId=$_.documentId;type=$_.type;identifier=$_.identifier;title=$_.title;description=$_.description;attachments=@($_.attachments)}})
    $version=if ($script:deliveryEdit) {[int]$script:deliveryEdit.version} else {0}
    $payload=@{case=$case.caseId;source=$script:loadedSource;expected=$version;status=$Status;documents=$docs}
    $request=Get-Request "delivery:$($case.caseId):$script:loadedSource" $payload
    Set-Status '검색 결과 저장 중…'
    Invoke-UiWork -Work {
        param($ctx,$case,$payload,$request)
        Save-Delivery -Context $ctx -Case $case -SourceId $payload.source -ExpectedVersion $payload.expected -RequestId $request -Status $payload.status -Documents $payload.documents
    } -Arguments @($context,$case,$payload,$request) -Completed {
        param($result)
        $script:deliveryEdit=$result[0];$script:editDirty=$false
        $ui.DeliveryStateText.Text=$(if ($script:deliveryEdit.data.status -eq 'Published') {'입력 완료'} else {'임시저장'})
        if ($script:deliveryEdit.data.status -eq 'Published' -and $SmokeTest) {$script:expectedRefreshSource=$script:loadedSource;$script:expectedRefreshVersion=$script:deliveryEdit.version}
        Set-Status '검색 결과 저장 완료 · 조회와 평가를 갱신합니다.'
        Show-CaseDetail $script:selectedCase
    }
}
function Save-EvaluationEditor {
    if ($script:activeWorkspace -ne 'Personal' -or $ui.PersonalTabs.SelectedItem -ne $ui.EvaluationTab) {throw '내 심사 업무의 2. AI 검색 결과 평가에서 저장하세요.'}
    $case=Require-Case
    foreach ($grid in @($ui.RatingsGrid,$ui.RankGrid)) {[void]$grid.CommitEdit([Windows.Controls.DataGridEditingUnit]::Cell,$true);[void]$grid.CommitEdit([Windows.Controls.DataGridEditingUnit]::Row,$true)}
    $ratingRows=@(foreach ($row in $script:ratings) {
        $score=$null;$cited=$null
        if (-not [string]::IsNullOrWhiteSpace($row.scoreText)) {if ($row.scoreText -notmatch '^[0-5]$') {throw '점수는 0~5의 정수 또는 공란이어야 합니다.'};$score=[int]$row.scoreText}
        if (-not [string]::IsNullOrWhiteSpace($row.citedText)) {if ($row.citedText.ToUpper() -notin @('O','X')) {throw '인용 여부는 O, X 또는 공란이어야 합니다.'};$cited=$row.citedText.ToUpper() -eq 'O'}
        [pscustomobject]@{sourceId=$row.sourceId;documentId=$row.documentId;score=$score;cited=$cited;note=$row.note}
    })
    $rankingRows=@(foreach ($row in $script:ranks) {
        if ($row.notDelivered) {continue}
        if ([string]::IsNullOrWhiteSpace($row.choice)) {continue}
        $rank=$null
        $state=switch ($row.choice) {'1' {$rank=1;'Ranked'};'2' {$rank=2;'Ranked'};'3' {$rank=3;'Ranked'};'순위 외' {'Outside'};'평가 제외' {'Excluded'};'비교 불가' {'Incomparable'};'' {'Excluded'};default {throw '순위는 1, 2, 3, 순위 외, 평가 제외, 비교 불가 중 입력하세요.'}}
        [pscustomobject]@{sourceId=$row.sourceId;rank=$rank;status=$state}
    })
    $versions=@{};foreach ($delivery in $script:published) {$versions[$delivery.data.sourceId]=[int]$delivery.version}
    $expected=if ($script:evaluation) {[int]$script:evaluation.version} else {0}
    $payload=@{expected=$expected;deliveryVersions=$versions;documentRatings=$ratingRows;rankings=$rankingRows}
    $request=Get-Request "evaluation:$($case.caseId)" $payload
    Set-Status '평가 저장 중…'
    Invoke-UiWork -Work {param($ctx,$case,$payload,$request) Save-Evaluation -Context $ctx -Case $case -ExpectedVersion $payload.expected -RequestId $request -DeliveryVersions $payload.deliveryVersions -DocumentRatings $payload.documentRatings -Rankings $payload.rankings} -Arguments @($context,$case,$payload,$request) -Completed {
        param($result)
        $script:evaluation=$result[0];$script:evalDirty=$false
        $script:evalBaselineStamp=Get-EvaluationInputStamp
        $ui.EvaluationSavedText.Text='저장된 평가'
        $ui.EvaluationStateText.Text='평가를 저장했습니다.'
        Set-Status '평가 저장 확인 완료.'
    }
}
function Export-Current([bool]$IncludeAttachments,[string]$CaseId='') {
    $period=Get-SelectedPeriod;$scope=Get-WorkspaceScope
    $folder=Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'AI검색평가'
    Set-Status '다운로드 파일 생성 중…'
    Invoke-UiWork -Work {
        param($ctx,$period,$scope,$folder,$attachments,$caseId)
        $exportArgs=@{Context=$ctx;Period=$period;Scope=$scope;OutputDirectory=$folder;IncludeAttachments=$attachments}
        if ($caseId) {$exportArgs['CaseId']=$caseId}
        Export-ReviewSnapshot @exportArgs
    } -Arguments @($context,$period,$scope,$folder,$IncludeAttachments,$CaseId) -Completed {
        param($result)
        $export=$result[0]
        $paths=@($export.workbookPath,$export.bundlePath) | Where-Object {$_}
        Set-Status "다운로드 완료 · 출원 $($export.caseCount)건 · 문헌 $($export.documentCount)건"
        [Windows.MessageBox]::Show($window,"문서 폴더의 AI검색평가 폴더에 저장했습니다.`n$($paths -join "`n")",'다운로드 완료','OK','Information') | Out-Null
    }
}

. (Join-Path $appDirectory 'Ui\Admin.ps1')
Initialize-AdminUi
. (Join-Path $appDirectory 'Ui/Statistics.ps1')
Initialize-StatisticsUi
. (Join-Path $appDirectory 'Ui/Layout.ps1')
foreach ($list in @($ui.CasesGrid,$ui.SubmissionRowsGrid,$ui.PublishedGrid,$ui.RankGrid,$ui.RatingsGrid,$ui.DocumentsGrid,$ui.PeriodMembersGrid,$ui.PeriodAssignmentsGrid,$ui.PeriodSourcesList)) {Enable-ReadableListScroll $list}
foreach ($list in @($ui.StatsGroupsGrid,$ui.StatsAccountsGrid,$ui.StatsCasesGrid,$ui.StatsSourcesGrid)) {Enable-ReadableListScroll $list}

function Set-SessionControls {
    $script:changing=$true
    Set-Identity
    $modeLabel=if ($AdminMode) {'관리자용 실행'} else {'일반사용자용 실행'}
    $ui.ModeText.Text=if ($Demo) {"로컬 시연 · $modeLabel"} else {"$modeLabel · $DataRoot"}
    $ui.SwitchUserButton.Visibility=if ($AdminMode) {'Collapsed'} else {'Visible'}
    $ui.AdminTab.Visibility=if ('Admin' -in $context.User.roles) {'Visible'} else {'Collapsed'}
    $ui.StatisticsTab.Visibility=$ui.AdminTab.Visibility
    Clear-Statistics
    $ui.PersonalWorkTab.Visibility=if ('examiner' -in $context.User.roles) {'Visible'} else {'Collapsed'}
    $ui.DeliveryTab.Visibility=if ('ai' -in $context.User.roles -or 'admin' -in $context.User.roles) {'Visible'} else {'Collapsed'}
    $script:workspaceCases=@{Personal='';AI='';Admin=''};$script:workspaceFilters=@{}
    $script:selectionLoadedKey='';$script:targetSelection=$null;$script:selectionBaselineStamp=Get-SelectionInputStamp
    $ui.SearchBox.Clear();$ui.CaseStateCombo.SelectedIndex=0
    $ui.PeriodCombo.ItemsSource=@(Get-Periods -Context $context);$ui.PeriodCombo.SelectedIndex=0
    Reset-AdminEditor
    $ui.PersonalTabs.SelectedItem=$ui.TargetSelectionTab;$ui.EvaluationTabs.SelectedIndex=1
    $ui.MainTabs.SelectedItem=if ($AdminMode) {$ui.StatisticsTab} elseif ('examiner' -in $context.User.roles) {$ui.PersonalWorkTab} else {$ui.DeliveryTab}
    $script:activeWorkspace=Get-CurrentWorkspace;$script:lastMainTab=$ui.MainTabs.SelectedItem
    Set-SelectionStage
    $script:requests=@{};$script:adminDirty=$false;$script:changing=$false
}
Set-SessionControls
$ui.DocumentsGrid.ItemsSource=$script:documents;$ui.RatingsGrid.ItemsSource=$script:ratings;$ui.RankGrid.ItemsSource=$script:ranks
$ui.SubmissionRowsGrid.ItemsSource=$script:submissionRows
foreach ($grid in @($ui.SubmissionRowsGrid,$ui.DocumentsGrid,$ui.RatingsGrid)) {Enable-GridSingleClick $grid}
$script:documents.Add_CollectionChanged({if (-not $script:changing) {Update-ActionAvailability}})
$ui.SwitchUserButton.Add_Click({Invoke-UiAction {
    if ($script:busy -or -not (Confirm-Discard)) {return}
    if ((Has-PendingSubmissions) -and [Windows.MessageBox]::Show($window,'등록 전 출원 입력을 버리고 사용자를 변경할까요?','사용자 변경','YesNo','Question') -ne 'Yes') {return}
    $selected=Show-LoginDialog $window
    if ($selected) {
        $script:context=$selected;$script:changing=$true
        Clear-Detail
        $script:submissionRows.Clear();$ui.ApplicationNumberBox.Clear();$ui.CpcBox.Clear();$ui.SubmissionBox.Clear();$ui.SubmissionResult.Text=''
        $script:editDirty=$false;$script:evalDirty=$false;$script:adminDirty=$false
        Set-SessionControls;Load-Cases
    }
}})
$ui.RefreshButton.Add_Click({Invoke-UiAction {
    if (-not (Confirm-Discard)) {return}
    if (-not $AdminMode -and (Has-PendingSubmissions) -and [Windows.MessageBox]::Show($window,'임시저장하지 않은 대상 건 변경을 버리고 저장된 목록을 불러올까요?','새로 고침','YesNo','Question') -ne 'Yes') {return}
    if (-not $AdminMode) {$script:selectionLoadedKey=''}
    Load-Cases
}})
$ui.PeriodCombo.Add_SelectionChanged({if (-not $script:changing -and -not $script:busy) {Invoke-UiAction {
    $canChange=Confirm-Discard
    if ($canChange -and (Has-PendingSubmissions)) {$canChange=[Windows.MessageBox]::Show($window,'임시저장하지 않은 대상 건 변경을 버리고 운영월을 바꿀까요?','운영월 변경','YesNo','Question') -eq 'Yes'}
    if ($canChange) {$script:selectionLoadedKey='';Load-Cases} else {$script:changing=$true;$ui.PeriodCombo.SelectedItem=$script:lastPeriod;$script:changing=$false}
}}})
$ui.SearchBox.Add_TextChanged({if (-not $script:changing) {Update-Filter}})
$ui.CaseStateCombo.Add_SelectionChanged({if (-not $script:changing -and -not $script:busy) {Update-Filter}})
foreach ($box in @($ui.ApplicationNumberBox,$ui.CpcBox)) {$box.Add_KeyDown({param($sender,$eventArgs) if ($eventArgs.Key -eq 'Return') {Invoke-UiAction {Add-SubmissionInputRows $true $false;[void]$ui.ApplicationNumberBox.Focus()};$eventArgs.Handled=$true}})}
$ui.CasesGrid.Add_SelectionChanged({
    if ($script:changing -or $script:busy -or -not $ui.CasesGrid.SelectedItem) {return}
    Invoke-UiAction {if (Confirm-Discard) {Show-CaseDetail $ui.CasesGrid.SelectedItem} else {$script:changing=$true;$ui.CasesGrid.SelectedItem=$script:lastCase;$script:changing=$false}}
})
$ui.SubmitButton.Add_Click({Invoke-UiAction {
    if (-not $AdminMode) {Save-SelectionEditor Confirmed;return}
    if (-not (Confirm-Discard)) {return}
    $period=Get-SelectedPeriod
    Add-SubmissionInputRows
    [void]$ui.SubmissionRowsGrid.CommitEdit([Windows.Controls.DataGridEditingUnit]::Cell,$true);[void]$ui.SubmissionRowsGrid.CommitEdit([Windows.Controls.DataGridEditingUnit]::Row,$true)
    if (-not $script:submissionRows.Count) {throw '출원번호와 CPC를 입력하거나 붙여넣으세요.'}
    $jobs=@(foreach ($row in $script:submissionRows) {
        $number=[string]$row.applicationNumber;$cpc=[string]$row.cpc
        [pscustomobject]@{rowKey=$row.rowKey;number=$number;cpc=$cpc;request=(Get-Request "submit:$period`:$number" @{period=$period;number=$number;cpc=$cpc})}
    })
    Set-Status '출원번호 확인 및 등록 중…'
    Invoke-UiWork -Work {
        param($ctx,$period,$jobs)
        foreach ($job in $jobs) {try {$record=New-Submission -Context $ctx -Period $period -ApplicationNumber $job.number -Cpc $job.cpc -RequestId $job.request;[pscustomobject]@{rowKey=$job.rowKey;number=$job.number;success=$true;message=$(if ($record.duplicateWarning) {'등록 확인 · 다른 심사관과 중복 확인 필요'} else {'등록 확인'})}} catch {[pscustomobject]@{rowKey=$job.rowKey;number=$job.number;success=$false;message=$_.Exception.Message}}}
    } -Arguments @($context,$period,$jobs) -Completed {
        param($result)
        $ui.SubmissionResult.Text=($result | ForEach-Object {"$($_.number) · $($_.message)"}) -join "`n"
        foreach ($entry in $result) {
            $row=@($script:submissionRows | Where-Object {$_.rowKey -eq $entry.rowKey}) | Select-Object -First 1
            if ($row) {if ($entry.success) {[void]$script:submissionRows.Remove($row)} else {$row.message=$entry.message}}
        }
        $ui.SubmissionRowsGrid.Items.Refresh()
        Load-Cases
    }
}})
$ui.SelectionDraftButton.Add_Click({Invoke-UiAction {Save-SelectionEditor Draft}})
$ui.AddSubmissionRowButton.Add_Click({Invoke-UiAction {Add-SubmissionInputRows $true $false}})
$ui.AddPastedRowsButton.Add_Click({Invoke-UiAction {Add-SubmissionInputRows $false $true}})
$ui.RemoveSubmissionRowButton.Add_Click({if ($ui.SubmissionRowsGrid.SelectedItem) {[void]$script:submissionRows.Remove($ui.SubmissionRowsGrid.SelectedItem)}})
$ui.AddResultNumbersButton.Add_Click({Invoke-UiAction {Add-ResultNumberRows}})
$ui.CancelSubmissionButton.Add_Click({Invoke-UiAction {
    if (-not (Confirm-Discard)) {return}
    $case=Require-Case
    $reason=''
    if ([Windows.MessageBox]::Show($window,"$($case.applicationNumber) 출원을 취소할까요?",'제출 취소','YesNo','Question') -ne 'Yes') {return}
    $request=Get-Request "cancel:$($case.caseId)" @{expected=$case.version;reason=$reason}
    Invoke-UiWork -Work {param($ctx,$case,$reason,$request) Cancel-Submission -Context $ctx -Case $case -ExpectedVersion $case.version -RequestId $request -Reason $reason} -Arguments @($context,$case,$reason,$request) -Completed {param($result) Load-Cases}
}})
$ui.LoadDeliveryButton.Add_Click({Invoke-UiAction {if (Confirm-Discard) {Load-DeliveryEditor}}})
$ui.GoEvaluationButton.Add_Click({$ui.EvaluationTabs.SelectedIndex=1})
$ui.MainTabs.Add_SelectionChanged({
    param($sender,$eventArgs)
    if ($eventArgs.Source -ne $ui.MainTabs -or $script:changing -or $script:busy) {return}
    Invoke-UiAction {Change-Workspace}
})
$ui.PersonalTabs.Add_SelectionChanged({
    param($sender,$eventArgs)
    if ($eventArgs.Source -ne $ui.PersonalTabs -or $script:changing -or $script:busy -or $script:activeWorkspace -ne 'Personal') {return}
    Set-WorkspacePresentation
    if ($ui.PersonalTabs.SelectedItem -eq $ui.EvaluationTab -and $script:selectedCase -and -not $script:caseDetailLoaded) {Invoke-UiAction {Show-CaseDetail $script:selectedCase}}
})
$ui.SourceCombo.Add_SelectionChanged({
    if ($script:changing -or $script:busy -or -not $script:selectedCase -or -not $ui.SourceCombo.SelectedItem -or $ui.MainTabs.SelectedItem -ne $ui.DeliveryTab) {return}
    Invoke-UiAction {
        if ($script:editDirty -and -not (Confirm-Discard)) {$script:changing=$true;$ui.SourceCombo.SelectedItem=@($ui.SourceCombo.ItemsSource | Where-Object {$_.sourceId -eq $script:loadedSource}) | Select-Object -First 1;$script:changing=$false;return}
        Load-DeliveryEditor
    }
})
$ui.AddDocumentButton.Add_Click({Invoke-UiAction {
    if (-not $script:loadedSource) {throw '편집 자료를 먼저 불러오세요.'}
    if ($script:documents.Count -ge 7) {throw '조사 방식별 문헌은 최대 7건입니다.'}
    $script:documents.Add([pscustomobject]@{documentId=[guid]::NewGuid().ToString('N');type='특허';identifier='';title='';description='';attachments=@()});$script:editDirty=$true
    $row=$script:documents[$script:documents.Count-1]
    $ui.DocumentsGrid.SelectedItem=$row;$ui.DocumentsGrid.ScrollIntoView($row)
    $ui.DocumentsGrid.CurrentCell=New-Object Windows.Controls.DataGridCellInfo($row,$ui.DocumentsGrid.Columns[0])
    [void]$ui.DocumentsGrid.Focus();[void]$ui.DocumentsGrid.BeginEdit()
}})
$ui.RemoveDocumentButton.Add_Click({if ($ui.DocumentsGrid.SelectedItem) {[void]$script:documents.Remove($ui.DocumentsGrid.SelectedItem);$script:editDirty=$true}})
$ui.DraftButton.Add_Click({Invoke-UiAction {Save-DeliveryEditor 'Draft'}})
$ui.PublishButton.Add_Click({Invoke-UiAction {Save-DeliveryEditor 'Published'}})
$ui.SaveEvaluationButton.Add_Click({Invoke-UiAction {Save-EvaluationEditor}})
$ui.ExportButton.Add_Click({Invoke-UiAction {Export-Current $false}})
$ui.ZipButton.Add_Click({Invoke-UiAction {Export-Current $true}})
$ui.CaseExportButton.Add_Click({Invoke-UiAction {$case=Require-Case;Export-Current $true $case.caseId}})
$ui.ResultNumbersBox.Add_TextChanged({if (-not $script:changing) {$script:editDirty=$true}})
$ui.DocumentsGrid.Add_CellEditEnding({if (-not $script:changing) {$script:editDirty=$true}})
$ui.RatingsGrid.Add_CellEditEnding({if (-not $script:changing) {$script:evalDirty=$true}})
$ui.RankGrid.Add_CellEditEnding({if (-not $script:changing) {$script:evalDirty=$true}})
foreach ($grid in @($ui.RatingsGrid,$ui.RankGrid)) {
    $grid.AddHandler([Windows.Controls.Primitives.Selector]::SelectionChangedEvent,[Windows.Controls.SelectionChangedEventHandler]{
        param($sender,$eventArgs)
        if (-not $script:changing -and -not $script:busy -and $eventArgs.OriginalSource -is [Windows.Controls.ComboBox]) {
            [void]$sender.Dispatcher.BeginInvoke([Action]{Update-EvaluationDirty})
        }
    })
}
$window.Add_Closing({
    param($sender,$eventArgs)
    if ($script:busy) {$eventArgs.Cancel=$true;Set-Status '진행 중인 저장·조회가 끝난 뒤 종료하세요.'}
    elseif (-not (Confirm-Discard)) {$eventArgs.Cancel=$true}
    elseif ((Has-PendingSubmissions) -and [Windows.MessageBox]::Show($window,'등록하지 않은 출원 입력이 있습니다. 입력내용을 버리고 종료할까요?','출원 입력 확인','YesNo','Question') -ne 'Yes') {$eventArgs.Cancel=$true}
})
$window.Add_ContentRendered({Invoke-UiAction {Load-Cases}})
if ($SmokeTest) {
    $script:smokeTimer=New-Object Windows.Threading.DispatcherTimer
    $script:smokeTimer.Interval=[TimeSpan]::FromSeconds(1)
    $script:smokeSelected=$false;$script:smokeStage=0;$script:selectionStepPreserved=$false
    $script:smokeTimer.Add_Tick({
        if ($script:busy) {return}
        if ($SmokeWorkspace -and -not $script:smokeWorkspaceSelected -and -not $script:smokeFailure) {
            $script:smokeWorkspaceSelected=$true
            Invoke-UiAction {Select-Workspace $SmokeWorkspace}
            return
        }
        if ($ExerciseStatistics -and -not $script:smokeFailure -and $script:smokeStage -lt 6) {
            Invoke-UiAction {
                switch ($script:smokeStage) {
                    0 {
                        if (-not $script:statisticsSnapshot -or $script:activeWorkspace -ne 'Statistics' -or $ui.CaseSidebar.Visibility -ne 'Collapsed' -or $ui.StatsCasesGrid.Items.Count -lt 1) {throw '통계 첫 화면의 출원 집계 또는 전체 너비 배치가 올바르지 않습니다.'}
                        $script:statsOriginalPeriod=Get-SelectedPeriod
                        Save-TestScreenshot '-Overview'
                        $ui.StatsGroupsGrid.SelectedItem=@($ui.StatsGroupsGrid.Items | Where-Object {$_.caseCount -gt 0}) | Select-Object -First 1
                    }
                    1 {
                        if (-not $ui.StatsGroupCombo.SelectedValue -or $ui.StatsDetailTabs.SelectedIndex -ne 1) {throw '그룹 클릭 후 계정별 현황으로 이동하지 않았습니다.'}
                        $ui.StatsAccountsGrid.SelectedItem=@($ui.StatsAccountsGrid.Items | Where-Object {$_.caseCount -gt 0}) | Select-Object -First 1
                        if (-not $ui.StatsUserCombo.SelectedValue -or $ui.StatsDetailTabs.SelectedIndex -ne 2 -or $script:statisticsFilter.summary.accountCount -ne 1 -or $ui.StatsAccountsGrid.Items.Count -ne 1 -or @($ui.StatsCasesGrid.Items | Where-Object {$_.userId -ne $ui.StatsUserCombo.SelectedValue}).Count) {throw '계정 클릭 후 다른 계정의 출원이나 집계가 남았습니다.'}
                        Save-TestScreenshot '-Account'
                        Export-StatisticsView -OutputDirectory ([IO.Path]::GetDirectoryName($SmokeReportPath))
                    }
                    2 {
                        if (-not [IO.File]::Exists($script:statisticsExportPath)) {throw '조회 범위의 통계 엑셀이 생성되지 않았습니다.'}
                        $script:statsExportCaseCount=$script:statisticsFilter.cases.Count
                        $ui.StatsDetailTabs.SelectedIndex=3;Save-TestScreenshot '-Sources'
                        $empty=@($ui.PeriodCombo.Items | Where-Object {$_.period -ne $script:statsOriginalPeriod}) | Select-Object -First 1
                        if (-not $empty) {throw '빈 운영월 검증 자료가 없습니다.'}
                        $ui.PeriodCombo.SelectedItem=$empty
                    }
                    3 {
                        if ($script:statisticsFilter.summary.caseCount -ne 0 -or $ui.StatsCasesGrid.Items.Count -ne 0 -or $ui.StatsCasesText.Text -ne '0') {throw '빈 운영월에 이전 월의 통계가 남았습니다.'}
                        $ui.StatsResetButton.RaiseEvent((New-Object Windows.RoutedEventArgs([Windows.Controls.Button]::ClickEvent)))
                        if ($ui.StatsGroupCombo.SelectedValue -or $ui.StatsUserCombo.SelectedValue) {throw '전체 보기에서 필터가 초기화되지 않았습니다.'}
                        Save-TestScreenshot '-Empty'
                        $ui.PeriodCombo.SelectedItem=@($ui.PeriodCombo.Items | Where-Object {$_.period -eq $script:statsOriginalPeriod}) | Select-Object -First 1
                    }
                    4 {
                        if ($script:statisticsFilter.summary.caseCount -lt 1 -or $script:statisticsSnapshot.period -ne $script:statsOriginalPeriod) {throw '운영월 복귀 후 통계가 갱신되지 않았습니다.'}
                        $window.Width=$window.MinWidth;$window.Height=$window.MinHeight
                        $ui.StatsScroll.ScrollToBottom();$window.UpdateLayout();Save-TestScreenshot '-Minimum'
                    }
                    5 {
                        $ui.StatsScroll.ScrollToTop();$window.UpdateLayout()
                        $script:statisticsFlowVerified=$true
                    }
                }
                $script:smokeStage++
            }
            return
        }
        if ($ExerciseUiWrites -and -not $script:smokeSelected -and $script:activeWorkspace -eq 'Statistics' -and -not $script:smokeFailure) {
            Invoke-UiAction {
                if (-not $script:statisticsSnapshot -or $ui.CaseSidebar.Visibility -ne 'Collapsed') {throw '관리자 첫 화면에 전체 통계가 표시되지 않았습니다.'}
                $script:statisticsLandingVerified=$true
                Save-TestScreenshot '-Statistics'
                Select-Workspace 'Admin'
            }
            return
        }
        if ($ExerciseTargetSelection -and -not $script:smokeFailure -and $script:smokeStage -lt 3) {
            Invoke-UiAction {
                switch ($script:smokeStage) {
                    0 {
                        if ($ui.PersonalTabs.SelectedItem -ne $ui.TargetSelectionTab -or $ui.EvaluationTab.Visibility -ne 'Collapsed' -or $ui.CaseSidebar.Visibility -ne 'Collapsed' -or $ui.DeliveryTab.Visibility -ne 'Collapsed') {throw '미확정 심사관에게 대상 건 선정 화면만 표시되지 않았습니다.'}
                        if ($ExerciseTargetSelection -eq 'Draft') {
                            if ($script:submissionRows.Count -ne 0) {throw '신규 월의 대상 건 목록이 비어 있지 않습니다.'}
                            $ui.ApplicationNumberBox.Text='10-2026-0000888';$ui.CpcBox.Text='G06F 16/00'
                            $ui.SelectionDraftButton.RaiseEvent((New-Object Windows.RoutedEventArgs([Windows.Controls.Button]::ClickEvent)))
                        } else {
                            if ($script:submissionRows.Count -ne 1 -or $script:submissionRows[0].cpc -ne 'G06F 16/00') {throw '다시 로그인한 뒤 운영월의 초안이 복원되지 않았습니다.'}
                            $ui.SubmitButton.RaiseEvent((New-Object Windows.RoutedEventArgs([Windows.Controls.Button]::ClickEvent)))
                        }
                    }
                    1 {
                        if ($ExerciseTargetSelection -eq 'Draft') {
                            if ($script:targetSelection.status -ne 'Draft' -or (Has-PendingSubmissions) -or $script:cases.Count -ne 0) {throw '초안 저장 또는 미확정 건 숨김이 적용되지 않았습니다.'}
                        } else {
                            if ($ui.PersonalTabs.SelectedItem -ne $ui.EvaluationTab -or $ui.TargetSelectionTab.Visibility -ne 'Collapsed' -or $ui.CaseSidebar.Visibility -ne 'Visible' -or $script:cases.Count -ne 1) {throw '확정 뒤 평가 화면으로 전환되지 않았습니다.'}
                            if ($script:cases[0].statusLabel -ne '대기 중' -or (@($script:ranks.sourceId) -join '|') -ne 'S01|S02|S03') {throw '입력 대기 상태 또는 전체 조사 방식이 표시되지 않았습니다.'}
                            $window.UpdateLayout()
                            if ($ui.RatingsGrid.ActualHeight -le $ui.RankGrid.ActualHeight) {throw '문헌별 평가 영역이 품질 순위 영역보다 높지 않습니다.'}
                            $script:smokeOriginalWidth=$window.Width;$script:smokeOriginalHeight=$window.Height
                            $window.Width=$window.MinWidth;$window.Height=$window.MinHeight;$window.UpdateLayout()
                            $savePoint=$ui.SaveEvaluationButton.TranslatePoint((New-Object Windows.Point(0,0)),$window)
                            if ($ui.RatingsGrid.ActualHeight -le $ui.RankGrid.ActualHeight -or $savePoint.Y+$ui.SaveEvaluationButton.ActualHeight -gt $window.ActualHeight-30) {throw '최소 창 크기에서 문헌 평가 영역이나 저장 버튼이 잘립니다.'}
                        }
                        $script:selectionFlowVerified=$true
                    }
                    2 {
                        if ($ExerciseTargetSelection -eq 'Confirm') {
                            Save-TestScreenshot '-Small'
                            $window.Width=$script:smokeOriginalWidth;$window.Height=$script:smokeOriginalHeight;$window.UpdateLayout()
                        }
                    }
                }
                $script:smokeStage++
            }
            return
        }
        if (-not $script:smokeSelected -and $ui.CasesGrid.Items.Count -gt 0 -and -not $script:smokeFailure) {
            $script:smokeSelected=$true;$ui.CasesGrid.SelectedIndex=0
            $script:selectionStepPreserved=$script:activeWorkspace -eq 'Personal' -and $ui.PersonalTabs.SelectedItem -eq $ui.TargetSelectionTab
            return
        }
        if ($ExerciseUiWrites -and $script:smokeSelected -and -not $script:smokeFailure -and $script:smokeStage -lt 24) {
            Invoke-UiAction {
                switch ($script:smokeStage) {
                    0 {$script:smokeOriginalCaseId=$script:selectedCase.caseId;Select-Workspace 'Admin'}
                    1 {
                        $ui.AdminKindCombo.SelectedIndex=0
                        $script:smokeUserCount=$script:adminOptions.Users.Count
                        $ui.AdminUserTabs.SelectedIndex=1
                        if ($ui.AdminUserSelectionPanel.Visibility -ne 'Collapsed' -or $ui.AdminSaveButton.Content -ne '사용자 등록' -or $ui.UserNameBox.Text -or $script:adminLoadedId) {throw '신규 등록 화면에 기존 사용자 선택이나 정보가 남습니다.'}
                    }
                    2 {
                        $ui.UserNameBox.Text='화면 등록 심사관';$ui.UserBureauBox.Text='화면 등록 심사국';$ui.UserDivisionBox.Text='화면 등록 심사과'
                        Save-TestScreenshot '-Users'
                        $ui.AdminSaveButton.RaiseEvent((New-Object Windows.RoutedEventArgs([Windows.Controls.Button]::ClickEvent)))
                    }
                    3 {
                        if ($ui.AdminUserTabs.SelectedIndex -ne 1 -or $ui.UserNameBox.Text -or $script:adminLoadedId -or $ui.AdminUserCombo.SelectedItem -or $script:adminDirty -or $ui.AdminSavedText.Text -notlike '*등록 완료') {throw '사용자 등록 후 신규 양식과 완료 안내가 표시되지 않았습니다.'}
                        $script:smokeRegisteredUserId=$script:adminLastUserId
                        $newUserId=$script:smokeRegisteredUserId
                        $ui.AdminUserTabs.SelectedIndex=0
                        if ($script:adminLoadedId -ne $newUserId -or $ui.AdminUserSelectionPanel.Visibility -ne 'Visible' -or $ui.AdminSaveButton.Content -ne '변경사항 저장') {throw '기존 사용자 관리 탭에서 방금 등록한 사용자를 조회하지 못했습니다.'}
                        $ui.AdminUserCombo.SelectedItem=@($ui.AdminUserCombo.ItemsSource | Where-Object {$_.userId -eq 'U002'}) | Select-Object -First 1
                        $ui.AdminUserCombo.SelectedItem=@($ui.AdminUserCombo.ItemsSource | Where-Object {$_.userId -eq $newUserId}) | Select-Object -First 1
                        $script:automaticUserLoadVerified=$ui.UserNameBox.Text -eq '화면 등록 심사관' -and $ui.UserDivisionBox.Text -eq '화면 등록 심사과'
                        if (-not $script:automaticUserLoadVerified) {throw '사용자 선택 즉시 폼이 표시되지 않았습니다.'}
                        Select-Workspace 'Selection'
                    }
                    4 {
                        $script:selectionStepPreserved=$ui.PersonalTabs.SelectedItem -eq $ui.TargetSelectionTab
                        $ui.ApplicationNumberBox.Text='10-2026-0000456';$ui.CpcBox.Text='H04L 9/00';Add-SubmissionInputRows $true $false
                        $ui.SubmissionBox.Text="bad-number`tB25J 9/00";Add-SubmissionInputRows $false $true
                        Save-TestScreenshot '-Submissions'
                        $ui.SubmitButton.RaiseEvent((New-Object Windows.RoutedEventArgs([Windows.Controls.Button]::ClickEvent)))
                    }
                    5 {
                        if ($script:submissionRows.Count -ne 1 -or $script:submissionRows[0].cpc -ne 'B25J 9/00' -or -not $script:submissionRows[0].message) {throw '실패한 출원 행의 CPC와 입력내용이 보존되지 않았습니다.'}
                        $script:submissionRows.Clear()
                        $ui.CasesGrid.SelectedItem=@($script:cases | Where-Object {$_.caseId -eq $script:smokeOriginalCaseId})[0]
                        Select-Workspace 'Evaluation'
                    }
                    6 {$script:workspaceCases.AI=$script:smokeOriginalCaseId;Select-Workspace 'AI'}
                    7 {
                        $window.UpdateLayout()
                        $numberCell=@(Find-VisualControls $ui.DocumentsGrid ([Windows.Controls.DataGridCell]) | Where-Object {$_.Column.DisplayIndex -eq 0}) | Select-Object -First 1
                        if (-not $numberCell) {throw '문헌번호 편집 셀을 찾을 수 없습니다.'}
                        $mouseEvent=New-Object Windows.Input.MouseButtonEventArgs([Windows.Input.Mouse]::PrimaryDevice,[Environment]::TickCount,[Windows.Input.MouseButton]::Left)
                        $mouseEvent.RoutedEvent=[Windows.UIElement]::PreviewMouseDownEvent
                        $numberCell.RaiseEvent($mouseEvent)
                        $script:singleClickEditingVerified=[bool]$numberCell.IsEditing
                        if (-not $script:singleClickEditingVerified) {throw '문헌번호 셀을 한 번 클릭해 편집할 수 없습니다.'}
                        [void]$ui.DocumentsGrid.CommitEdit([Windows.Controls.DataGridEditingUnit]::Cell,$true);[void]$ui.DocumentsGrid.CommitEdit([Windows.Controls.DataGridEditingUnit]::Row,$true)
                        $ui.ResultNumbersBox.Text="10-2025-0000007`n10-2025-0000008"
                        Save-TestScreenshot '-Results'
                        Save-DeliveryEditor 'Draft'
                    }
                    8 {Save-DeliveryEditor 'Published'}
                    9 {
                        if (-not $script:automaticDeliveryRefreshVerified) {throw '결과 등록 후 조회가 자동 갱신되지 않았습니다.'}
                        Select-Workspace 'Evaluation'
                    }
                    10 {
                        $window.UpdateLayout()
                        $combos=@(Find-VisualControls $ui.RatingsGrid ([Windows.Controls.ComboBox]))
                        $scoreCombo=@($combos | Where-Object {$binding=$_.GetBindingExpression([Windows.Controls.Primitives.Selector]::SelectedValueProperty);$binding -and $binding.ParentBinding.Path.Path -eq 'scoreText'}) | Select-Object -First 1
                        $citeCombo=@($combos | Where-Object {$binding=$_.GetBindingExpression([Windows.Controls.Primitives.Selector]::SelectedValueProperty);$binding -and $binding.ParentBinding.Path.Path -eq 'citedText'}) | Select-Object -First 1
                        if (-not $scoreCombo -or -not $citeCombo) {throw '실제 평가 드롭다운을 찾을 수 없습니다.'}
                        $scoreCombo.SelectedValue='5';Update-EvaluationDirty
                        $changedBySelection=$script:evalDirty
                        $scoreCombo.SelectedValue='0';$citeCombo.SelectedValue='O';$citeCombo.SelectedValue='X'
                        $rankCombo=@(Find-VisualControls $ui.RankGrid ([Windows.Controls.ComboBox])) | Select-Object -First 1
                        if (-not $rankCombo) {throw '실제 순위 드롭다운을 찾을 수 없습니다.'}
                        $priorRank=$rankCombo.SelectedValue;$rankCombo.SelectedValue='';$rankCombo.SelectedValue=$priorRank
                        $script:dropdownBindingsVerified=$scoreCombo.DataContext.scoreText -eq '0' -and $citeCombo.DataContext.citedText -eq 'X' -and $rankCombo.DataContext.choice -eq $priorRank -and $changedBySelection
                        if (-not $script:dropdownBindingsVerified) {throw '드롭다운 선택이 평가값과 수정상태에 반영되지 않았습니다.'}
                        Save-EvaluationEditor
                    }
                    11 {Select-Workspace 'Admin'}
                    12 {$ui.AdminKindCombo.SelectedIndex=3;$window.UpdateLayout()}
                    13 {
                        if ($script:adminDirty) {throw '월 운영 폼을 표시한 것만으로 수정 상태가 되었습니다.'}
                        $script:adminFormDisplayVerified=$ui.PeriodMembersGrid.Items.Count -gt 0 -and $ui.PeriodAssignmentsGrid.Columns.Count -eq 3
                        if (-not $script:adminFormDisplayVerified) {throw '월 운영의 참여자·AI 배정 폼을 표시하지 못했습니다.'}
                        Save-TestScreenshot '-Period'
                        Change-PeriodStatus 'Closed'
                    }
                    14 {if ($script:loadedPeriod.status -ne 'Closed') {throw '사유 없이 월 마감이 저장되지 않았습니다.'};Change-PeriodStatus 'Open'}
                    15 {if ($script:loadedPeriod.status -ne 'Open') {throw '사유 없이 재개방이 저장되지 않았습니다.'};Select-Workspace 'Evaluation'}
                    16 {
                        $script:scrollTestHeight=$window.Height;$window.Height=$window.MinHeight;$window.UpdateLayout()
                        $scroll=$window.FindName('EvaluationScroll');$scroll.ScrollToBottom();$window.UpdateLayout()
                        $script:scrollTestOffset=$scroll.VerticalOffset
                        if ($script:scrollTestOffset -le 0) {throw '출원 전환 스크롤 검증에 필요한 스크롤 영역이 없습니다.'}
                        Show-CaseDetail $script:selectedCase
                    }
                    17 {
                        $window.UpdateLayout()
                        if ([Math]::Abs($window.FindName('EvaluationScroll').VerticalOffset-$script:scrollTestOffset) -gt 1) {throw '같은 출원을 새로 조회할 때 평가 위치가 초기화됩니다.'}
                        $ui.CasesGrid.SelectedItem=@($script:cases | Where-Object {$_.caseId -ne $script:selectedCase.caseId}) | Select-Object -First 1
                    }
                    18 {
                        $window.UpdateLayout()
                        if ($window.FindName('EvaluationScroll').VerticalOffset -gt 1) {throw '다른 출원을 선택해도 이전 평가의 스크롤 위치가 남습니다.'}
                        $script:caseScrollVerified=$true
                        $ui.CasesGrid.SelectedItem=@($script:cases | Where-Object {$_.caseId -eq $script:smokeOriginalCaseId}) | Select-Object -First 1
                    }
                    19 {$window.Height=$script:scrollTestHeight;$window.UpdateLayout()}
                    20 {Select-Workspace 'Admin'}
                    21 {
                        $ui.AdminKindCombo.SelectedIndex=0;$ui.AdminUserTabs.SelectedIndex=0
                        $ui.AdminUserCombo.SelectedItem=@($ui.AdminUserCombo.ItemsSource | Where-Object {$_.userId -eq $script:smokeRegisteredUserId}) | Select-Object -First 1
                        $ui.UserBureauBox.Text='수정한 심사국';$ui.UserAiCheck.IsChecked=$true;$ui.UserActiveCheck.IsChecked=$false
                        $ui.AdminSaveButton.RaiseEvent((New-Object Windows.RoutedEventArgs([Windows.Controls.Button]::ClickEvent)))
                    }
                    22 {
                        if ($script:adminLoadedId -ne $script:smokeRegisteredUserId -or $script:adminVersion -ne 2 -or $script:adminOptions.Users.Count -ne $script:smokeUserCount+1 -or $ui.AdminUserTabs.SelectedIndex -ne 0 -or $ui.UserActiveCheck.IsChecked) {throw '기존 사용자 수정이 신규 등록으로 처리되거나 수정 화면이 유지되지 않았습니다.'}
                        Save-TestScreenshot '-Users-Manage'
                        $ui.AdminUserTabs.SelectedIndex=1
                        if ($ui.UserNameBox.Text -or $ui.UserBureauBox.Text -or $ui.UserDivisionBox.Text -or -not $ui.UserExaminerCheck.IsChecked -or $ui.UserAiCheck.IsChecked -or $ui.UserAdminCheck.IsChecked -or -not $ui.UserActiveCheck.IsChecked) {throw '신규 등록 양식에 수정한 사용자의 소속·역할·비활성 상태가 남았습니다.'}
                        $ui.UserNameBox.Text='연속 등록 심사관'
                        $ui.AdminSaveButton.RaiseEvent((New-Object Windows.RoutedEventArgs([Windows.Controls.Button]::ClickEvent)))
                    }
                    23 {
                        if ($script:adminOptions.Users.Count -ne $script:smokeUserCount+2 -or $script:adminLastUserId -eq $script:smokeRegisteredUserId -or $ui.UserNameBox.Text -or $script:adminDirty) {throw '연속 등록이 별개의 사용자로 저장되지 않았습니다.'}
                        Save-TestScreenshot '-Users-Registered'
                        $ui.AdminUserTabs.SelectedIndex=0
                        $ui.AdminUserCombo.SelectedItem=@($ui.AdminUserCombo.ItemsSource | Where-Object {$_.userId -eq $script:smokeRegisteredUserId}) | Select-Object -First 1
                        if ($ui.UserBureauBox.Text -ne '수정한 심사국' -or $ui.UserActiveCheck.IsChecked -or -not $ui.UserAiCheck.IsChecked -or $ui.AdminUserCombo.SelectedItem.displayName -notlike '*(비활성)') {throw '비활성 사용자의 변경사항을 다시 불러오지 못했습니다.'}
                        $script:separateUserFlowsVerified=$true
                        Select-Workspace 'Evaluation'
                    }
                }
                $script:smokeStage++
            }
            return
        }
        if ($ExerciseWorkspaceSwitch -and $script:smokeSelected -and -not $script:smokeFailure -and $script:smokeStage -lt 5) {
            Invoke-UiAction {
                switch ($script:smokeStage) {
                    0 {
                        $script:personalCaseIds=@($script:cases | ForEach-Object {$_.caseId});$script:personalScope=Get-WorkspaceScope
                        Select-Workspace 'Evaluation'
                    }
                    1 {
                        if ((Get-WorkspaceScope) -ne 'Mine' -or $script:selectedCase.userId -ne $context.User.userId) {throw '본인 평가 단계에 담당 그룹 출원이 섞였습니다.'}
                        Save-TestScreenshot '-Evaluation'
                        Select-Workspace 'AI'
                    }
                    2 {
                        $script:aiCaseIds=@($script:cases | ForEach-Object {$_.caseId});$script:aiScope=Get-WorkspaceScope
                        $script:aiSourceIds=@($ui.SourceCombo.ItemsSource | ForEach-Object {$_.sourceId})
                        $other=@($script:cases | Where-Object {$_.userId -ne $context.User.userId}) | Select-Object -First 1
                        if (-not $other) {throw '담당 업무에서 다른 심사관의 대상 건을 조회하지 못했습니다.'}
                        $ui.CasesGrid.SelectedItem=$other
                    }
                    3 {
                        if ($script:activeWorkspace -ne 'AI' -or -not $script:loadedSource -or @($ui.SourceCombo.ItemsSource | Where-Object {$_.type -ne 'AI'}).Count) {throw 'AI 담당 업무에 일반 조사 또는 미배정 방식이 섞였습니다.'}
                        Save-TestScreenshot '-AI'
                        Select-Workspace 'Selection'
                    }
                    4 {
                        $personalOnly=@($script:cases | Where-Object {$_.userId -ne $context.User.userId}).Count -eq 0
                        $script:workspaceSeparationVerified=$personalOnly -and (Get-WorkspaceScope) -eq 'Mine' -and $ui.PersonalTabs.SelectedItem -eq $ui.EvaluationTab -and $script:documents.Count -eq 0 -and $script:selectedCase.userId -eq $context.User.userId
                        if (-not $script:workspaceSeparationVerified) {throw '개인 선정·평가와 AI 담당 업무의 조회 범위 또는 입력내용이 분리되지 않았습니다.'}
                    }
                }
                $script:smokeStage++
            }
            return
        }
        $script:smokeTimer.Stop()
        Save-TestScreenshot
        if ($SmokeReportPath) {
            $report=[pscustomobject]@{
                userId=$context.User.userId;caseIds=@($script:cases | ForEach-Object {$_.caseId})
                menus=@($ui.MainTabs.Items | ForEach-Object {$_.Header})
                visibleMenus=@($ui.MainTabs.Items | Where-Object {$_.Visibility -eq 'Visible'} | ForEach-Object {$_.Header})
                personalSteps=@($ui.PersonalTabs.Items | ForEach-Object {$_.Header})
                visiblePersonalSteps=@($ui.PersonalTabs.Items | Where-Object {$_.Visibility -eq 'Visible'} | ForEach-Object {$_.Header})
                selectionFlowVerified=[bool]$script:selectionFlowVerified;selectionStatus=(Get-UiValue $script:targetSelection 'status');sidebarVisible=$ui.CaseSidebar.Visibility -eq 'Visible'
                evaluationOpened=$ui.PersonalTabs.SelectedItem -eq $ui.EvaluationTab;caseStatusLabels=@($script:cases | ForEach-Object {$_.statusLabel});rankSourceIds=@($script:ranks | ForEach-Object {$_.sourceId})
                rankHeight=$ui.RankGrid.ActualHeight;ratingsHeight=$ui.RatingsGrid.ActualHeight
                workspace=$script:activeWorkspace;ratingCount=$script:ratings.Count;documentCount=$script:documents.Count
                caseListWidth=$ui.CasesGrid.ActualWidth
                sourceIds=@($ui.SourceCombo.ItemsSource | ForEach-Object {$_.sourceId})
                selectionRowCount=$script:submissionRows.Count
                statisticsSummary=$(if ($script:statisticsFilter) {$script:statisticsFilter.summary} else {$null})
                evaluationColumns=@($ui.RatingsGrid.Columns | ForEach-Object {$_.Header})
                adminVisible=($ui.AdminTab.Visibility -eq 'Visible');deliveryVisible=($ui.DeliveryTab.Visibility -eq 'Visible')
                launchMode=$script:launchMode;accessMode=$context.accessMode;loginDialogShown=[bool]$script:loginDialogShown
                scopeIds=@($ui.ScopeCombo.ItemsSource | ForEach-Object {$_.id});workspaceScope=(Get-WorkspaceScope)
                selectionStepPreserved=[bool]$script:selectionStepPreserved;workspaceSeparationVerified=[bool]$script:workspaceSeparationVerified
                aiCaseIds=@($script:aiCaseIds);personalCaseIds=@($script:personalCaseIds);aiSourceIds=@($script:aiSourceIds)
                aiScope=$script:aiScope;personalScope=$script:personalScope
                dropdownBindingsVerified=[bool]$script:dropdownBindingsVerified;automaticDeliveryRefreshVerified=[bool]$script:automaticDeliveryRefreshVerified
                automaticUserLoadVerified=[bool]$script:automaticUserLoadVerified;singleClickEditingVerified=[bool]$script:singleClickEditingVerified
                adminFormDisplayVerified=[bool]$script:adminFormDisplayVerified
                caseScrollVerified=[bool]$script:caseScrollVerified
                separateUserFlowsVerified=[bool]$script:separateUserFlowsVerified
                statisticsVisible=($ui.StatisticsTab.Visibility -eq 'Visible');statisticsLandingVerified=[bool]$script:statisticsLandingVerified
                statisticsFlowVerified=[bool]$script:statisticsFlowVerified;statisticsExportPath=$script:statisticsExportPath;statisticsExportCaseCount=$script:statsExportCaseCount
            }
            [IO.File]::WriteAllText($SmokeReportPath,(ConvertTo-Json -InputObject $report -Depth 10),(New-Object Text.UTF8Encoding($false)))
        }
        $script:submissionRows.Clear();$ui.ApplicationNumberBox.Clear();$ui.CpcBox.Clear();$ui.SubmissionBox.Clear()
        $script:selectionBaselineStamp=Get-SelectionInputStamp
        $script:editDirty=$false;$script:evalDirty=$false;$script:adminDirty=$false;$window.Close()
    });$script:smokeTimer.Start()
}

[void]$window.ShowDialog()
if ($SmokeTest -and $script:smokeFailure) {throw $script:smokeFailure}
