param(
    [string]$DataRoot,
    [switch]$Demo,
    [switch]$SmokeTest,
    [switch]$ExerciseUiWrites,
    [string]$ScreenshotPath,
    [string]$SmokeLoginUserId = 'U001',
    [string]$SmokeReportPath
)
$ErrorActionPreference = 'Stop'
if ($ExerciseUiWrites -and (-not $SmokeTest -or -not $Demo)) {throw '화면 쓰기 검증은 -Demo -SmokeTest에서만 실행할 수 있습니다.'}
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
        try {$login.users=@(& $login.listCommand -Root $login.root);& $login.filter}
        catch {$login.LoginStatusText.Text=$_.Exception.Message}
    }.GetNewClosure()
    $login.LoginSearchBox.Add_TextChanged({& $login.filter}.GetNewClosure())
    $login.LoginRefreshButton.Add_Click({& $login.reload}.GetNewClosure())
    $login.LoginButton.Add_Click({
        if (-not $login.LoginUsersList.SelectedItem) {$login.LoginStatusText.Text='목록에서 본인을 선택하세요.';return}
        try {$login.context=& $login.contextCommand -Root $login.root -UserId $login.LoginUsersList.SelectedItem.userId;$login.window.DialogResult=$true}
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
$script:context = Show-LoginDialog
if (-not $script:context) {return}
$reader = [Xml.XmlReader]::Create((Join-Path $PSScriptRoot 'Views\Main.xaml'))
try { $script:window = [Windows.Markup.XamlReader]::Load($reader) } finally { $reader.Close() }
$script:ui = @{}
$names = @('IdentityText','PeriodCombo','RefreshButton','ScopeCombo','SearchBox','CaseCountText','ExportButton','ZipButton','CasesGrid','CaseTitle','CaseSubtitle','MainTabs','SubmissionBox','SubmitButton','CancelSubmissionButton','SubmissionResult','DeliverySummaryText','PublishedGrid','DeliveryTab','SourceCombo','LoadDeliveryButton','EditVersionText','NoResultsCheck','NoResultsReasonBox','DocumentsGrid','AddDocumentButton','RemoveDocumentButton','AttachButton','HistoryButton','DeliveryReasonBox','DraftButton','PublishButton','EvaluationStateText','RankGrid','RatingsGrid','EvaluationReasonBox','SaveEvaluationButton','EvaluationVersionText','AdminTab','PeriodStateText','ClosePeriodButton','ReopenPeriodButton','AdminReasonBox','AdminKindCombo','AdminIdBox','AdminLoadButton','AdminSaveButton','AdminJsonBox','ModeText','StatusText')
$names += 'CaseExportButton'
$names += @('CaseMetadataText','ApplicationNumberBox','CpcBox','AddSubmissionRowButton','AddPastedRowsButton','RemoveSubmissionRowButton','SubmissionRowsGrid','UserFieldsPanel','UserNameBox','UserBureauBox','UserDivisionBox','UserExaminerCheck','UserAiCheck','UserAdminCheck','UserActiveCheck','AdminUserCombo','NewUserButton','SwitchUserButton','TargetTabs','TargetSelectionTab','EvaluationTab','ResultNumbersBox','AddResultNumbersButton')
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

function Set-Status([string]$Message) { $ui.StatusText.Text = $Message }
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
function Confirm-Discard {
    if (-not ($script:editDirty -or $script:evalDirty -or $script:adminDirty)) { return $true }
    return ([Windows.MessageBox]::Show($window,'저장하지 않은 입력을 버리고 이동할까요?','입력내용 확인','YesNo','Question') -eq 'Yes')
}
function Has-PendingSubmissions {
    return $script:submissionRows.Count -gt 0 -or -not [string]::IsNullOrWhiteSpace($ui.ApplicationNumberBox.Text) -or -not [string]::IsNullOrWhiteSpace($ui.CpcBox.Text) -or -not [string]::IsNullOrWhiteSpace($ui.SubmissionBox.Text)
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
function Set-Identity {
    $affiliation=@((Get-UiValue $context.User 'bureau'),(Get-UiValue $context.User 'division')) | Where-Object {$_}
    $ui.IdentityText.Text=(@($context.User.name,($affiliation -join ' / ')) | Where-Object {$_}) -join ' · '
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
function Update-AdminEditor {
    $isUser=$ui.AdminKindCombo.SelectedItem -and $ui.AdminKindCombo.SelectedItem.Tag -eq 'Users'
    $ui.UserFieldsPanel.Visibility=if ($isUser) {'Visible'} else {'Collapsed'}
    $ui.AdminJsonBox.Visibility=if ($isUser) {'Collapsed'} else {'Visible'}
    $ui.AdminUserCombo.Visibility=if ($isUser) {'Visible'} else {'Collapsed'}
    $ui.NewUserButton.Visibility=if ($isUser) {'Visible'} else {'Collapsed'}
    $ui.AdminIdBox.Visibility=if ($isUser) {'Collapsed'} else {'Visible'}
    $ui.UserFieldsPanel.IsEnabled=$isUser -and ($script:adminIsNewUser -or ($script:adminLoadedKind -eq 'Users' -and $ui.AdminIdBox.Text.Trim() -eq $script:adminLoadedId))
}
function Refresh-AdminUserList([string]$SelectedUserId='') {
    if ('admin' -notin $context.User.roles) {$ui.AdminUserCombo.ItemsSource=@();return}
    $users=@(Get-Users -Context $context | Sort-Object division,name,userId)
    foreach ($user in $users) {
        $division=[string](Get-UiValue $user 'division');if (-not $division) {$division='미지정'}
        $label="$division - $($user.name)";if (-not $user.active) {$label+=' (비활성)'}
        $user | Add-Member -NotePropertyName displayName -NotePropertyValue $label -Force
    }
    $previous=$script:changing;$script:changing=$true
    $ui.AdminUserCombo.ItemsSource=$users
    if ($SelectedUserId) {$ui.AdminUserCombo.SelectedItem=@($users | Where-Object {$_.userId -eq $SelectedUserId}) | Select-Object -First 1}
    $script:changing=$previous
}
function Set-UserForm($Data) {
    $ui.UserNameBox.Text=[string](Get-UiValue $Data 'name')
    $ui.UserBureauBox.Text=[string](Get-UiValue $Data 'bureau')
    $ui.UserDivisionBox.Text=[string](Get-UiValue $Data 'division')
    $roles=@(Get-UiValue $Data 'roles' @('examiner'))
    $ui.UserExaminerCheck.IsChecked='examiner' -in $roles
    $ui.UserAiCheck.IsChecked='ai' -in $roles
    $ui.UserAdminCheck.IsChecked='admin' -in $roles
    $ui.UserActiveCheck.IsChecked=[bool](Get-UiValue $Data 'active' $true)
}
function Read-UserForm {
    $data=ConvertFrom-Json -InputObject $ui.AdminJsonBox.Text
    if ($null -eq $data -or $data -isnot [pscustomobject]) {throw '사용자 정보를 먼저 불러오세요.'}
    $roles=@();if ($ui.UserExaminerCheck.IsChecked) {$roles+='examiner'};if ($ui.UserAiCheck.IsChecked) {$roles+='ai'};if ($ui.UserAdminCheck.IsChecked) {$roles+='admin'}
    $fields=@{name=$ui.UserNameBox.Text.Trim();bureau=$ui.UserBureauBox.Text.Trim();division=$ui.UserDivisionBox.Text.Trim();roles=@($roles);active=[bool]$ui.UserActiveCheck.IsChecked}
    if (-not $script:adminIsNewUser) {$fields.userId=$script:adminLoadedId}
    foreach ($key in $fields.Keys) {$data | Add-Member -NotePropertyName $key -NotePropertyValue $fields[$key] -Force}
    return $data
}
function Update-Filter {
    $search = $ui.SearchBox.Text.Trim()
    $filtered = @($script:cases | Where-Object {
        $text=@($_.applicationNumber,$_.examinerName,(Get-UiValue $_ 'cpc'),(Get-UiValue $_ 'examinerBureau'),(Get-UiValue $_ 'examinerDivision')) -join "`n"
        -not $search -or $text.IndexOf($search,[StringComparison]::OrdinalIgnoreCase) -ge 0
    })
    $ui.CasesGrid.ItemsSource = $filtered
    $ui.CaseCountText.Text = "전체 $($script:cases.Count)건 · 표시 $($filtered.Count)건"
}
function Clear-Detail {
    $script:selectedCase=$null; $script:deliveryEdit=$null; $script:evaluation=$null; $script:loadedSource=$null
    $script:published=@(); $script:editDirty=$false; $script:evalDirty=$false
    $script:documents.Clear(); $script:ratings.Clear(); $script:ranks.Clear()
    $ui.ResultNumbersBox.Clear()
    $ui.PublishedGrid.ItemsSource=@(); $ui.CaseTitle.Text='출원을 등록하거나 목록에서 선택하세요'
    $ui.CaseSubtitle.Text='대상건 선정 → AI 검색 결과 입력 → AI 검색 결과 평가'
    $ui.CaseMetadataText.Text=''
    $ui.EditVersionText.Text=''; $ui.EvaluationVersionText.Text=''
}
function Load-Cases {
    $period = Get-SelectedPeriod
    $scope = [string]$ui.ScopeCombo.SelectedItem.id
    $script:lastPeriod=$ui.PeriodCombo.SelectedItem
    $script:lastScope=$ui.ScopeCombo.SelectedItem
    $script:lastCase=$null
    $script:changing=$true; Clear-Detail; $script:changing=$false
    Set-Status '출원 및 월 운영정보 조회 중…'
    Invoke-UiWork -Work {
        param($ctx,$month,$scope)
        [pscustomobject]@{cases=@(Get-Cases -Context $ctx -Period $month -Scope $scope);period=(Get-Period -Context $ctx -Period $month);sources=@(Get-Sources -Context $ctx)}
    } -Arguments @($context,$period,$scope) -Completed {
        param($result)
        $snapshot=$result[0]; $script:loadedPeriod=$snapshot.period
        $script:cases=@($snapshot.cases)
        foreach ($case in $script:cases) { $case | Add-Member -NotePropertyName statusLabel -NotePropertyValue $(if ($case.status -eq 'Cancelled') {'취소'} elseif ($case.duplicateWarning) {'중복 확인'} else {'접수'}) -Force }
        $script:allSources=@($snapshot.sources)
        $script:changing=$true
        $ui.SourceCombo.ItemsSource=@($snapshot.sources | Where-Object { $_.sourceId -in $script:loadedPeriod.sourceIds })
        $ui.SourceCombo.SelectedIndex=0
        Update-Filter
        $script:changing=$false
        $ui.PeriodStateText.Text="운영월 $($snapshot.period.period) · 상태 $($snapshot.period.status) · 설정 $($snapshot.period.version)판"
        Set-Status '조회 완료 · 목록에서 출원을 선택하세요.'
    }
}
function Show-CaseDetail($Case) {
    Set-Status '납품 및 평가 자료 조회 중…'
    $script:changing=$true
    $script:deliveryEdit=$null;$script:loadedSource=$null;$script:evaluation=$null;$script:published=@()
    $script:documents.Clear();$script:ratings.Clear();$script:ranks.Clear();$ui.ResultNumbersBox.Clear();$ui.PublishedGrid.ItemsSource=@()
    $script:editDirty=$false;$script:evalDirty=$false;$script:changing=$false
    $script:selectedCase=$Case
    $script:lastCase=$Case
    $ui.CaseTitle.Text=$Case.applicationNumber
    if ($ui.MainTabs.SelectedItem -eq $ui.TargetSelectionTab) {$ui.TargetTabs.SelectedIndex=1}
    $cpc=[string](Get-UiValue $Case 'cpc');if (-not $cpc) {$cpc='미입력'}
    $affiliation=@((Get-UiValue $Case 'examinerBureau'),(Get-UiValue $Case 'examinerDivision')) | Where-Object {$_}
    $ui.CaseMetadataText.Text=(@("CPC: $cpc",($affiliation -join ' / ')) | Where-Object {$_}) -join ' · '
    $ui.CaseMetadataText.ToolTip=$ui.CaseMetadataText.Text
    $ui.CaseSubtitle.Text="$($Case.examinerName) · $($Case.groupId) · $($Case.caseId)"
    Invoke-UiWork -Work {
        param($ctx,$case)
        $evaluation=$null
        if ($case.userId -eq $ctx.User.userId -or 'Admin' -in $ctx.User.roles) { $evaluation=Get-Evaluation -Context $ctx -Case $case }
        [pscustomobject]@{deliveries=@(Get-Deliveries -Context $ctx -Case $case);evaluation=$evaluation;progress=(Get-CaseStatus -Context $ctx -Case $case)}
    } -Arguments @($context,$Case) -Completed {
        param($result)
        $snapshot=$result[0]
        $script:changing=$true
        $script:deliveryEdit=$null; $script:loadedSource=$null; $script:documents.Clear(); $ui.EditVersionText.Text='편집 자료를 불러오세요'
        $script:published=@($snapshot.deliveries); $script:evaluation=$snapshot.evaluation
        $editable=@($script:allSources | Where-Object {
            $source=$_
            $source.sourceId -in $script:loadedPeriod.sourceIds -and (
                'admin' -in $context.User.roles -or
                ($source.type -eq 'General' -and $script:selectedCase.userId -eq $context.User.userId -and 'examiner' -in $context.User.roles) -or
                ($source.type -eq 'AI' -and 'ai' -in $context.User.roles -and @($script:loadedPeriod.assignments | Where-Object {$_.userId -eq $context.User.userId -and $_.groupId -eq $script:selectedCase.groupId -and $_.sourceId -eq $source.sourceId}).Count -gt 0)
            )
        })
        $ui.SourceCombo.ItemsSource=$editable;$ui.SourceCombo.SelectedIndex=0
        $rows=@()
        foreach ($delivery in $script:published) {
            if ($delivery.data.noResults) { $rows += [pscustomobject]@{sourceName=$delivery.data.sourceName;version=$delivery.version;type='결과 없음';identifier=('결과 없음 · '+$delivery.data.noResultsReason);title=$delivery.data.noResultsReason;description=''} }
            foreach ($doc in @($delivery.data.documents)) { $rows += [pscustomobject]@{sourceName=$delivery.data.sourceName;version=$delivery.version;type=$doc.type;identifier=$doc.identifier;title=$doc.title;description=$doc.description} }
        }
        $ui.PublishedGrid.ItemsSource=$rows
        $ui.DeliverySummaryText.Text="등록된 검색 결과 $($script:published.Count)개 · 문헌 $(@($rows | Where-Object {$_.type -ne '결과 없음'}).Count)건"
        $states=@{NotStarted='미착수';Draft='작성 중';Published='등록 완료';NoResults='결과 없음';PublishedWithDraft='등록 후 수정 중'}
        $sourceStatus=@($snapshot.progress.deliveries | ForEach-Object {"$($_.sourceId): $($states[$_.status]) $($_.documentCount)건"}) -join ' / '
        $ui.CaseSubtitle.Text="$($script:selectedCase.examinerName) · $($script:selectedCase.groupId) · $sourceStatus"
        Initialize-EvaluationRows
        $script:editDirty=$false; $script:evalDirty=$false; $script:changing=$false
        Set-Status '조회 완료 · 평가와 순위는 불러온 납품 판에 연결됩니다.'
        if ($ui.MainTabs.SelectedItem -eq $ui.DeliveryTab -and $ui.SourceCombo.SelectedItem) {Load-DeliveryEditor}
    }
}
function Initialize-EvaluationRows {
    $script:ratings.Clear(); $script:ranks.Clear()
    $own=$script:selectedCase.userId -eq $context.User.userId
    $ui.SaveEvaluationButton.IsEnabled=$own
    $ui.RatingsGrid.IsReadOnly=-not $own; $ui.RankGrid.IsReadOnly=-not $own
    $stale=$false
    $old=$null; if ($script:evaluation) { $old=$script:evaluation.data }
    foreach ($delivery in $script:published) {
        if ($old) { $prop=$old.deliveryVersions.PSObject.Properties[$delivery.data.sourceId]; if (-not $prop -or [int]$prop.Value -ne $delivery.version) { $stale=$true } }
        $prior=@(); if ($old) { $prior=@($old.rankings | Where-Object {$_.sourceId -eq $delivery.data.sourceId}) }
        $choice=''; if ($delivery.data.noResults) {$choice='결과 없음'} elseif ($prior.Count) {
            switch ($prior[0].status) { 'Ranked' {$choice=[string]$prior[0].rank}; 'Outside' {$choice='순위 외'}; 'Excluded' {$choice='평가 제외'}; 'Incomparable' {$choice='비교 불가'} }
        }
        $script:ranks.Add([pscustomobject]@{sourceId=$delivery.data.sourceId;sourceName=$delivery.data.sourceName;deliveryVersion=$delivery.version;choice=$choice;noResults=[bool]$delivery.data.noResults})
        foreach ($doc in @($delivery.data.documents)) {
            $priorRating=@(); if ($old) { $priorRating=@($old.documentRatings | Where-Object {$_.sourceId -eq $delivery.data.sourceId -and $_.documentId -eq $doc.documentId}) }
            $score='';$cited='';$note=''
            if ($priorRating.Count) { $rating=$priorRating[0];if ($null -ne $rating.score) {$score=[string]$rating.score};if ($null -ne $rating.cited) {$cited=if ($rating.cited) {'O'} else {'X'}};$note=$rating.note }
            $script:ratings.Add([pscustomobject]@{sourceId=$delivery.data.sourceId;sourceName=$delivery.data.sourceName;documentId=$doc.documentId;identifier=$doc.identifier;title=$doc.title;scoreText=$score;citedText=$cited;note=$note})
        }
    }
    $ui.EvaluationVersionText.Text=if ($script:evaluation) {"평가 $($script:evaluation.version)판"} else {'평가 저장 전'}
    $ui.EvaluationStateText.Text=if ($stale) {'재확인 필요: 납품 판이 변경되었습니다. 문헌 평가와 전체 순위를 확인한 뒤 저장하세요.'} elseif ($own) {'점수와 인용 여부의 공란은 미평가입니다. 0점과 X는 실제 평가값으로 저장합니다.'} else {'본인 출원의 평가만 입력할 수 있습니다.'}
}
function Load-DeliveryEditor {
    $case=Require-Case
    if (-not $ui.SourceCombo.SelectedItem) {throw '조사 방식을 선택하세요.'}
    $source=[string]$ui.SourceCombo.SelectedItem.sourceId
    Set-Status '편집할 최신 버전 조회 중…'
    Invoke-UiWork -Work {param($ctx,$case,$source) Get-DeliveryForEdit -Context $ctx -Case $case -SourceId $source} -Arguments @($context,$case,$source) -Completed {
        param($result)
        $script:changing=$true
        $script:deliveryEdit=if ($result.Count) {$result[0]} else {$null}
        $script:loadedSource=[string]$ui.SourceCombo.SelectedItem.sourceId
        $script:documents.Clear()
        $ui.ResultNumbersBox.Clear()
        if ($script:deliveryEdit) {
            foreach ($doc in @($script:deliveryEdit.data.documents)) {
                $script:documents.Add([pscustomobject]@{documentId=$doc.documentId;type=(Get-UiValue $doc 'type' '특허');identifier=$doc.identifier;title=(Get-UiValue $doc 'title');description=(Get-UiValue $doc 'description');attachments=@($doc.attachments);attachmentCount=@($doc.attachments).Count})
            }
            $ui.NoResultsCheck.IsChecked=[bool]$script:deliveryEdit.data.noResults; $ui.NoResultsReasonBox.Text=[string]$script:deliveryEdit.data.noResultsReason
        } else {$ui.NoResultsCheck.IsChecked=$false;$ui.NoResultsReasonBox.Text=''}
        $ui.DeliveryReasonBox.Text=''
        $ui.EditVersionText.Text=if ($script:deliveryEdit) {"$($script:deliveryEdit.version)판 · $($script:deliveryEdit.data.status)"} else {'최초 저장'}
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
    foreach ($number in $numbers) {$script:documents.Add([pscustomobject]@{documentId=[guid]::NewGuid().ToString('N');type='특허';identifier=$number;title='';description='';attachments=@();attachmentCount=0})}
    $ui.ResultNumbersBox.Clear();$script:editDirty=$true
}
function Save-DeliveryEditor([string]$Status) {
    $case=Require-Case
    if (-not $script:loadedSource -or $script:loadedSource -ne $ui.SourceCombo.SelectedItem.sourceId) {throw '조사 방식의 편집 자료를 먼저 불러오세요.'}
    Add-ResultNumberRows
    [void]$ui.DocumentsGrid.CommitEdit([Windows.Controls.DataGridEditingUnit]::Cell,$true);[void]$ui.DocumentsGrid.CommitEdit([Windows.Controls.DataGridEditingUnit]::Row,$true)
    $docs=@($script:documents | ForEach-Object {[pscustomobject]@{documentId=$_.documentId;type=$_.type;identifier=$_.identifier;title=$_.title;description=$_.description;attachments=@($_.attachments)}})
    $version=if ($script:deliveryEdit) {[int]$script:deliveryEdit.version} else {0}
    $payload=@{case=$case.caseId;source=$script:loadedSource;expected=$version;status=$Status;documents=$docs;noResults=[bool]$ui.NoResultsCheck.IsChecked;noResultsReason=$ui.NoResultsReasonBox.Text;reason=$ui.DeliveryReasonBox.Text}
    $request=Get-Request "delivery:$($case.caseId):$script:loadedSource" $payload
    Set-Status '납품 저장 및 버전 확인 중…'
    Invoke-UiWork -Work {
        param($ctx,$case,$payload,$request)
        Save-Delivery -Context $ctx -Case $case -SourceId $payload.source -ExpectedVersion $payload.expected -RequestId $request -Status $payload.status -Documents $payload.documents -NoResults $payload.noResults -NoResultsReason $payload.noResultsReason -Reason $payload.reason
    } -Arguments @($context,$case,$payload,$request) -Completed {
        param($result)
        $script:deliveryEdit=$result[0];$script:editDirty=$false
        $ui.EditVersionText.Text="$($script:deliveryEdit.version)판 · $($script:deliveryEdit.data.status)"
        Set-Status "저장 확인 완료 · 납품 $($script:deliveryEdit.version)판. 공개 자료·평가를 갱신하려면 새로 고침하세요."
    }
}
function Save-EvaluationEditor {
    $case=Require-Case
    foreach ($grid in @($ui.RatingsGrid,$ui.RankGrid)) {[void]$grid.CommitEdit([Windows.Controls.DataGridEditingUnit]::Cell,$true);[void]$grid.CommitEdit([Windows.Controls.DataGridEditingUnit]::Row,$true)}
    $ratingRows=@(foreach ($row in $script:ratings) {
        $score=$null;$cited=$null
        if (-not [string]::IsNullOrWhiteSpace($row.scoreText)) {if ($row.scoreText -notmatch '^[0-5]$') {throw '점수는 0~5의 정수 또는 공란이어야 합니다.'};$score=[int]$row.scoreText}
        if (-not [string]::IsNullOrWhiteSpace($row.citedText)) {if ($row.citedText.ToUpper() -notin @('O','X')) {throw '인용 여부는 O, X 또는 공란이어야 합니다.'};$cited=$row.citedText.ToUpper() -eq 'O'}
        [pscustomobject]@{sourceId=$row.sourceId;documentId=$row.documentId;score=$score;cited=$cited;note=$row.note}
    })
    $rankingRows=@(foreach ($row in $script:ranks) {
        if ([string]::IsNullOrWhiteSpace($row.choice)) {continue}
        $rank=$null
        $state=switch ($row.choice) {'1' {$rank=1;'Ranked'};'2' {$rank=2;'Ranked'};'3' {$rank=3;'Ranked'};'순위 외' {'Outside'};'평가 제외' {'Excluded'};'비교 불가' {'Incomparable'};'결과 없음' {if (-not $row.noResults) {throw '문헌이 있는 납품에 결과 없음을 지정할 수 없습니다.'};'NoResults'};'' {'Excluded'};default {throw '순위는 1, 2, 3, 순위 외, 평가 제외, 비교 불가 중 입력하세요.'}}
        [pscustomobject]@{sourceId=$row.sourceId;rank=$rank;status=$state}
    })
    $versions=@{};foreach ($delivery in $script:published) {$versions[$delivery.data.sourceId]=[int]$delivery.version}
    $expected=if ($script:evaluation) {[int]$script:evaluation.version} else {0}
    $payload=@{expected=$expected;deliveryVersions=$versions;documentRatings=$ratingRows;rankings=$rankingRows;reason=$ui.EvaluationReasonBox.Text}
    $request=Get-Request "evaluation:$($case.caseId)" $payload
    Set-Status '평가 저장 및 납품 판 확인 중…'
    Invoke-UiWork -Work {param($ctx,$case,$payload,$request) Save-Evaluation -Context $ctx -Case $case -ExpectedVersion $payload.expected -RequestId $request -DeliveryVersions $payload.deliveryVersions -DocumentRatings $payload.documentRatings -Rankings $payload.rankings -Reason $payload.reason} -Arguments @($context,$case,$payload,$request) -Completed {
        param($result)
        $script:evaluation=$result[0];$script:evalDirty=$false
        $ui.EvaluationVersionText.Text="평가 $($script:evaluation.version)판"
        $ui.EvaluationStateText.Text='불러온 납품 판에 대한 평가를 저장했습니다.'
        Set-Status '평가 저장 확인 완료.'
    }
}
function Export-Current([bool]$IncludeAttachments,[string]$CaseId='') {
    $period=Get-SelectedPeriod;$scope=[string]$ui.ScopeCombo.SelectedItem.id
    $folder=Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'AI검색평가'
    Set-Status '고정된 납품 판으로 다운로드 파일 생성 중…'
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

function Set-SessionControls {
    $script:changing=$true
    Set-Identity
    $ui.ModeText.Text=if ($Demo) {'로컬 시연 · 사용자 선택 로그인'} else {"공유 저장소 · $DataRoot"}
    $ui.AdminTab.Visibility=if ('Admin' -in $context.User.roles) {'Visible'} else {'Collapsed'}
    $scopes=@([pscustomobject]@{id='Mine';name='내 대상건 · 검색 결과 · 평가'})
    if ('AI' -in $context.User.roles) {$scopes += [pscustomobject]@{id='Assigned';name='담당 그룹 대상건'}}
    if ('Admin' -in $context.User.roles) {$scopes += [pscustomobject]@{id='All';name='전체 현황 · 관리자'}}
    $ui.ScopeCombo.ItemsSource=$scopes;$ui.ScopeCombo.SelectedIndex=0
    $ui.PeriodCombo.ItemsSource=@(Get-Periods -Context $context);$ui.PeriodCombo.SelectedIndex=0
    $ui.AdminKindCombo.SelectedIndex=0
    $script:adminLoadedKind=$null;$script:adminLoadedId=$null;$script:adminIsNewUser=$false;$script:adminVersion=0
    $ui.AdminIdBox.Clear();$ui.AdminJsonBox.Text='{}';Set-UserForm $null
    Refresh-AdminUserList
    Update-AdminEditor
    $ui.MainTabs.SelectedItem=$ui.TargetSelectionTab;$ui.TargetTabs.SelectedIndex=0
    $script:requests=@{};$script:adminDirty=$false;$script:changing=$false
}
Set-SessionControls
$ui.DocumentsGrid.ItemsSource=$script:documents;$ui.RatingsGrid.ItemsSource=$script:ratings;$ui.RankGrid.ItemsSource=$script:ranks
$ui.SubmissionRowsGrid.ItemsSource=$script:submissionRows
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
$ui.RefreshButton.Add_Click({Invoke-UiAction {if (Confirm-Discard) {Load-Cases}}})
$ui.PeriodCombo.Add_SelectionChanged({if (-not $script:changing -and -not $script:busy) {Invoke-UiAction {
    $canChange=Confirm-Discard
    if ($canChange -and (Has-PendingSubmissions)) {$canChange=[Windows.MessageBox]::Show($window,'등록 전 출원 입력이 있습니다. 입력내용을 유지하고 새 운영월로 이동할까요?','등록 대상 월 확인','YesNo','Question') -eq 'Yes'}
    if ($canChange) {Load-Cases} else {$script:changing=$true;$ui.PeriodCombo.SelectedItem=$script:lastPeriod;$script:changing=$false}
}}})
$ui.ScopeCombo.Add_SelectionChanged({if (-not $script:changing -and -not $script:busy) {Invoke-UiAction {if (Confirm-Discard) {Load-Cases} else {$script:changing=$true;$ui.ScopeCombo.SelectedItem=$script:lastScope;$script:changing=$false}}}})
$ui.SearchBox.Add_TextChanged({if (-not $script:changing) {Update-Filter}})
$ui.CasesGrid.Add_SelectionChanged({
    if ($script:changing -or $script:busy -or -not $ui.CasesGrid.SelectedItem) {return}
    Invoke-UiAction {if (Confirm-Discard) {Show-CaseDetail $ui.CasesGrid.SelectedItem} else {$script:changing=$true;$ui.CasesGrid.SelectedItem=$script:lastCase;$script:changing=$false}}
})
$ui.SubmitButton.Add_Click({Invoke-UiAction {
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
$ui.AddSubmissionRowButton.Add_Click({Invoke-UiAction {Add-SubmissionInputRows $true $false}})
$ui.AddPastedRowsButton.Add_Click({Invoke-UiAction {Add-SubmissionInputRows $false $true}})
$ui.RemoveSubmissionRowButton.Add_Click({if ($ui.SubmissionRowsGrid.SelectedItem) {[void]$script:submissionRows.Remove($ui.SubmissionRowsGrid.SelectedItem)}})
$ui.AddResultNumbersButton.Add_Click({Invoke-UiAction {Add-ResultNumberRows}})
$ui.CancelSubmissionButton.Add_Click({Invoke-UiAction {
    if (-not (Confirm-Discard)) {return}
    $case=Require-Case
    $reason=$ui.DeliveryReasonBox.Text
    if (-not $reason) {$reason='심사관 요청에 따른 제출 취소'}
    if ([Windows.MessageBox]::Show($window,"$($case.applicationNumber) 출원을 취소할까요?",'제출 취소','YesNo','Question') -ne 'Yes') {return}
    $request=Get-Request "cancel:$($case.caseId)" @{expected=$case.version;reason=$reason}
    Invoke-UiWork -Work {param($ctx,$case,$reason,$request) Cancel-Submission -Context $ctx -Case $case -ExpectedVersion $case.version -RequestId $request -Reason $reason} -Arguments @($context,$case,$reason,$request) -Completed {param($result) Load-Cases}
}})
$ui.LoadDeliveryButton.Add_Click({Invoke-UiAction {if (Confirm-Discard) {Load-DeliveryEditor}}})
$ui.MainTabs.Add_SelectionChanged({
    param($sender,$eventArgs)
    if ($eventArgs.Source -ne $ui.MainTabs -or $script:changing -or $script:busy) {return}
    if ($ui.MainTabs.SelectedItem -eq $ui.DeliveryTab -and $script:selectedCase -and $ui.SourceCombo.SelectedItem -and -not $script:loadedSource) {Invoke-UiAction {Load-DeliveryEditor}}
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
    $script:documents.Add([pscustomobject]@{documentId=[guid]::NewGuid().ToString('N');type='특허';identifier='';title='';description='';attachments=@();attachmentCount=0});$script:editDirty=$true
}})
$ui.RemoveDocumentButton.Add_Click({if ($ui.DocumentsGrid.SelectedItem) {[void]$script:documents.Remove($ui.DocumentsGrid.SelectedItem);$script:editDirty=$true}})
$ui.AttachButton.Add_Click({Invoke-UiAction {
    $case=Require-Case;$doc=$ui.DocumentsGrid.SelectedItem
    if (-not $doc -or -not $script:loadedSource) {throw '편집 중인 문헌을 선택하세요.'}
    $dialog=New-Object Microsoft.Win32.OpenFileDialog;$dialog.Filter='문헌 파일|*.pdf;*.txt;*.docx;*.png;*.jpg|모든 파일|*.*'
    if ($dialog.ShowDialog($window)) {
        $script:attachmentDocument=$doc
        Invoke-UiWork -Work {param($ctx,$case,$source,$path) Import-DeliveryAttachment -Context $ctx -Case $case -SourceId $source -Path $path} -Arguments @($context,$case,$script:loadedSource,$dialog.FileName) -Completed {
            param($result)
            $script:attachmentDocument.attachments=@($script:attachmentDocument.attachments)+@($result[0]);$script:attachmentDocument.attachmentCount=$script:attachmentDocument.attachments.Count
            $ui.DocumentsGrid.Items.Refresh();$script:editDirty=$true;Set-Status '첨부파일 복사 완료. 납품을 저장하면 문헌에 연결됩니다.'
        }
    }
}})
$ui.HistoryButton.Add_Click({Invoke-UiAction {
    if (-not $script:deliveryEdit) {throw '저장된 편집 자료를 먼저 불러오세요.'}
    $case=Require-Case
    Invoke-UiWork -Work {param($ctx,$case,$source) Get-DeliveryHistory -Context $ctx -Case $case -SourceId $source} -Arguments @($context,$case,$script:loadedSource) -Completed {
        param($result)
        $text=($result | ForEach-Object {"$($_.version)판 · $($_.data.status) · $($_.actorId) · $($_.writtenAt)`n$($_.reason)"}) -join "`n`n"
        [Windows.MessageBox]::Show($window,$text,'납품 버전 이력','OK','Information') | Out-Null
    }
}})
$ui.DraftButton.Add_Click({Invoke-UiAction {Save-DeliveryEditor 'Draft'}})
$ui.PublishButton.Add_Click({Invoke-UiAction {Save-DeliveryEditor 'Published'}})
$ui.SaveEvaluationButton.Add_Click({Invoke-UiAction {Save-EvaluationEditor}})
$ui.ExportButton.Add_Click({Invoke-UiAction {Export-Current $false}})
$ui.ZipButton.Add_Click({Invoke-UiAction {Export-Current $true}})
$ui.CaseExportButton.Add_Click({Invoke-UiAction {$case=Require-Case;Export-Current $true $case.caseId}})
foreach ($control in @($ui.NoResultsReasonBox,$ui.DeliveryReasonBox,$ui.ResultNumbersBox)) {$control.Add_TextChanged({if (-not $script:changing) {$script:editDirty=$true}})}
$ui.NoResultsCheck.Add_Click({if (-not $script:changing) {$script:editDirty=$true}})
$ui.DocumentsGrid.Add_CellEditEnding({if (-not $script:changing) {$script:editDirty=$true}})
$ui.RatingsGrid.Add_CellEditEnding({if (-not $script:changing) {$script:evalDirty=$true}})
$ui.RankGrid.Add_CellEditEnding({if (-not $script:changing) {$script:evalDirty=$true}})
$ui.EvaluationReasonBox.Add_TextChanged({if (-not $script:changing) {$script:evalDirty=$true}})
function Change-PeriodStatus([string]$Status) {
    if (-not (Confirm-Discard)) {return}
    $period=Get-SelectedPeriod;$reason=$ui.AdminReasonBox.Text
    if ([string]::IsNullOrWhiteSpace($reason)) {throw '월 마감·재개방 사유를 입력하세요.'}
    $payload=@{period=$period;status=$Status;version=$script:loadedPeriod.version;reason=$reason}
    $request=Get-Request "period:$period" $payload
    Invoke-UiWork -Work {param($ctx,$payload,$request) Set-PeriodStatus -Context $ctx -Period $payload.period -Status $payload.status -ExpectedVersion $payload.version -RequestId $request -Reason $payload.reason} -Arguments @($context,$payload,$request) -Completed {param($result) Load-Cases}
}
$ui.ClosePeriodButton.Add_Click({Invoke-UiAction {Change-PeriodStatus 'Closed'}})
$ui.ReopenPeriodButton.Add_Click({Invoke-UiAction {Change-PeriodStatus 'Open'}})
$ui.NewUserButton.Add_Click({Invoke-UiAction {
    if ($script:adminDirty -and -not (Confirm-Discard)) {return}
    $script:changing=$true
    $ui.AdminKindCombo.SelectedIndex=0;$ui.AdminUserCombo.SelectedIndex=-1;$ui.AdminIdBox.Clear()
    $script:adminLoadedKind='Users';$script:adminLoadedId='';$script:adminIsNewUser=$true;$script:adminVersion=0
    [void]$script:requests.Remove('master:Users:')
    $data=[pscustomobject]@{name='';bureau='';division='';roles=@('examiner');active=$true}
    $ui.AdminJsonBox.Text=ConvertTo-Json -InputObject $data -Depth 20
    Set-UserForm $data;Update-AdminEditor
    $script:adminDirty=$false;$script:changing=$false
    Set-Status '성명과 소속을 입력하고 저장하세요. 사용자 식별번호는 자동 생성됩니다.'
}})
$ui.AdminLoadButton.Add_Click({Invoke-UiAction {
    if ($script:adminDirty -and -not (Confirm-Discard)) {return}
    $kind=[string]$ui.AdminKindCombo.SelectedItem.Tag
    $id=if ($kind -eq 'Users' -and $ui.AdminUserCombo.SelectedItem) {[string]$ui.AdminUserCombo.SelectedItem.userId} else {$ui.AdminIdBox.Text.Trim()}
    if ($kind -eq 'Users' -and -not $ui.AdminUserCombo.SelectedItem) {throw '사용자를 목록에서 선택하거나 신규 사용자 버튼을 누르세요.'}
    $ui.AdminIdBox.Text=$id
    if (-not $id) {throw '불러올 ID 또는 운영월을 입력하세요.'}
    Invoke-UiWork -Work {
        param($ctx,$kind,$id)
        if ($kind -eq 'Period') {$data=Get-Period -Context $ctx -Period $id}
        else {$list=switch ($kind) {'Users' {Get-Users -Context $ctx};'Groups' {Get-Groups -Context $ctx};'Sources' {Get-Sources -Context $ctx}};$field=switch ($kind) {'Users' {'userId'};'Groups' {'groupId'};'Sources' {'sourceId'}};$data=@($list | Where-Object {$_.$field -eq $id}) | Select-Object -First 1}
        if ($data) {$data} else {[pscustomobject]@{version=0}}
    } -Arguments @($context,$kind,$id) -Completed {
        param($result)
        $data=$result[0];$script:adminVersion=[int]$data.version
        $script:adminIsNewUser=$false
        $script:adminLoadedKind=[string]$ui.AdminKindCombo.SelectedItem.Tag;$script:adminLoadedId=$ui.AdminIdBox.Text.Trim()
        $clean=[ordered]@{};foreach ($prop in $data.PSObject.Properties) {if ($prop.Name -notin @('version','path')) {$clean[$prop.Name]=$prop.Value}}
        if ($script:adminLoadedKind -eq 'Users' -and -not $clean.Contains('userId')) {
            $clean=[ordered]@{userId=$script:adminLoadedId;name='';bureau='';division='';roles=@('examiner');active=$true}
        }
        $script:changing=$true
        $ui.AdminJsonBox.Text=ConvertTo-Json -InputObject $clean -Depth 40
        if ($script:adminLoadedKind -eq 'Users') {Set-UserForm $clean}
        Update-AdminEditor
        $script:adminDirty=$false;$script:changing=$false
        Set-Status "기준정보 $script:adminVersion 판을 불러왔습니다."
    }
}})
$ui.AdminSaveButton.Add_Click({Invoke-UiAction {
    $kind=[string]$ui.AdminKindCombo.SelectedItem.Tag;$id=$ui.AdminIdBox.Text.Trim()
    if ($kind -ne $script:adminLoadedKind -or (-not $script:adminIsNewUser -and $id -ne $script:adminLoadedId)) {throw '대상을 먼저 불러오거나 신규 사용자를 선택하세요.'}
    $data=if ($kind -eq 'Users') {Read-UserForm} else {ConvertFrom-Json -InputObject $ui.AdminJsonBox.Text}
    $payload=@{kind=$kind;id=$id;data=$data;expected=$script:adminVersion;reason=$ui.AdminReasonBox.Text;newUser=[bool]$script:adminIsNewUser}
    $request=Get-Request "master:$kind`:$id" $payload
    Invoke-UiWork -Work {
        param($ctx,$payload,$request)
        if ($payload.kind -eq 'Users' -and $payload.newUser) {New-User -Context $ctx -Name $payload.data.name -Bureau $payload.data.bureau -Division $payload.data.division -Roles $payload.data.roles -Active $payload.data.active -RequestId $request -Reason $payload.reason}
        elseif ($payload.kind -eq 'Period') {Save-Period -Context $ctx -Period $payload.id -ExpectedVersion $payload.expected -RequestId $request -Data $payload.data -Reason $payload.reason}
        else {Save-MasterRecord -Context $ctx -Kind $payload.kind -Id $payload.id -ExpectedVersion $payload.expected -RequestId $request -Data $payload.data -Reason $payload.reason}
    } -Arguments @($context,$payload,$request) -Completed {
        param($result)
        $saved=$result[0];$script:adminVersion=[int]$saved.version
        $script:changing=$true
        if ($script:adminLoadedKind -eq 'Users') {
            $script:adminLoadedId=[string]$saved.userId;$script:adminIsNewUser=$false;$ui.AdminIdBox.Text=$script:adminLoadedId
            $ui.AdminJsonBox.Text=ConvertTo-Json -InputObject ($saved | Select-Object * -ExcludeProperty version,path,displayName) -Depth 40
            Refresh-AdminUserList $script:adminLoadedId;Update-AdminEditor
        }
        if ($script:adminLoadedKind -eq 'Users' -and $script:adminLoadedId -eq $context.User.userId) {$script:context=Get-Context -Root $context.Root -UserId $context.User.userId;Set-Identity}
        $script:adminDirty=$false;$script:changing=$false
        Set-Status "설정 $script:adminVersion 판 저장 완료. 새 운영월은 프로그램을 다시 실행하면 목록에 표시됩니다."
    }
}})
$ui.AdminJsonBox.Add_TextChanged({if (-not $script:changing) {$script:adminDirty=$true}})
$ui.AdminKindCombo.Add_SelectionChanged({Update-AdminEditor})
$ui.AdminIdBox.Add_TextChanged({Update-AdminEditor})
$ui.AdminUserCombo.Add_SelectionChanged({
    if ($script:changing -or -not $ui.AdminUserCombo.SelectedItem) {return}
    if ($script:adminDirty -and -not (Confirm-Discard)) {
        $script:changing=$true
        $ui.AdminUserCombo.SelectedItem=@($ui.AdminUserCombo.ItemsSource | Where-Object {$_.userId -eq $script:adminLoadedId}) | Select-Object -First 1
        $script:changing=$false;return
    }
    $script:adminIsNewUser=$false;$script:adminLoadedKind=$null;$script:adminDirty=$false
    $ui.AdminIdBox.Text=[string]$ui.AdminUserCombo.SelectedItem.userId
    Update-AdminEditor
})
foreach ($control in @($ui.UserNameBox,$ui.UserBureauBox,$ui.UserDivisionBox)) {$control.Add_TextChanged({if (-not $script:changing) {$script:adminDirty=$true}})}
foreach ($control in @($ui.UserExaminerCheck,$ui.UserAiCheck,$ui.UserAdminCheck,$ui.UserActiveCheck)) {$control.Add_Click({if (-not $script:changing) {$script:adminDirty=$true}})}
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
    $script:smokeTicks=0
    $script:smokeSelected=$false
    $script:smokeStage=0
    $script:smokeTimer.Add_Tick({
        $script:smokeTicks++
        if ($script:busy -and $script:smokeTicks -lt 30) {return}
        if (-not $script:smokeSelected -and $ui.CasesGrid.Items.Count -gt 0 -and -not $script:smokeFailure) {
            $script:smokeSelected=$true
            $ui.CasesGrid.SelectedIndex=0
            $ui.MainTabs.SelectedItem=$ui.EvaluationTab
            return
        }
        if ($ExerciseUiWrites -and $script:smokeSelected -and -not $script:smokeFailure -and $script:smokeStage -lt 9) {
            Invoke-UiAction {
                switch ($script:smokeStage) {
                    0 {
                        $script:smokeOriginalCaseId=$script:selectedCase.caseId
                        $ui.MainTabs.SelectedItem=$ui.AdminTab;$ui.AdminKindCombo.SelectedIndex=0
                        $ui.NewUserButton.RaiseEvent((New-Object Windows.RoutedEventArgs([Windows.Controls.Button]::ClickEvent)))
                    }
                    1 {
                        $ui.UserNameBox.Text='화면 등록 심사관';$ui.UserBureauBox.Text='화면 등록 심사국';$ui.UserDivisionBox.Text='화면 등록 심사과'
                        Save-TestScreenshot '-Users'
                        $ui.AdminSaveButton.RaiseEvent((New-Object Windows.RoutedEventArgs([Windows.Controls.Button]::ClickEvent)))
                    }
                    2 {
                        $ui.MainTabs.SelectedIndex=0
                        $ui.ApplicationNumberBox.Text='10-2026-0000456';$ui.CpcBox.Text='H04L 9/00';Add-SubmissionInputRows $true $false
                        $ui.SubmissionBox.Text="bad-number`tB25J 9/00";Add-SubmissionInputRows $false $true
                        Save-TestScreenshot '-Submissions'
                        $ui.SubmitButton.RaiseEvent((New-Object Windows.RoutedEventArgs([Windows.Controls.Button]::ClickEvent)))
                    }
                    3 {
                        if ($script:submissionRows.Count -ne 1 -or $script:submissionRows[0].cpc -ne 'B25J 9/00' -or -not $script:submissionRows[0].message) {throw '실패한 출원 행의 CPC와 입력내용이 보존되지 않았습니다.'}
                        $script:submissionRows.Clear()
                        $ui.MainTabs.SelectedItem=$ui.EvaluationTab
                        $ui.CasesGrid.SelectedItem=@($script:cases | Where-Object {$_.caseId -eq $script:smokeOriginalCaseId})[0]
                    }
                    4 {$ui.MainTabs.SelectedItem=$ui.DeliveryTab}
                    5 {
                        $ui.ResultNumbersBox.Text="10-2025-0000007`n10-2025-0000008"
                        Save-TestScreenshot '-Results'
                        Save-DeliveryEditor 'Draft'
                    }
                    6 {$ui.DeliveryReasonBox.Text='화면 자동 검증 재납품';Save-DeliveryEditor 'Published'}
                    7 {$script:editDirty=$false;Show-CaseDetail $script:selectedCase}
                    8 {Save-EvaluationEditor}
                }
                $script:smokeStage++
            }
            return
        }
        $script:smokeTimer.Stop()
        if ($ExerciseUiWrites) {$ui.MainTabs.SelectedItem=$ui.EvaluationTab}
        Save-TestScreenshot
        if ($SmokeReportPath) {
            $report=[pscustomobject]@{userId=$context.User.userId;caseIds=@($script:cases | ForEach-Object {$_.caseId});menus=@($ui.MainTabs.Items | ForEach-Object {$_.Header});evaluationColumns=@($ui.RatingsGrid.Columns | ForEach-Object {$_.Header});adminVisible=($ui.AdminTab.Visibility -eq 'Visible')}
            [IO.File]::WriteAllText($SmokeReportPath,(ConvertTo-Json -InputObject $report -Depth 10),(New-Object Text.UTF8Encoding($false)))
        }
        $script:submissionRows.Clear();$ui.ApplicationNumberBox.Clear();$ui.CpcBox.Clear();$ui.SubmissionBox.Clear()
        $script:editDirty=$false;$script:evalDirty=$false;$window.Close()
    });$script:smokeTimer.Start()
}
[void]$window.ShowDialog()
if ($SmokeTest -and $script:smokeFailure) {throw $script:smokeFailure}
