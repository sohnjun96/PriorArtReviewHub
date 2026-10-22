# Administrator forms are loaded into Main.xaml's AdminContentHost.
# Identity, locking, validation and immutable revisions remain in Workflow.psm1.
function Get-AdminKind {
    if ($ui.AdminKindCombo.SelectedItem) {return [string]$ui.AdminKindCombo.SelectedItem.Tag}
    return 'Users'
}
function Copy-AdminData($Data) {
    if ($null -eq $Data) {return [pscustomobject]@{}}
    $clean=$Data | Select-Object * -ExcludeProperty version,path,displayName
    $json=ConvertTo-Json -InputObject $clean -Depth 60
    if ((Get-Command ConvertFrom-Json).Parameters.ContainsKey('DateKind')) {return ConvertFrom-Json -InputObject $json -DateKind String}
    return ConvertFrom-Json -InputObject $json
}
function Set-AdminField($Data,[string]$Name,$Value) {$Data | Add-Member -NotePropertyName $Name -NotePropertyValue $Value -Force}
function Get-AdminSelector([string]$Kind) {
    switch ($Kind) {'Users' {$ui.AdminUserCombo};'Groups' {$ui.AdminGroupCombo};'Sources' {$ui.AdminSourceCombo};'Period' {$ui.AdminPeriodCombo}}
}
function Get-AdminId($Data,[string]$Kind) {
    $field=@{Users='userId';Groups='groupId';Sources='sourceId';Period='period'}[$Kind]
    return [string](Get-UiValue $Data $field)
}
function Get-AdminInputStamp {
    $kind=Get-AdminKind
    $values=[ordered]@{kind=$kind}
    switch ($kind) {
        'Users' {
            $values.name=$ui.UserNameBox.Text;$values.bureau=$ui.UserBureauBox.Text;$values.division=$ui.UserDivisionBox.Text
            $values.examiner=[bool]$ui.UserExaminerCheck.IsChecked;$values.ai=[bool]$ui.UserAiCheck.IsChecked;$values.admin=[bool]$ui.UserAdminCheck.IsChecked;$values.active=[bool]$ui.UserActiveCheck.IsChecked
        }
        'Groups' {$values.name=$ui.GroupNameBox.Text;$values.active=[bool]$ui.GroupActiveCheck.IsChecked}
        'Sources' {$values.name=$ui.SourceNameBox.Text;$values.type=[string]$ui.SourceTypeCombo.SelectedValue;$values.active=[bool]$ui.SourceActiveCheck.IsChecked}
        'Period' {
            $values.year=$ui.PeriodYearCombo.SelectedItem;$values.month=$ui.PeriodMonthCombo.SelectedItem;$values.limit=$ui.PeriodLimitCombo.Text
            $values.submission=[string]$ui.SubmissionDeadlinePicker.SelectedDate;$values.delivery=[string]$ui.DeliveryDeadlinePicker.SelectedDate;$values.evaluation=[string]$ui.EvaluationDeadlinePicker.SelectedDate
            $values.sources=@($ui.PeriodSourcesList.SelectedItems | ForEach-Object {$_.sourceId} | Sort-Object)
            $values.members=@($script:adminMembers | ForEach-Object {[ordered]@{userId=$_.userId;groupId=$_.groupId}})
            $values.assignments=@($script:adminAssignments | ForEach-Object {[ordered]@{userId=$_.userId;groupId=$_.groupId;sourceId=$_.sourceId}})
        }
    }
    return ConvertTo-Json -InputObject $values -Depth 20 -Compress
}
function Set-AdminDirty {
    if ($script:changing -or -not $script:adminLoadedKind) {return}
    $script:adminDirty=((Get-AdminInputStamp) -cne $script:adminBaselineStamp)
    if ($script:adminDirty) {$ui.AdminSavedText.Text='저장하지 않은 변경사항';$ui.AdminSavedText.Foreground='#AD6500'}
    elseif ($ui.AdminSavedText.Text -eq '저장하지 않은 변경사항') {$ui.AdminSavedText.Text='변경사항 없음';$ui.AdminSavedText.Foreground='#63738A'}
}
function Set-AdminOptionLabel($Records,[string]$Kind) {
    foreach ($record in @($Records)) {
        $label=[string](Get-UiValue $record 'name')
        if ($Kind -eq 'Users') {$division=[string](Get-UiValue $record 'division');if (-not $division) {$division='미지정'};$label="$division - $($record.name)"}
        if ($Kind -eq 'Period') {$label="$($record.period) · $(if ($record.status -eq 'Closed') {'마감'} else {'운영 중'})"}
        elseif (-not [bool](Get-UiValue $record 'active' $true)) {$label+=' (비활성)'}
        $record | Add-Member -NotePropertyName displayName -NotePropertyValue $label -Force
    }
}
function Refresh-AdminOptions([string]$SelectedKind='',[string]$SelectedId='') {
    if ('admin' -notin $context.User.roles) {return}
    $script:adminOptions=@{
        Users=@(Get-Users -Context $context | Sort-Object division,name,userId)
        Groups=@(Get-Groups -Context $context | Sort-Object name,groupId)
        Sources=@(Get-Sources -Context $context | Sort-Object name,sourceId)
        Period=@(Get-Periods -Context $context)
    }
    $before=$script:changing;$script:changing=$true
    try {
        foreach ($kind in @('Users','Groups','Sources','Period')) {
            Set-AdminOptionLabel $script:adminOptions[$kind] $kind
            $selector=Get-AdminSelector $kind
            $oldId=Get-AdminId $selector.SelectedItem $kind
            if ($SelectedKind -eq $kind) {$oldId=$SelectedId}
            $selector.ItemsSource=@($script:adminOptions[$kind])
            $selector.SelectedItem=@($script:adminOptions[$kind] | Where-Object {(Get-AdminId $_ $kind) -eq $oldId}) | Select-Object -First 1
            if (-not $selector.SelectedItem -and $selector.Items.Count -gt 0) {$selector.SelectedIndex=0}
        }
        $ui.UserBureauBox.ItemsSource=@($script:adminOptions.Users | ForEach-Object {[string](Get-UiValue $_ 'bureau')} | Where-Object {$_} | Sort-Object -Unique)
        Update-AdminDivisionOptions
        $script:adminView.Resources['AdminExaminerOptions']=@($script:adminOptions.Users | Where-Object {'examiner' -in $_.roles})
        $script:adminView.Resources['AdminAiUserOptions']=@($script:adminOptions.Users | Where-Object {'ai' -in $_.roles})
        $script:adminView.Resources['AdminGroupOptions']=@($script:adminOptions.Groups)
        $script:adminView.Resources['AdminAiSourceOptions']=@($script:adminOptions.Sources | Where-Object {$_.type -eq 'AI'})
        $ui.PeriodSourcesList.ItemsSource=@($script:adminOptions.Sources)
        $currentYear=(Get-Date).Year
        $periodYears=@($script:adminOptions.Period | ForEach-Object {[int]$_.period.Substring(0,4)})
        $firstYear=[int](@($currentYear-10)+$periodYears | Measure-Object -Minimum).Minimum
        $lastYear=[int](@($currentYear+10)+$periodYears | Measure-Object -Maximum).Maximum
        $ui.PeriodYearCombo.ItemsSource=@($firstYear..$lastYear)
    } finally {$script:changing=$before}
}
function Refresh-AdminUserList([string]$SelectedUserId='') {Refresh-AdminOptions -SelectedKind Users -SelectedId $SelectedUserId}
function Update-AdminDivisionOptions {
    $text=$ui.UserDivisionBox.Text
    $bureau=$ui.UserBureauBox.Text.Trim()
    $values=@($script:adminOptions.Users | Where-Object {-not $bureau -or (Get-UiValue $_ 'bureau') -eq $bureau} | ForEach-Object {[string](Get-UiValue $_ 'division')} | Where-Object {$_} | Sort-Object -Unique)
    $ui.UserDivisionBox.ItemsSource=$values;$ui.UserDivisionBox.Text=$text
}
function Set-UserForm($Data) {
    $ui.UserNameBox.Text=[string](Get-UiValue $Data 'name')
    $ui.UserBureauBox.Text=[string](Get-UiValue $Data 'bureau')
    Update-AdminDivisionOptions
    $ui.UserDivisionBox.Text=[string](Get-UiValue $Data 'division')
    $roles=@(Get-UiValue $Data 'roles' @('examiner'))
    $ui.UserExaminerCheck.IsChecked='examiner' -in $roles;$ui.UserAiCheck.IsChecked='ai' -in $roles;$ui.UserAdminCheck.IsChecked='admin' -in $roles
    $ui.UserActiveCheck.IsChecked=[bool](Get-UiValue $Data 'active' $true)
}
function Set-AdminDate($Picker,$Value) {
    $Picker.SelectedDate=$null
    if ($Value) {$Picker.SelectedDate=[datetimeoffset]::Parse([string]$Value).LocalDateTime.Date}
}
function Get-AdminDate($Picker,$Original) {
    if ($null -eq $Picker.SelectedDate) {return ''}
    $selected=([datetime]$Picker.SelectedDate).Date
    if ($Original -and [datetimeoffset]::Parse([string]$Original).LocalDateTime.Date -eq $selected) {return [string]$Original}
    return ([datetimeoffset]($selected.AddDays(1).AddTicks(-1))).ToString('o')
}
function Update-AdminEditor {
    $kind=Get-AdminKind
    foreach ($item in @(@('Users','UserFieldsPanel'),@('Groups','GroupFieldsPanel'),@('Sources','SourceFieldsPanel'),@('Period','PeriodFieldsPanel'))) {
        $ui[$item[1]].Visibility=if ($kind -eq $item[0]) {'Visible'} else {'Collapsed'}
        $ui[$item[1]].IsEnabled=($kind -eq $script:adminLoadedKind)
        $selector=Get-AdminSelector $item[0];$selector.Visibility=if ($kind -eq $item[0]) {'Visible'} else {'Collapsed'}
    }
    $ui.AdminSaveButton.IsEnabled=($kind -eq $script:adminLoadedKind)
    $ui.NewUserButton.Content=@{Users='신규 사용자';Groups='신규 그룹';Sources='신규 조사 방식';Period='신규 월'}[$kind]
}
function Show-AdminRecord($Data,[string]$Kind) {
    $before=$script:changing;$script:changing=$true
    try {
        $script:adminLoadedKind=$Kind;$script:adminLoadedId=Get-AdminId $Data $Kind
        $script:adminVersion=[int](Get-UiValue $Data 'version' 0)
        $script:adminOriginal=Copy-AdminData $Data
        $script:adminIsNew=($script:adminVersion -eq 0)
        $script:adminIsNewUser=$script:adminIsNew -and $Kind -eq 'Users'
        $ui.AdminReasonBox.Clear();$ui.AdminSavedText.Text=''
        switch ($Kind) {
            'Users' {Set-UserForm $Data}
            'Groups' {$ui.GroupNameBox.Text=[string](Get-UiValue $Data 'name');$ui.GroupActiveCheck.IsChecked=[bool](Get-UiValue $Data 'active' $true)}
            'Sources' {$ui.SourceNameBox.Text=[string](Get-UiValue $Data 'name');$ui.SourceTypeCombo.SelectedValue=[string](Get-UiValue $Data 'type' 'AI');$ui.SourceActiveCheck.IsChecked=[bool](Get-UiValue $Data 'active' $true)}
            'Period' {
                $month=[string](Get-UiValue $Data 'period' (Get-Date -Format 'yyyy-MM'))
                $ui.PeriodYearCombo.SelectedItem=[int]$month.Substring(0,4);$ui.PeriodMonthCombo.SelectedItem=[int]$month.Substring(5,2)
                $ui.PeriodYearCombo.IsEnabled=$script:adminIsNew;$ui.PeriodMonthCombo.IsEnabled=$script:adminIsNew
                $ui.PeriodLimitCombo.Text=[string](Get-UiValue $Data 'submissionLimit' 0)
                Set-AdminDate $ui.SubmissionDeadlinePicker (Get-UiValue $Data 'submissionDeadline')
                Set-AdminDate $ui.DeliveryDeadlinePicker (Get-UiValue $Data 'deliveryDeadline')
                Set-AdminDate $ui.EvaluationDeadlinePicker (Get-UiValue $Data 'evaluationDeadline')
                $ui.PeriodSourcesList.UnselectAll()
                foreach ($source in @($ui.PeriodSourcesList.ItemsSource)) {if ($source.sourceId -in @(Get-UiValue $Data 'sourceIds' @())) {$ui.PeriodSourcesList.SelectedItems.Add($source) | Out-Null}}
                $script:adminMembers.Clear();foreach ($member in @(Get-UiValue $Data 'members' @())) {$script:adminMembers.Add((Copy-AdminData $member))}
                $script:adminAssignments.Clear();foreach ($assignment in @(Get-UiValue $Data 'assignments' @())) {$script:adminAssignments.Add((Copy-AdminData $assignment))}
                $state=if ((Get-UiValue $Data 'status' 'Open') -eq 'Closed') {'마감'} else {'운영 중'}
                $ui.PeriodStateText.Text="$month · $state"
                $ui.ClosePeriodButton.IsEnabled=-not $script:adminIsNew -and $state -ne '마감';$ui.ReopenPeriodButton.IsEnabled=-not $script:adminIsNew -and $state -eq '마감'
            }
        }
        Update-AdminEditor
        $label=if ($script:adminIsNew) {'새 항목 · 입력한 내용을 저장하세요.'} else {"선택 즉시 불러옴 · 현재 $script:adminVersion 판 · 수정 후 저장하세요."}
        $ui.AdminEditorStateText.Text=$label
        $script:adminBaselineStamp=Get-AdminInputStamp
        $script:adminDirty=$false
    } finally {$script:changing=$before}
}
function Select-AdminRecord {
    $kind=Get-AdminKind;$selector=Get-AdminSelector $kind
    if ($selector.SelectedItem) {Show-AdminRecord $selector.SelectedItem $kind}
    else {
        $script:adminLoadedKind=$null;$script:adminLoadedId='';$script:adminVersion=0;$script:adminDirty=$false
        Update-AdminEditor;$ui.AdminEditorStateText.Text='등록된 항목이 없습니다. 신규 버튼을 눌러 등록하세요.'
    }
}
function Restore-AdminSelection {
    $before=$script:changing;$script:changing=$true
    try {
        foreach ($item in $ui.AdminKindCombo.Items) {if ($item.Tag -eq $script:adminLoadedKind) {$ui.AdminKindCombo.SelectedItem=$item;break}}
        $selector=Get-AdminSelector $script:adminLoadedKind
        if ($selector) {$selector.SelectedItem=@($selector.ItemsSource | Where-Object {(Get-AdminId $_ $script:adminLoadedKind) -eq $script:adminLoadedId}) | Select-Object -First 1}
        Update-AdminEditor
    } finally {$script:changing=$before}
}
function New-AdminRecord {
    if ($script:adminDirty -and -not (Confirm-Discard)) {return}
    $kind=Get-AdminKind;$id=''
    switch ($kind) {
        'Users' {$data=[pscustomobject]@{name='';bureau='';division='';roles=@('examiner');active=$true}}
        'Groups' {$id='G'+[guid]::NewGuid().ToString('N');$data=[pscustomobject]@{groupId=$id;name='';active=$true}}
        'Sources' {$id='S'+[guid]::NewGuid().ToString('N');$data=[pscustomobject]@{sourceId=$id;name='';type='AI';active=$true}}
        'Period' {
            $template=$ui.AdminPeriodCombo.SelectedItem
            if (-not $template -and $script:adminOptions.Period.Count -gt 0) {$template=$script:adminOptions.Period[0]}
            $date=Get-Date
            if ($template) {$date=[datetime]::ParseExact($template.period+'-01','yyyy-MM-dd',[Globalization.CultureInfo]::InvariantCulture).AddMonths(1)}
            while (@($script:adminOptions.Period | Where-Object {$_.period -eq $date.ToString('yyyy-MM')}).Count -gt 0) {$date=$date.AddMonths(1)}
            $data=Copy-AdminData $template
            foreach ($key in @('statusChangedAt','statusReason','requestFingerprint')) {$data.PSObject.Properties.Remove($key)}
            $fields=@{period=$date.ToString('yyyy-MM');status='Open';submissionDeadline='';deliveryDeadline='';evaluationDeadline='';submissionLimit=0}
            foreach ($key in $fields.Keys) {Set-AdminField $data $key $fields[$key]}
            if (-not $template) {Set-AdminField $data 'members' @();Set-AdminField $data 'assignments' @();Set-AdminField $data 'sourceIds' @($script:adminOptions.Sources | Where-Object {$_.active} | ForEach-Object {$_.sourceId})}
            $id=$data.period
        }
    }
    $before=$script:changing;$script:changing=$true;(Get-AdminSelector $kind).SelectedIndex=-1;$script:changing=$before
    [void]$script:requests.Remove("master:$kind`:$id")
    Show-AdminRecord $data $kind
    $ui.AdminSavedText.Text='신규 항목';$ui.AdminSavedText.Foreground='#63738A'
    Set-Status '내용을 입력한 뒤 변경사항 저장을 누르세요. 식별번호는 자동으로 부여됩니다.'
}
function Read-UserForm {
    $data=Copy-AdminData $script:adminOriginal
    $roles=@();if ($ui.UserExaminerCheck.IsChecked) {$roles+='examiner'};if ($ui.UserAiCheck.IsChecked) {$roles+='ai'};if ($ui.UserAdminCheck.IsChecked) {$roles+='admin'}
    foreach ($entry in @{name=$ui.UserNameBox.Text.Trim();bureau=$ui.UserBureauBox.Text.Trim();division=$ui.UserDivisionBox.Text.Trim();roles=@($roles);active=[bool]$ui.UserActiveCheck.IsChecked}.GetEnumerator()) {Set-AdminField $data $entry.Key $entry.Value}
    return $data
}
function Read-AdminForm {
    $kind=Get-AdminKind
    if ($kind -ne $script:adminLoadedKind) {throw '수정할 대상을 먼저 선택하거나 신규 버튼을 누르세요.'}
    if ($kind -eq 'Users') {return Read-UserForm}
    $data=Copy-AdminData $script:adminOriginal
    switch ($kind) {
        'Groups' {Set-AdminField $data 'name' $ui.GroupNameBox.Text.Trim();Set-AdminField $data 'active' ([bool]$ui.GroupActiveCheck.IsChecked)}
        'Sources' {Set-AdminField $data 'name' $ui.SourceNameBox.Text.Trim();Set-AdminField $data 'type' ([string]$ui.SourceTypeCombo.SelectedValue);Set-AdminField $data 'active' ([bool]$ui.SourceActiveCheck.IsChecked)}
        'Period' {
            foreach ($grid in @($ui.PeriodMembersGrid,$ui.PeriodAssignmentsGrid)) {[void]$grid.CommitEdit([Windows.Controls.DataGridEditingUnit]::Cell,$true);[void]$grid.CommitEdit([Windows.Controls.DataGridEditingUnit]::Row,$true)}
            $limit=0;if (-not [int]::TryParse($ui.PeriodLimitCombo.Text.Trim(),[ref]$limit) -or $limit -lt 0) {throw '제출 상한은 0 이상의 정수로 입력하세요. 0은 제한 없음입니다.'}
            $month='{0:D4}-{1:D2}' -f [int]$ui.PeriodYearCombo.SelectedItem,[int]$ui.PeriodMonthCombo.SelectedItem
            if ($script:adminIsNew -and @($script:adminOptions.Period | Where-Object {$_.period -eq $month}).Count -gt 0) {throw '이미 등록된 운영월입니다. 목록에서 해당 월을 선택하세요.'}
            Set-AdminField $data 'period' $month;Set-AdminField $data 'submissionLimit' $limit
            Set-AdminField $data 'sourceIds' @($ui.PeriodSourcesList.SelectedItems | ForEach-Object {$_.sourceId})
            Set-AdminField $data 'members' @($script:adminMembers | ForEach-Object {Copy-AdminData $_})
            Set-AdminField $data 'assignments' @($script:adminAssignments | ForEach-Object {Copy-AdminData $_})
            Set-AdminField $data 'submissionDeadline' (Get-AdminDate $ui.SubmissionDeadlinePicker (Get-UiValue $script:adminOriginal 'submissionDeadline'))
            Set-AdminField $data 'deliveryDeadline' (Get-AdminDate $ui.DeliveryDeadlinePicker (Get-UiValue $script:adminOriginal 'deliveryDeadline'))
            Set-AdminField $data 'evaluationDeadline' (Get-AdminDate $ui.EvaluationDeadlinePicker (Get-UiValue $script:adminOriginal 'evaluationDeadline'))
        }
    }
    return $data
}
function Save-AdminForm {
    $kind=Get-AdminKind;$data=Read-AdminForm;$id=Get-AdminId $data $kind
    $payload=@{kind=$kind;id=$id;data=$data;expected=$script:adminVersion;reason=$ui.AdminReasonBox.Text.Trim();newUser=[bool]$script:adminIsNewUser}
    $request=Get-Request "master:$kind`:$id" $payload
    Set-Status '설정을 저장하는 중…'
    Invoke-UiWork -Work {
        param($ctx,$payload,$request)
        if ($payload.kind -eq 'Users' -and $payload.newUser) {New-User -Context $ctx -Name $payload.data.name -Bureau $payload.data.bureau -Division $payload.data.division -Roles $payload.data.roles -Active $payload.data.active -RequestId $request -Reason $payload.reason}
        elseif ($payload.kind -eq 'Period') {Save-Period -Context $ctx -Period $payload.id -ExpectedVersion $payload.expected -RequestId $request -Data $payload.data -Reason $payload.reason}
        else {Save-MasterRecord -Context $ctx -Kind $payload.kind -Id $payload.id -ExpectedVersion $payload.expected -RequestId $request -Data $payload.data -Reason $payload.reason}
    } -Arguments @($context,$payload,$request) -Completed {
        param($result)
        $saved=$result[0];$kind=$script:adminLoadedKind;$id=Get-AdminId $saved $kind
        if ($kind -eq 'Users' -and $id -eq $context.User.userId) {$script:context=Get-Context -Root $context.Root -UserId $id -AccessMode $script:launchMode;Set-Identity}
        Refresh-AdminOptions -SelectedKind $kind -SelectedId $id
        $script:allGroups=@($script:adminOptions.Groups);$script:allSources=@($script:adminOptions.Sources)
        Show-AdminRecord $saved $kind
        $ui.AdminSavedText.Text='저장 완료';$ui.AdminSavedText.Foreground='#147D57'
        if ($kind -eq 'Period') {Sync-AdminPeriodList $id;Load-Cases}
        else {Set-Status "설정 $script:adminVersion 판 저장 완료 · 목록에 즉시 반영했습니다."}
    }
}
function Sync-AdminPeriodList([string]$SelectedPeriod='') {
    if (-not $SelectedPeriod -and $ui.PeriodCombo.SelectedItem) {$SelectedPeriod=$ui.PeriodCombo.SelectedItem.period}
    $before=$script:changing;$script:changing=$true
    try {$ui.PeriodCombo.ItemsSource=@($script:adminOptions.Period);$ui.PeriodCombo.SelectedItem=@($script:adminOptions.Period | Where-Object {$_.period -eq $SelectedPeriod}) | Select-Object -First 1;if (-not $ui.PeriodCombo.SelectedItem) {$ui.PeriodCombo.SelectedIndex=0}}
    finally {$script:changing=$before}
}
function Change-PeriodStatus([string]$Status) {
    if ($script:adminDirty -and -not (Confirm-Discard)) {return}
    $month=if ((Get-AdminKind) -eq 'Period' -and -not $script:adminIsNew) {$script:adminLoadedId} else {Get-SelectedPeriod}
    $settings=@($script:adminOptions.Period | Where-Object {$_.period -eq $month}) | Select-Object -First 1
    if (-not $settings) {throw '저장된 운영월을 먼저 선택하세요.'}
    $reason=$ui.AdminReasonBox.Text.Trim();if (-not $reason) {throw '월 마감·재개방 사유를 입력하세요.'}
    $payload=@{period=$month;status=$Status;version=[int]$settings.version;reason=$reason}
    $request=Get-Request "period:$month" $payload
    Invoke-UiWork -Work {param($ctx,$payload,$request) Set-PeriodStatus -Context $ctx -Period $payload.period -Status $payload.status -ExpectedVersion $payload.version -RequestId $request -Reason $payload.reason} -Arguments @($context,$payload,$request) -Completed {
        param($result)
        $saved=$result[0];Refresh-AdminOptions -SelectedKind Period -SelectedId $saved.period
        Show-AdminRecord $saved Period;Sync-AdminPeriodList $saved.period;Load-Cases
    }
}
function Reset-AdminEditor {
    $script:adminDirty=$false;$script:adminBaselineStamp='';$script:adminVersion=0;$script:adminLoadedKind=$null;$script:adminLoadedId='';$script:adminOriginal=[pscustomobject]@{};$script:adminIsNew=$false;$script:adminIsNewUser=$false
    $script:adminOptions=@{Users=@();Groups=@();Sources=@();Period=@()}
    $before=$script:changing;$script:changing=$true
    try {
        $script:adminMembers.Clear();$script:adminAssignments.Clear();$ui.AdminReasonBox.Clear();$ui.AdminKindCombo.SelectedIndex=0
        foreach ($kind in @('Users','Groups','Sources','Period')) {(Get-AdminSelector $kind).ItemsSource=@()}
        Update-AdminEditor
        if ('admin' -in $context.User.roles) {Refresh-AdminOptions;Select-AdminRecord}
        else {$ui.AdminEditorStateText.Text='관리자용 실행 파일에서 운영 관리 기능을 사용할 수 있습니다.'}
    } finally {$script:changing=$before}
}
function Add-AdminDropColumn($Grid,[string]$Header,[string]$Field,[string]$Resource,[string]$ValuePath) {
    $column=New-Object Windows.Controls.DataGridTemplateColumn;$column.Header=$Header;$column.Width=New-Object Windows.Controls.DataGridLength(1,[Windows.Controls.DataGridLengthUnitType]::Star)
    $markup='<DataTemplate xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"><ComboBox ItemsSource="{DynamicResource '+$Resource+'}" DisplayMemberPath="displayName" SelectedValuePath="'+$ValuePath+'" SelectedValue="{Binding '+$Field+', Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}" BorderThickness="0" Background="Transparent" Padding="4" MinWidth="70" IsTextSearchEnabled="True"/></DataTemplate>'
    $column.CellTemplate=[Windows.Markup.XamlReader]::Parse($markup);$Grid.Columns.Add($column) | Out-Null
}
function Initialize-AdminUi {
    $reader=[Xml.XmlReader]::Create((Join-Path $appDirectory 'Views\Admin.xaml'))
    try {$script:adminView=[Windows.Markup.XamlReader]::Load($reader)} finally {$reader.Close()}
    $ui.AdminContentHost.Content=$script:adminView
    foreach ($name in @('AdminKindCombo','AdminUserCombo','AdminGroupCombo','AdminSourceCombo','AdminPeriodCombo','NewUserButton','AdminEditorStateText','AdminReasonBox','AdminSaveButton','AdminReloadButton','AdminSavedText','UserFieldsPanel','UserNameBox','UserBureauBox','UserDivisionBox','UserExaminerCheck','UserAiCheck','UserAdminCheck','UserActiveCheck','GroupFieldsPanel','GroupNameBox','GroupActiveCheck','SourceFieldsPanel','SourceNameBox','SourceTypeCombo','SourceActiveCheck','PeriodFieldsPanel','PeriodStateText','ClosePeriodButton','ReopenPeriodButton','PeriodYearCombo','PeriodMonthCombo','PeriodLimitCombo','SubmissionDeadlinePicker','DeliveryDeadlinePicker','EvaluationDeadlinePicker','PeriodSourcesList','PeriodMembersGrid','PeriodAssignmentsGrid','AddPeriodMemberButton','RemovePeriodMemberButton','AddPeriodAssignmentButton','RemovePeriodAssignmentButton')) {$ui[$name]=$script:adminView.FindName($name)}
    $script:adminMembers=New-Object 'System.Collections.ObjectModel.ObservableCollection[object]';$script:adminAssignments=New-Object 'System.Collections.ObjectModel.ObservableCollection[object]'
    $ui.PeriodMembersGrid.AutoGenerateColumns=$false;$ui.PeriodAssignmentsGrid.AutoGenerateColumns=$false
    $ui.PeriodMembersGrid.CanUserAddRows=$false;$ui.PeriodAssignmentsGrid.CanUserAddRows=$false
    $ui.PeriodMembersGrid.ItemsSource=$script:adminMembers;$ui.PeriodAssignmentsGrid.ItemsSource=$script:adminAssignments
    Add-AdminDropColumn $ui.PeriodMembersGrid '심사관' 'userId' 'AdminExaminerOptions' 'userId';Add-AdminDropColumn $ui.PeriodMembersGrid '그룹' 'groupId' 'AdminGroupOptions' 'groupId'
    Add-AdminDropColumn $ui.PeriodAssignmentsGrid 'AI 담당자' 'userId' 'AdminAiUserOptions' 'userId';Add-AdminDropColumn $ui.PeriodAssignmentsGrid '그룹' 'groupId' 'AdminGroupOptions' 'groupId';Add-AdminDropColumn $ui.PeriodAssignmentsGrid 'AI 조사 방식' 'sourceId' 'AdminAiSourceOptions' 'sourceId'
    $year=(Get-Date).Year;$ui.PeriodYearCombo.ItemsSource=@(($year-10)..($year+10));$ui.PeriodMonthCombo.ItemsSource=@(1..12);$ui.PeriodLimitCombo.ItemsSource=@('0','1','2','3','5','10','15','20','30','50','100')
    $ui.AdminKindCombo.Add_SelectionChanged({if ($script:changing -or $script:busy) {return};Invoke-UiAction {if ($script:adminDirty -and -not (Confirm-Discard)) {Restore-AdminSelection;return};Update-AdminEditor;Select-AdminRecord}})
    foreach ($selector in @($ui.AdminUserCombo,$ui.AdminGroupCombo,$ui.AdminSourceCombo,$ui.AdminPeriodCombo)) {$selector.Add_SelectionChanged({if ($script:changing -or $script:busy) {return};Invoke-UiAction {if ($script:adminDirty -and -not (Confirm-Discard)) {Restore-AdminSelection;return};Select-AdminRecord}})}
    $ui.NewUserButton.Add_Click({Invoke-UiAction {New-AdminRecord}});$ui.AdminSaveButton.Add_Click({Invoke-UiAction {Save-AdminForm}})
    $ui.AdminReloadButton.Add_Click({Invoke-UiAction {if ($script:adminDirty -and -not (Confirm-Discard)) {return};$kind=Get-AdminKind;$id=$script:adminLoadedId;Refresh-AdminOptions -SelectedKind $kind -SelectedId $id;Select-AdminRecord;Set-Status '운영 목록을 새로 불러왔습니다.'}})
    $ui.ClosePeriodButton.Add_Click({Invoke-UiAction {Change-PeriodStatus 'Closed'}});$ui.ReopenPeriodButton.Add_Click({Invoke-UiAction {Change-PeriodStatus 'Open'}})
    foreach ($box in @($ui.UserNameBox,$ui.GroupNameBox,$ui.SourceNameBox)) {$box.Add_TextChanged({Set-AdminDirty})}
    foreach ($combo in @($ui.UserBureauBox,$ui.UserDivisionBox,$ui.PeriodLimitCombo)) {$combo.AddHandler([Windows.Controls.Primitives.TextBoxBase]::TextChangedEvent,[Windows.Controls.TextChangedEventHandler]{Set-AdminDirty})}
    $ui.UserBureauBox.Add_SelectionChanged({if (-not $script:changing) {[void]$ui.UserBureauBox.Dispatcher.BeginInvoke([Action]{$before=$script:changing;$script:changing=$true;try {Update-AdminDivisionOptions} finally {$script:changing=$before};Set-AdminDirty})}})
    foreach ($combo in @($ui.SourceTypeCombo,$ui.PeriodYearCombo,$ui.PeriodMonthCombo)) {$combo.Add_SelectionChanged({Set-AdminDirty})}
    foreach ($check in @($ui.UserExaminerCheck,$ui.UserAiCheck,$ui.UserAdminCheck,$ui.UserActiveCheck,$ui.GroupActiveCheck,$ui.SourceActiveCheck)) {$check.Add_Click({Set-AdminDirty})}
    foreach ($picker in @($ui.SubmissionDeadlinePicker,$ui.DeliveryDeadlinePicker,$ui.EvaluationDeadlinePicker)) {$picker.Add_SelectedDateChanged({Set-AdminDirty})}
    $ui.PeriodSourcesList.Add_SelectionChanged({Set-AdminDirty})
    foreach ($grid in @($ui.PeriodMembersGrid,$ui.PeriodAssignmentsGrid)) {
        $grid.AddHandler([Windows.Controls.Primitives.Selector]::SelectionChangedEvent,[Windows.Controls.SelectionChangedEventHandler]{param($sender,$eventArgs) if (-not $script:changing -and $eventArgs.OriginalSource -is [Windows.Controls.ComboBox]) {[void]$sender.Dispatcher.BeginInvoke([Action]{Set-AdminDirty})}},$true)
        $grid.Add_CellEditEnding({Set-AdminDirty})
    }
    $ui.AddPeriodMemberButton.Add_Click({Invoke-UiAction {$script:adminMembers.Add([pscustomobject]@{userId='';groupId=''});Set-AdminDirty}})
    $ui.RemovePeriodMemberButton.Add_Click({if ($ui.PeriodMembersGrid.SelectedItem) {[void]$script:adminMembers.Remove($ui.PeriodMembersGrid.SelectedItem);Set-AdminDirty}})
    $ui.AddPeriodAssignmentButton.Add_Click({Invoke-UiAction {$script:adminAssignments.Add([pscustomobject]@{userId='';groupId='';sourceId=''});Set-AdminDirty}})
    $ui.RemovePeriodAssignmentButton.Add_Click({if ($ui.PeriodAssignmentsGrid.SelectedItem) {[void]$script:adminAssignments.Remove($ui.PeriodAssignmentsGrid.SelectedItem);Set-AdminDirty}})
}
