function Initialize-StatisticsUi {
    $reader=[Xml.XmlReader]::Create((Join-Path $appDirectory 'Views/Statistics.xaml'))
    try {$script:statisticsView=[Windows.Markup.XamlReader]::Load($reader)} finally {$reader.Close()}
    $ui.StatisticsHost.Content=$script:statisticsView
    foreach ($name in @('StatsScopeText','StatsGroupCombo','StatsUserCombo','StatsResetButton','StatsExportButton','StatsUpdatedText','StatsScroll','StatsParticipantsText','StatsCasesText','StatsDeliveriesText','StatsEvaluationsText','StatsDeliveryRateText','StatsEvaluationRateText','StatsPendingText','StatsAccountText','StatsDetailTabs','StatsGroupsGrid','StatsAccountsGrid','StatsCasesGrid','StatsSourcesGrid')) {$ui[$name]=$script:statisticsView.FindName($name)}
    $ui.StatsGroupCombo.Add_SelectionChanged({if (-not $script:changing -and -not $script:busy) {Invoke-UiAction {Set-StatisticsUsers;Show-StatisticsFilter}}})
    $ui.StatsUserCombo.Add_SelectionChanged({if (-not $script:changing -and -not $script:busy) {Invoke-UiAction {Show-StatisticsFilter}}})
    $ui.StatsResetButton.Add_Click({Invoke-UiAction {$script:changing=$true;$ui.StatsGroupCombo.SelectedIndex=0;$script:changing=$false;Set-StatisticsUsers;$ui.StatsUserCombo.SelectedIndex=0;Show-StatisticsFilter;$ui.StatsDetailTabs.SelectedIndex=0}})
    $ui.StatsGroupsGrid.Add_SelectionChanged({if (-not $script:changing -and $ui.StatsGroupsGrid.SelectedItem) {$id=$ui.StatsGroupsGrid.SelectedItem.groupId;$ui.StatsGroupCombo.SelectedValue=$id;$ui.StatsDetailTabs.SelectedIndex=1;$ui.StatsScroll.ScrollToTop()}})
    $ui.StatsAccountsGrid.Add_SelectionChanged({if (-not $script:changing -and $ui.StatsAccountsGrid.SelectedItem) {$id=$ui.StatsAccountsGrid.SelectedItem.userId;$ui.StatsUserCombo.SelectedValue=$id;$ui.StatsDetailTabs.SelectedIndex=2;$ui.StatsScroll.ScrollToTop()}})
    $ui.StatsExportButton.Add_Click({Invoke-UiAction {Export-StatisticsView}})
}
function Clear-Statistics {
    $script:statisticsSnapshot=$null;$script:statisticsFilter=$null
    $before=$script:changing;$script:changing=$true
    try {
        foreach ($name in @('StatsGroupsGrid','StatsAccountsGrid','StatsCasesGrid','StatsSourcesGrid','StatsGroupCombo','StatsUserCombo')) {$ui[$name].ItemsSource=@()}
        foreach ($name in @('StatsParticipantsText','StatsCasesText','StatsDeliveriesText','StatsEvaluationsText')) {$ui[$name].Text='—'}
        foreach ($name in @('StatsDeliveryRateText','StatsEvaluationRateText','StatsAccountText','StatsPendingText')) {$ui[$name].Text=''}
        $ui.StatsExportButton.IsEnabled=$false
        $ui.StatsUpdatedText.Text='조회가 완료되면 통계를 표시합니다. 오류가 발생하면 상단 새로 고침을 눌러 다시 조회하세요.'
    } finally {$script:changing=$before}
}
function Load-Statistics {
    $period=Get-SelectedPeriod;$script:lastPeriod=$ui.PeriodCombo.SelectedItem
    $script:statisticsOldGroup=[string](Get-UiValue $ui.StatsGroupCombo.SelectedItem 'groupId');$script:statisticsOldUser=[string](Get-UiValue $ui.StatsUserCombo.SelectedItem 'userId')
    Clear-Statistics
    $ui.StatsScopeText.Text="$period 전체 통계 조회 중…"
    Set-Status '전체·그룹·계정 통계를 집계하는 중…'
    Invoke-UiWork -Work {param($ctx,$month) Get-AdminStatistics -Context $ctx -Period $month} -Arguments @($context,$period) -Completed {
        param($result)
        $script:statisticsSnapshot=$result[0]
        $before=$script:changing;$script:changing=$true
        try {
            $ui.StatsGroupCombo.ItemsSource=@([pscustomobject]@{groupId='';name='전체 그룹'})+@($script:statisticsSnapshot.groups | Sort-Object name)
            $ui.StatsGroupCombo.SelectedValue=$script:statisticsOldGroup
            if (-not $ui.StatsGroupCombo.SelectedItem) {$ui.StatsGroupCombo.SelectedIndex=0}
            Set-StatisticsUsers $script:statisticsOldUser
        } finally {$script:changing=$before}
        Show-StatisticsFilter
        Set-Status '통계 조회 완료 · 그룹 또는 계정을 선택하면 바로 좁혀 볼 수 있습니다.'
    }
}
function Set-StatisticsUsers([string]$UserId='') {
    if (-not $script:statisticsSnapshot) {return}
    $group=[string](Get-UiValue $ui.StatsGroupCombo.SelectedItem 'groupId')
    $before=$script:changing;$script:changing=$true
    try {
        $items=@(foreach ($account in @($script:statisticsSnapshot.accounts | Where-Object {-not $group -or $group -in $_.groupIds} | Sort-Object division,name)) {
            [pscustomobject]@{userId=$account.userId;label=(@($account.division,$account.name) | Where-Object {$_}) -join ' - '}
        })
        $ui.StatsUserCombo.ItemsSource=@([pscustomobject]@{userId='';label='전체 계정'})+$items
        $ui.StatsUserCombo.SelectedValue=$UserId
        if (-not $ui.StatsUserCombo.SelectedItem) {$ui.StatsUserCombo.SelectedIndex=0}
    } finally {$script:changing=$before}
}
function Show-StatisticsFilter {
    if (-not $script:statisticsSnapshot) {return}
    $group=[string](Get-UiValue $ui.StatsGroupCombo.SelectedItem 'groupId');$user=[string](Get-UiValue $ui.StatsUserCombo.SelectedItem 'userId')
    $script:statisticsFilter=Select-AdminStatistics -Snapshot $script:statisticsSnapshot -GroupId $group -UserId $user
    $view=$script:statisticsFilter;$summary=$view.summary
    $before=$script:changing;$script:changing=$true
    try {
        $ui.StatsGroupsGrid.ItemsSource=@($view.groups);$ui.StatsAccountsGrid.ItemsSource=@($view.accounts);$ui.StatsCasesGrid.ItemsSource=@($view.cases);$ui.StatsSourcesGrid.ItemsSource=@($view.sources)
        $ui.StatsParticipantsText.Text="$($summary.participantCount) / $($summary.unselectedCount)";$ui.StatsCasesText.Text=[string]$summary.caseCount
        $ui.StatsDeliveriesText.Text="$($summary.deliveredCount) / $($summary.caseCount)";$ui.StatsDeliveryRateText.Text=$summary.deliveryRate
        $ui.StatsEvaluationsText.Text="$($summary.completeCount) / $($summary.caseCount)";$ui.StatsEvaluationRateText.Text=$summary.evaluationRate
        $ui.StatsPendingText.Text="미선정 $($summary.unselectedCount)명  ·  납품 대기 $($summary.waitingCount)건  ·  미평가 $($summary.unreviewedCount)건  ·  평가 중 $($summary.partialCount)건  ·  재확인 필요 $($summary.staleCount)건"
        $state=if ($script:statisticsSnapshot.periodStatus -eq 'Closed') {'마감'} else {'운영 중'}
        $ui.StatsScopeText.Text="$($script:statisticsSnapshot.period) · $state · $($ui.StatsGroupCombo.SelectedItem.name) · $($ui.StatsUserCombo.SelectedItem.label)"
        $ui.StatsAccountText.Text=if ($user -and $view.accounts.Count) {"$($view.accounts[0].name) · $($view.accounts[0].roles) · $($view.accounts[0].activeLabel) · AI 배정: $($view.accounts[0].assignments)"} else {"조회 계정 $($summary.accountCount)명 · 계정별 현황에는 미참여·비활성 계정도 표시됩니다."}
        $ui.StatsUpdatedText.Text="집계 $($script:statisticsSnapshot.capturedAt) · 취소 출원 $($script:statisticsSnapshot.cancelledCount)건 제외(월 전체) · 최신 상태는 상단 새로 고침"
        $ui.StatsExportButton.IsEnabled=$true
    } finally {$script:changing=$before}
}
function Export-StatisticsView([string]$OutputDirectory='') {
    if (-not $script:statisticsFilter) {throw '통계를 먼저 조회하세요.'}
    $folder=if ($OutputDirectory) {$OutputDirectory} else {Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'AI검색평가'}
    $path=Join-Path $folder ('통계-'+$script:statisticsSnapshot.period+'-'+(Get-Date -Format 'yyyyMMdd-HHmmss')+'-'+[guid]::NewGuid().ToString('N').Substring(0,6)+'.xlsx')
    $view=$script:statisticsFilter
    $columns=[ordered]@{
        '그룹별 통계'=[ordered]@{groupName='그룹';participantCount='참여 심사관';unselectedCount='미선정';caseCount='출원';deliveredCount='AI 납품 완료';waitingCount='납품 대기';completeCount='평가 완료';partialCount='평가 중';unreviewedCount='미평가';staleCount='재확인';evaluationRate='평가 완료율'}
        '계정별 현황'=[ordered]@{name='이름';bureau='심사국';division='심사과';groupName='그룹';roles='역할';activeLabel='계정';selectionStatus='대상건 선정';caseCount='본인 출원';completeCount='평가 완료';pendingCount='평가 남음';staleCount='재확인';aiPublished='담당 납품 완료';aiAssigned='담당 납품 대상';assignments='AI 배정'}
        '출원별 현황'=[ordered]@{applicationNumber='출원번호';cpc='CPC';examiner='심사관';groupName='그룹';deliveryStatus='AI 납품';aiProgress='납품 방식';evaluationStatus='평가';documentCount='문헌 수'}
        '조사 방식별 통계'=[ordered]@{sourceName='조사 방식';expectedCount='대상';publishedCount='납품 완료';pendingCount='납품 대기';documentCount='문헌 수';scoreCount='점수 입력 문헌';averageScore='평균 점수';citedCount='인용 O';citationCount='O/X 입력 문헌';citationRate='인용률';firstCount='1위 횟수'}
        '조회 기준'=[ordered]@{scope='조회 범위';capturedAt='집계 시각';rule='집계 기준'}
    }
    $sheets=[ordered]@{'그룹별 통계'=@($view.groups);'계정별 현황'=@($view.accounts);'출원별 현황'=@($view.cases);'조사 방식별 통계'=@($view.sources);'조회 기준'=@([pscustomobject]@{scope=$ui.StatsScopeText.Text;capturedAt=$script:statisticsSnapshot.capturedAt;rule='취소·선정 초안 제외. 점수·인용은 최신 납품과 일치하는 입력값만, 순위는 재확인 불필요한 평가만 집계.'})}
    Invoke-UiWork -Work {param($path,$sheets,$columns) Export-Workbook -Path $path -Sheets $sheets -Columns $columns;return $path} -Arguments @($path,$sheets,$columns) -Completed {
        param($result)
        $script:statisticsExportPath=$result[-1]
        $ui.StatsUpdatedText.Text="저장 완료 · $($result[-1])";Set-Status '현재 조회 범위의 통계를 문서 폴더의 AI검색평가 폴더에 저장했습니다.'
    }
}
