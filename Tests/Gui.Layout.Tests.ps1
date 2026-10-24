$ErrorActionPreference='Stop'
if ([Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {throw 'Run layout checks in STA mode.'}
Add-Type -AssemblyName PresentationFramework,PresentationCore,WindowsBase
$projectPath=Split-Path $PSScriptRoot -Parent
$appPath=Join-Path $projectPath 'App'
. (Join-Path $appPath 'Ui/Layout.ps1')
$outputPath=Join-Path $projectPath 'TestResults'
[void][IO.Directory]::CreateDirectory($outputPath)
function Read-LayoutView([string]$Name) {
    $reader=[Xml.XmlReader]::Create((Join-Path $appPath "Views/$Name.xaml"))
    try {return [Windows.Markup.XamlReader]::Load($reader)} finally {$reader.Close()}
}
function Update-LayoutTest {
    $window.UpdateLayout()
    $window.Dispatcher.Invoke([Windows.Threading.DispatcherPriority]::Background,[Action]{})
    $window.UpdateLayout()
}
function Assert-Layout([bool]$Condition,[string]$Message) {if (-not $Condition) {throw $Message}}
function Point-In($Element,$Parent) {return $Element.TranslatePoint((New-Object Windows.Point(0,0)),$Parent)}
function Assert-ButtonVisible($Button) {
    $point=Point-In $Button $window
    Assert-Layout ($Button.IsVisible -and $Button.ActualHeight -ge 30 -and $point.Y -ge 0 -and $point.Y+$Button.ActualHeight -lt $window.ActualHeight-25) "$($Button.Name) is clipped at $($window.Width)x$($window.Height)."
}
function Save-LayoutImage([string]$Name) {
    Update-LayoutTest
    $bitmap=New-Object Windows.Media.Imaging.RenderTargetBitmap ([int]$window.ActualWidth),([int]$window.ActualHeight),96,96,([Windows.Media.PixelFormats]::Pbgra32)
    $bitmap.Render($window)
    $encoder=New-Object Windows.Media.Imaging.PngBitmapEncoder
    $encoder.Frames.Add([Windows.Media.Imaging.BitmapFrame]::Create($bitmap))
    $stream=[IO.File]::Create((Join-Path $outputPath ("Gui.Layout-$Name.png")))
    try {$encoder.Save($stream)} finally {$stream.Dispose()}
}
function Send-ListWheel($List,[int]$Delta) {
    $eventArgs=New-Object Windows.Input.MouseWheelEventArgs([Windows.Input.Mouse]::PrimaryDevice,0,$Delta)
    $eventArgs.RoutedEvent=[Windows.UIElement]::PreviewMouseWheelEvent
    $List.RaiseEvent($eventArgs)
    Update-LayoutTest
    return $eventArgs.Handled
}

$window=Read-LayoutView Main
try {
    $controls=@{}
    foreach ($name in @('MainTabs','PersonalTabs','EvaluationTabs','PersonalWorkTab','EvaluationTab','TargetSelectionTab','DeliveryTab','AdminTab','CaseTitle','CaseSubtitle','IdentityText','WorkflowHintText','EvaluationStateText','RankGrid','RatingsGrid','CasesGrid','DocumentsGrid','PublishedGrid','SubmissionRowsGrid','EvaluationScroll','DeliveryScroll','PublishedScroll','CaseSidebarScroll','ExportOptions','SaveEvaluationButton','PublishButton','SubmitButton','GoEvaluationButton','AdminContentHost')) {$controls[$name]=$window.FindName($name)}
    $adminView=Read-LayoutView Admin
    $controls.AdminContentHost.Content=$adminView
    $controls.IdentityText.Text='심사관 · 심사국 / 심사과'
    $controls.CaseTitle.Text='10-2026-0000123'
    $controls.CaseSubtitle.Text='ChatGPT · 클로드 입력 완료'
    $controls.WorkflowHintText.Text='AI 검색 결과 평가 · 왼쪽 출원을 선택하면 바로 평가할 수 있습니다.'
    $controls.EvaluationStateText.Text='선행기술조사는 결과 입력 없이 품질 순위만 비교합니다. 문헌별 평가는 입력된 AI 검색 결과에만 작성하세요.'
    $rankRows=@(foreach ($name in @('선행기술조사','ChatGPT','클로드')) {[pscustomobject]@{sourceName=$name;choice='1';notDelivered=$false;noResults=$false}})
    $ratingRows=@(foreach ($source in @('ChatGPT','클로드')) {foreach ($index in 1..7) {[pscustomobject]@{sourceName=$source;identifier=('10-2025-{0:0000000}' -f $index);scoreText='';citedText='';note=''}}})
    $controls.RankGrid.ItemsSource=$rankRows
    $controls.RatingsGrid.ItemsSource=$ratingRows
    $controls.PublishedGrid.ItemsSource=$ratingRows
    $controls.DocumentsGrid.ItemsSource=@($ratingRows | Select-Object -First 7)
    $controls.CasesGrid.ItemsSource=@(1..10 | ForEach-Object {[pscustomobject]@{applicationNumber=('10-2026-{0:0000000}' -f $_);cpc='G06F 16/00';examinerName='테스트 김가람';statusLabel='입력 완료'}})
    $controls.SubmissionRowsGrid.ItemsSource=$controls.CasesGrid.ItemsSource
    foreach ($name in @('RankGrid','RatingsGrid','CasesGrid','DocumentsGrid','PublishedGrid','SubmissionRowsGrid')) {Enable-ReadableListScroll $controls[$name]}
    $window.ShowActivated=$false;$window.ShowInTaskbar=$false
    $window.Show()
    foreach ($role in @('Examiner','AiExaminer','Admin')) {
        foreach ($name in @('MainTabs','PersonalTabs','EvaluationTabs')) {
            $tabs=$controls[$name]
            if ($role -eq 'Admin' -or ($role -eq 'AiExaminer' -and $name -eq 'MainTabs')) {$tabs.ClearValue([Windows.Controls.Control]::TemplateProperty)} else {$tabs.Template=$window.Resources['ContentOnlyTabs']}
            if ($name -ne 'MainTabs') {$tabs.Margin=if ($role -eq 'Admin') {12} else {0}}
        }
        $controls.MainTabs.SelectedItem=$controls.PersonalWorkTab
        $controls.PersonalTabs.SelectedItem=$controls.EvaluationTab
        $controls.EvaluationTabs.SelectedIndex=1
        foreach ($size in @(@(1480,940),@(1366,728),@(1220,760),@(1100,640))) {
            $window.Width=$size[0];$window.Height=$size[1]
            $controls.EvaluationScroll.ScrollToTop();Update-LayoutTest
            $rank=$controls.RankGrid;$ratings=$controls.RatingsGrid;$scroll=$controls.EvaluationScroll
            Assert-Layout ($rank.ActualHeight -ge 148 -and $ratings.ActualHeight -ge 300 -and $ratings.ActualHeight -gt $rank.ActualHeight) "$role $size collapses a review list."
            Assert-Layout ($scroll.ViewportHeight -gt 60) "$role $size has no usable page viewport."
            Assert-Layout ((Point-In $ratings $scroll.Content).Y+$ratings.ActualHeight -le $scroll.Content.ActualHeight+1) "$role $size clips the last document rows outside the scroll extent."
            Assert-ButtonVisible $controls.SaveEvaluationButton
            # The basic comparison fits three full rows, not just its header or one row.
            Assert-Layout ((Get-ListInnerScroll $rank).ScrollableHeight -lt 1) 'Three comparison methods must fit without an inner scrollbar.'
            $scroll.ScrollToBottom();Update-LayoutTest
            Assert-Layout ((Point-In $ratings $scroll).Y+$ratings.ActualHeight -le $scroll.ViewportHeight+1) 'Page scrolling cannot reach the bottom of the ratings list.'
            $ratings.ScrollIntoView($ratingRows[-1]);Update-LayoutTest
            $lastRow=$ratings.ItemContainerGenerator.ContainerFromItem($ratingRows[-1])
            Assert-Layout ($null -ne $lastRow) 'The last of 14 AI documents cannot be reached.'
            $rowPoint=Point-In $lastRow $scroll
            Assert-Layout ($rowPoint.Y -ge -1 -and $rowPoint.Y+$lastRow.ActualHeight -le $scroll.ViewportHeight+1) 'Last document row is hidden after scrolling.'
            Assert-ButtonVisible $controls.SaveEvaluationButton
            if ($role -eq 'Examiner' -and $size[0] -eq 1366) {
                Save-LayoutImage 'Evaluation-1366-Bottom'
                $scroll.ScrollToTop();Update-LayoutTest
                Save-LayoutImage 'Evaluation-1366-Top'
            }
        }
        Write-Output "PASS $role review at four window sizes: 3 comparison methods, 14 AI documents, minimum list heights, reachable rows and pinned save"
    }
    # Wheel over a list must move the surrounding page when the list is clipped.
    $controls.EvaluationScroll.ScrollToTop();Update-LayoutTest
    $handled=Send-ListWheel $controls.RankGrid -120
    Assert-Layout ($handled -and $controls.EvaluationScroll.VerticalOffset -gt 0) 'Nested ratings list traps downward wheel input.'
    $controls.EvaluationScroll.ScrollToBottom();Update-LayoutTest
    $previousOffset=$controls.EvaluationScroll.VerticalOffset
    $handled=Send-ListWheel $controls.RatingsGrid 120
    Assert-Layout ($handled -and $controls.EvaluationScroll.VerticalOffset -lt $previousOffset) 'Nested ratings list traps upward wheel input.'

    $window.FindName('CaseSidebarColumn').MinWidth=375
    $window.FindName('CaseSidebarColumn').Width=390
    $controls.MainTabs.SelectedItem=$controls.DeliveryTab;Update-LayoutTest
    $columnWidth=($controls.CasesGrid.Columns | Measure-Object ActualWidth -Sum).Sum
    Assert-Layout ($controls.CasesGrid.Columns[1].ActualWidth -ge 100 -and $columnWidth -le $controls.CasesGrid.ActualWidth) "AI case list clips examiner names or the status column: examiner=$($controls.CasesGrid.Columns[1].ActualWidth), columns=$columnWidth, grid=$($controls.CasesGrid.ActualWidth)."
    Assert-Layout ($controls.DocumentsGrid.ActualHeight -ge 250) 'AI input list collapsed.'
    Assert-ButtonVisible $controls.PublishButton
    $controls.DeliveryScroll.ScrollToBottom();Update-LayoutTest
    $controls.DocumentsGrid.ScrollIntoView($controls.DocumentsGrid.Items[6]);Update-LayoutTest
    Assert-Layout ((Point-In $controls.DocumentsGrid $controls.DeliveryScroll.Content).Y+$controls.DocumentsGrid.ActualHeight -le $controls.DeliveryScroll.Content.ActualHeight+1) 'AI input list exceeds scroll extent.'
    Save-LayoutImage 'AI-Minimum'
    $controls.ExportOptions.IsExpanded=$true;Update-LayoutTest
    Assert-Layout ($controls.CasesGrid.ActualHeight -ge 250) 'Opening downloads collapses the case list.'
    $controls.CaseSidebarScroll.ScrollToBottom();Update-LayoutTest
    $zipButton=$window.FindName('ZipButton')
    Assert-Layout ((Point-In $zipButton $controls.CaseSidebarScroll).Y+$zipButton.ActualHeight -le $controls.CaseSidebarScroll.ViewportHeight+1) 'Expanded download actions cannot be reached.'

    $controls.MainTabs.SelectedItem=$controls.PersonalWorkTab
    $controls.PersonalTabs.SelectedItem=$controls.TargetSelectionTab;Update-LayoutTest
    Assert-Layout ($controls.SubmissionRowsGrid.ActualHeight -ge 250) 'Selection list collapsed.'
    Assert-ButtonVisible $controls.SubmitButton
    $controls.PersonalTabs.SelectedItem=$controls.EvaluationTab;$controls.EvaluationTabs.SelectedIndex=0;Update-LayoutTest
    Assert-Layout ($controls.PublishedGrid.ActualHeight -ge 300) 'Published result list collapsed.'
    Assert-ButtonVisible $controls.GoEvaluationButton
    $controls.MainTabs.SelectedItem=$controls.AdminTab
    $window.FindName('CaseSidebar').Visibility='Collapsed';$window.FindName('CaseSidebarColumn').MinWidth=0;$window.FindName('CaseSidebarColumn').Width=0;$window.FindName('CaseGapColumn').Width=0
    $window.FindName('CaseHeaderPanel').Visibility='Collapsed';$window.FindName('CaseHeaderRow').Height=0
    foreach ($mode in @(0,1)) {
        $adminView.FindName('AdminUserTabs').SelectedIndex=$mode
        $adminView.FindName('AdminUserSelectionPanel').Visibility=if ($mode -eq 0) {'Visible'} else {'Collapsed'}
        $adminView.FindName('AdminRosterPanel').Visibility=if ($mode -eq 0) {'Visible'} else {'Collapsed'}
        $adminView.FindName('AdminRosterColumn').Width=if ($mode -eq 0) {470} else {0};$adminView.FindName('AdminRosterGap').Width=if ($mode -eq 0) {18} else {0}
        $adminView.FindName('AdminSaveButton').Content=if ($mode -eq 0) {'변경사항 저장'} else {'신규 사용자 등록'}
        $adminView.FindName('ViewRegisteredUserButton').Visibility=if ($mode -eq 1) {'Visible'} else {'Collapsed'}
        $adminScroll=$adminView.FindName('AdminFormScroll')
        $adminScroll.ScrollToBottom();Update-LayoutTest
        Save-LayoutImage "UserMode-$mode-Minimum"
        foreach ($tab in $adminView.FindName('AdminUserTabs').Items) {
            $point=Point-In $tab $adminView
            Assert-Layout ($tab.IsVisible -and $tab.ActualWidth -ge 120 -and $point.X -ge 0 -and $point.X+$tab.ActualWidth -le $adminView.ActualWidth) 'User registration/management tabs are clipped at the minimum window size.'
        }
        $active=$adminView.FindName('UserActiveCheck');$point=Point-In $active $adminScroll
        Assert-Layout ($adminScroll.ViewportHeight -ge 100 -and $point.Y -ge 0 -and $point.Y+$active.ActualHeight -le $adminScroll.ViewportHeight+1) "User role/activation fields cannot be reached by scrolling: mode=$mode viewport=$($adminScroll.ViewportHeight) activeTop=$($point.Y) activeHeight=$($active.ActualHeight) offset=$($adminScroll.VerticalOffset)."
        Assert-ButtonVisible ($adminView.FindName('AdminSaveButton'))
        if ($mode -eq 0) {Assert-Layout ($adminView.FindName('AdminUsersGrid').ActualHeight -ge 250 -and $adminView.FindName('AdminFormScroll').ActualWidth -ge 450) 'User roster or details collapse at the minimum size.'}
        else {Assert-ButtonVisible ($adminView.FindName('ViewRegisteredUserButton'))}
    }
    $adminView.FindName('AdminUserTabs').Visibility='Collapsed'
    $adminView.FindName('AdminUserSelectionPanel').Visibility='Collapsed'
    $adminView.FindName('UserFieldsPanel').Visibility='Collapsed'
    $adminView.FindName('PeriodFieldsPanel').Visibility='Visible';Update-LayoutTest
    Assert-Layout ($adminView.FindName('PeriodMembersGrid').ActualHeight -ge 250) 'Administrator participant list collapsed.'
    Assert-ButtonVisible ($adminView.FindName('AdminSaveButton'))
    $statsView=Read-LayoutView Statistics
    $window.FindName('StatisticsHost').Content=$statsView
    $controls.MainTabs.SelectedItem=$window.FindName('StatisticsTab')
    $window.FindName('CaseSidebar').Visibility='Collapsed';$window.FindName('CaseSidebarColumn').MinWidth=0;$window.FindName('CaseSidebarColumn').Width=0;$window.FindName('CaseGapColumn').Width=0
    $window.FindName('CaseHeaderPanel').Visibility='Collapsed';$window.FindName('CaseHeaderRow').Height=0
    $statsView.FindName('StatsScopeText').Text='2026-10 · 운영 중 · 전체 그룹 · 전체 계정'
    $statsView.FindName('StatsUpdatedText').Text='집계 2026-10-24 15:00:00 · 취소 출원 제외 · 최신 상태는 상단 새로 고침'
    $statsView.FindName('StatsPendingText').Text='미선정 3명 · 입력 대기 12건 · 미평가 8건 · 평가 중 2건 · 재확인 필요 1건'
    $statsScroll=$statsView.FindName('StatsScroll')
    $statsRatings=$statsView.FindName('StatsRatingsGrid')
    $statsRatings.ItemsSource=@(1..40 | ForEach-Object {[pscustomobject]@{applicationNumber=('10-2026-{0:0000000}' -f $_);examiner='담당 심사관';aiAuthor='AI 담당자';sourceName='ChatGPT';identifier='10-2020-1234567';citedText='O';score=5;evaluationStatus='평가 완료'}})
    $statsView.FindName('StatsRatingsCountText').Text='AI 검색 결과 40건 · 열 제목을 누르면 정렬할 수 있습니다.'
    foreach ($size in @(@(1480,940),@(1366,728),@(1100,640))) {
        $window.Width=$size[0];$window.Height=$size[1]
        foreach ($tabIndex in 0..4) {
            $statsView.FindName('StatsDetailTabs').SelectedIndex=$tabIndex
            $statsScroll.ScrollToBottom();Update-LayoutTest
            $list=$statsView.FindName(@('StatsGroupsGrid','StatsAccountsGrid','StatsCasesGrid','StatsSourcesGrid','StatsRatingsGrid')[$tabIndex])
            if ($list.ActualHeight -lt 300 -or $statsScroll.ViewportHeight -lt 200) {Save-LayoutImage 'Statistics-Failure'}
            Assert-Layout ($list.ActualHeight -ge 300 -and $statsScroll.ViewportHeight -ge 200) "Statistics list or viewport collapsed: size=$size tab=$tabIndex list=$($list.ActualHeight) viewport=$($statsScroll.ViewportHeight)."
            $point=Point-In $list $statsScroll.Content
            Assert-Layout ($point.Y+$list.ActualHeight -le $statsScroll.ExtentHeight+1) 'Statistics list is outside the scrollable body.'
            Assert-ButtonVisible ($statsView.FindName('StatsExportButton'))
            if ($tabIndex -eq 4) {
                Assert-Layout ($list.ActualHeight -ge 360 -and ($list.Columns | Measure-Object ActualWidth -Sum).Sum -le $list.ActualWidth) 'Evaluation detail columns or list height do not fit the minimum window.'
                $list.ScrollIntoView($list.Items[$list.Items.Count-1]);Update-LayoutTest
                Assert-Layout ($null -ne $list.ItemContainerGenerator.ContainerFromIndex($list.Items.Count-1)) 'The last document evaluation cannot be reached.'
                $list.ScrollIntoView($list.Items[0]);Update-LayoutTest
            }
        }
        Save-LayoutImage "Statistics-$($size[0])"
    }
    Write-Output 'PASS administrator statistics at three window sizes: readable lists and pinned export'
    Write-Output 'PASS wheel chaining, AI input, expanded downloads, selection, result lookup and administrator list layout'
} finally {$window.Close()}
