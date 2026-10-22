function Get-ListInnerScroll($Visual) {
    if ($Visual -is [Windows.Controls.ScrollViewer]) {return $Visual}
    for ($index=0;$index -lt [Windows.Media.VisualTreeHelper]::GetChildrenCount($Visual);$index++) {
        $found=Get-ListInnerScroll ([Windows.Media.VisualTreeHelper]::GetChild($Visual,$index))
        if ($found) {return $found}
    }
}

function Enable-ReadableListScroll($List) {
    # Nested lists should not trap the wheel while the page hides part of the list.
    $List.Add_PreviewMouseWheel({
        param($sender,$eventArgs)
        if ([Windows.Input.Keyboard]::Modifiers -ne [Windows.Input.ModifierKeys]::None) {return}
        $outer=[Windows.Media.VisualTreeHelper]::GetParent($sender)
        while ($outer -and $outer -isnot [Windows.Controls.ScrollViewer]) {$outer=[Windows.Media.VisualTreeHelper]::GetParent($outer)}
        if (-not $outer -or $outer.ScrollableHeight -le 0) {return}
        $inner=Get-ListInnerScroll $sender
        if (-not $inner) {return}
        $point=$sender.TranslatePoint((New-Object Windows.Point(0,0)),$outer)
        $down=$eventArgs.Delta -lt 0
        $partlyHidden=if ($down) {$point.Y+$sender.ActualHeight -gt $outer.ViewportHeight+1} else {$point.Y -lt -1}
        $atListEnd=if ($down) {$inner.VerticalOffset -ge $inner.ScrollableHeight} else {$inner.VerticalOffset -le 0}
        $canMovePage=if ($down) {$outer.VerticalOffset -lt $outer.ScrollableHeight} else {$outer.VerticalOffset -gt 0}
        if ($canMovePage -and ($partlyHidden -or $atListEnd)) {
            $outer.ScrollToVerticalOffset($outer.VerticalOffset-($eventArgs.Delta / 3))
            $eventArgs.Handled=$true
        }
    })
}
