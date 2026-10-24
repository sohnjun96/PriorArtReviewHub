$script:lastExportDirectory = ''

function Show-ExportSaveDialog($Dialog) {
    return $Dialog.ShowDialog($window)
}

function Confirm-ExportBundleOverwrite([string]$Path) {
    return ([Windows.MessageBox]::Show($window,"같은 이름의 ZIP 파일이 있습니다. 함께 덮어쓸까요?`n$Path",'ZIP 덮어쓰기 확인','YesNo','Question') -eq 'Yes')
}

function Select-ExcelExportTarget([string]$FileName, [bool]$IncludeAttachments=$false) {
    $dialog = New-Object Microsoft.Win32.SaveFileDialog
    $dialog.Title = if ($IncludeAttachments) {'엑셀 저장 위치 선택 · ZIP도 같은 폴더에 저장'} else {'엑셀 저장 위치 선택'}
    $dialog.Filter = 'Excel 통합 문서 (*.xlsx)|*.xlsx'
    $dialog.DefaultExt = '.xlsx'
    $dialog.AddExtension = $true
    $dialog.OverwritePrompt = $true
    $dialog.CheckPathExists = $true
    $dialog.FileName = [regex]::Replace($FileName, '[\x00-\x1F<>:"/\\|?*]', '_')
    $dialog.InitialDirectory = if ($script:lastExportDirectory -and [IO.Directory]::Exists($script:lastExportDirectory)) {$script:lastExportDirectory} else {[Environment]::GetFolderPath('MyDocuments')}
    if ((Show-ExportSaveDialog $dialog) -ne $true) {return $null}
    $path = [IO.Path]::GetFullPath($dialog.FileName)
    if ([IO.Path]::GetExtension($path) -ine '.xlsx') {throw '엑셀 파일명은 .xlsx로 끝나야 합니다. 다시 저장해 주세요.'}
    $overwriteBundle = $false
    if ($IncludeAttachments) {
        $bundlePath = [IO.Path]::ChangeExtension($path, '.zip')
        if ([IO.File]::Exists($bundlePath)) {
            if (-not (Confirm-ExportBundleOverwrite $bundlePath)) {return $null}
            $overwriteBundle = $true
        }
    }
    $script:lastExportDirectory = [IO.Path]::GetDirectoryName($path)
    return [pscustomobject]@{Path=$path;OverwriteWorkbook=[IO.File]::Exists($path);OverwriteBundle=$overwriteBundle}
}

function Export-Current([bool]$IncludeAttachments,[string]$CaseId='') {
    $period=Get-SelectedPeriod;$scope=Get-WorkspaceScope
    $label=if ($CaseId) {$script:selectedCase.applicationNumber} else {switch ($scope) {'Mine' {'내 대상 건'} 'Assigned' {'담당 그룹 대상 건'} 'All' {'전체 대상 건'}}}
    $target=Select-ExcelExportTarget -FileName ("AI 검색 결과_${period}_${label}.xlsx") -IncludeAttachments $IncludeAttachments
    if (-not $target) {Set-Status '엑셀 저장을 취소했습니다.';return}
    Set-Status '다운로드 파일 생성 중…'
    Invoke-UiWork -Work {
        param($ctx,$period,$scope,$target,$attachments,$caseId)
        $exportArgs=@{Context=$ctx;Period=$period;Scope=$scope;WorkbookPath=$target.Path;OverwriteWorkbook=$target.OverwriteWorkbook;OverwriteBundle=$target.OverwriteBundle;IncludeAttachments=$attachments}
        if ($caseId) {$exportArgs['CaseId']=$caseId}
        Export-ReviewSnapshot @exportArgs
    } -Arguments @($context,$period,$scope,$target,$IncludeAttachments,$CaseId) -Completed {
        param($result)
        $export=$result[0]
        $paths=@($export.workbookPath,$export.bundlePath) | Where-Object {$_}
        Set-Status "다운로드 완료 · 출원 $($export.caseCount)건 · 문헌 $($export.documentCount)건"
        [Windows.MessageBox]::Show($window,"선택한 위치에 저장했습니다.`n$($paths -join "`n")",'다운로드 완료','OK','Information') | Out-Null
    }
}
