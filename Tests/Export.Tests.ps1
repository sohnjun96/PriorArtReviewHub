Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot '..\App\Modules\Export.psm1') -Force

function Assert-ExportTest {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw "FAIL: $Message" }
}

function Assert-ExportThrows {
    param([scriptblock]$Action, [string]$Message)
    $thrown = $false
    try { & $Action | Out-Null } catch { $thrown = $true }
    Assert-ExportTest $thrown $Message
}

function Read-ExportZipEntry {
    param([System.IO.Compression.ZipArchive]$Archive, [string]$Name)
    $entry = $Archive.GetEntry($Name)
    if ($null -eq $entry) { throw "Missing ZIP entry: $Name" }
    $reader = New-Object IO.StreamReader($entry.Open(), [Text.Encoding]::UTF8)
    try { return $reader.ReadToEnd() } finally { $reader.Dispose() }
}

function Open-ExportTestZip {
    param([string]$Path)
    $stream = [IO.File]::OpenRead($Path)
    return (New-Object IO.Compression.ZipArchive($stream, [IO.Compression.ZipArchiveMode]::Read, $false))
}

function Assert-ExportCaseFields {
    param([System.IO.Compression.ZipArchive]$Archive, [string]$Bureau, [string]$Division, [string]$Cpc)
    $expected = [ordered]@{ '심사국' = $Bureau; '심사과' = $Division; 'CPC' = $Cpc }
    foreach ($sheetIndex in 1..4) {
        [xml]$sheet = Read-ExportZipEntry $Archive ('xl/worksheets/sheet{0}.xml' -f $sheetIndex)
        $sheetRows = @($sheet.GetElementsByTagName('row'))
        Assert-ExportTest ($sheetRows.Count -gt 1) "Snapshot sheet $sheetIndex has data rows to verify."
        Assert-ExportTest (@($sheetRows[0].GetElementsByTagName('c') | Where-Object {$_.InnerText -match '버전'}).Count -eq 0) 'Export does not expose revision columns.'
        $columnReferences = @{}
        foreach ($cell in $sheetRows[0].GetElementsByTagName('c')) {
            if ($expected.Contains($cell.InnerText)) { $columnReferences[$cell.InnerText] = $cell.GetAttribute('r') -replace '\d+$', '' }
        }
        foreach ($label in $expected.Keys) { Assert-ExportTest ($columnReferences.ContainsKey($label)) "Snapshot sheet $sheetIndex contains $label header." }
        foreach ($row in @($sheetRows | Select-Object -Skip 1)) {
            $cells = @{}
            foreach ($cell in $row.GetElementsByTagName('c')) { $cells[$cell.GetAttribute('r')] = $cell }
            foreach ($label in $expected.Keys) {
                $reference = $columnReferences[$label] + $row.GetAttribute('r')
                $cell = $cells[$reference]
                Assert-ExportTest ($null -ne $cell -and $cell.InnerText -eq $expected[$label]) "Snapshot sheet $sheetIndex row $($row.GetAttribute('r')) preserves $label from the submitted case."
                Assert-ExportTest ($cell.GetAttribute('t') -eq 'inlineStr') "Snapshot $label remains explicit text."
            }
        }
    }
}

$tempRoot = Join-Path ([IO.Path]::GetTempPath()) ('AiReviewExportTests-' + [guid]::NewGuid().ToString('N'))
[void][IO.Directory]::CreateDirectory($tempRoot)
try {
    $unicodeText = '한글 & <문헌> ' + [char]1 + [char]::ConvertFromUtf32(0x1F600) + [char]0xD800
    $rows = @([pscustomobject][ordered]@{
        applicationNumber = '001234567890'; title = '=HYPERLINK("https://invalid.example","문헌")'; unicode = $unicodeText
        score = 0; cited = $false; empty = $null; negative = -2; exactText = '  +계산하지 않음  '
    })
    $workbookPath = Join-Path $tempRoot '문헌목록.xlsx'
    $sheets = [ordered]@{ 'A/B' = $rows; 'A\B' = @() }
    $columns = @{ 'A/B' = [ordered]@{ applicationNumber = '출원번호'; title = '명칭'; unicode = '설명'; score = '점수'; cited = '인용'; empty = '미평가'; negative = '숫자'; exactText = '그대로' } }
    $writtenPath = Export-Workbook -Path $workbookPath -Sheets $sheets -Columns $columns
    Assert-ExportTest ($writtenPath -eq $workbookPath) 'Workbook returns the absolute output path.'
    Assert-ExportTest ([IO.File]::Exists($workbookPath)) 'Workbook was created.'
    Assert-ExportThrows { Export-Workbook -Path $workbookPath -Sheets $sheets } 'Existing workbook must not be overwritten.'
    Assert-ExportThrows { Export-Workbook -Path (Join-Path $tempRoot 'oversize.xlsx') -Sheets ([ordered]@{ '내용' = @([pscustomobject]@{ value = ('가' * 32768) }) }) } 'Oversized Excel cells fail clearly instead of creating a corrupt workbook.'
    Assert-ExportTest (-not [IO.File]::Exists((Join-Path $tempRoot 'oversize.xlsx'))) 'Failed workbook was never published.'
    $zip = Open-ExportTestZip $workbookPath
    try {
        foreach ($required in @('[Content_Types].xml', '_rels/.rels', 'xl/workbook.xml', 'xl/_rels/workbook.xml.rels', 'xl/styles.xml', 'xl/worksheets/sheet1.xml', 'xl/worksheets/sheet2.xml')) {
            Assert-ExportTest ($null -ne $zip.GetEntry($required)) "Workbook package contains $required."
        }
        foreach ($entry in $zip.Entries) {
            if ($entry.FullName.EndsWith('.xml') -or $entry.FullName.EndsWith('.rels')) {
                [xml]$validXml = Read-ExportZipEntry $zip $entry.FullName
                Assert-ExportTest ($null -ne $validXml.DocumentElement) "Valid XML: $($entry.FullName)"
            }
        }
        [xml]$workbook = Read-ExportZipEntry $zip 'xl/workbook.xml'
        $sheetNames = @($workbook.GetElementsByTagName('sheet') | ForEach-Object { $_.GetAttribute('name') })
        Assert-ExportTest ($sheetNames[0] -eq 'A B' -and $sheetNames[1] -eq 'A B (2)') 'Invalid and duplicate sheet names are made legal.'
        [xml]$sheet = Read-ExportZipEntry $zip 'xl/worksheets/sheet1.xml'
        $cells = @{}
        foreach ($cell in $sheet.GetElementsByTagName('c')) { $cells[$cell.GetAttribute('r')] = $cell }
        Assert-ExportTest ($cells['A2'].GetAttribute('t') -eq 'inlineStr' -and $cells['A2'].InnerText -eq '001234567890') 'Leading zeros remain literal strings.'
        Assert-ExportTest ($cells['B2'].GetAttribute('t') -eq 'inlineStr' -and $cells['B2'].InnerText.StartsWith('=')) 'Formula-like user text remains inline text.'
        Assert-ExportTest ($sheet.GetElementsByTagName('f').Count -eq 0) 'There are no formula elements.'
        Assert-ExportTest ($cells['C2'].InnerText -eq ('한글 & <문헌> ' + [char]::ConvertFromUtf32(0x1F600))) 'Unicode survives and invalid XML controls/surrogates are removed.'
        Assert-ExportTest ($cells['D2'].InnerText -eq '0' -and $cells['D2'].GetAttribute('t') -eq '') 'A real zero is numeric.'
        Assert-ExportTest ($cells['E2'].InnerText -eq '아니오') 'Boolean false is exported explicitly.'
        Assert-ExportTest ($cells['F2'].InnerText -eq '' -and $cells['F2'].ChildNodes.Count -eq 0) 'Null is an empty cell, not zero or false.'
        Assert-ExportTest ($cells['G2'].InnerText -eq '-2') 'Negative numeric values are numeric.'
        Assert-ExportTest ($cells['H2'].InnerText -eq '  +계산하지 않음  ') 'Whitespace and formula prefixes remain unchanged.'
        Assert-ExportTest ($sheet.GetElementsByTagName('pane')[0].GetAttribute('state') -eq 'frozen') 'Header row is frozen.'
        Assert-ExportTest ($sheet.GetElementsByTagName('autoFilter')[0].GetAttribute('ref') -eq 'A1:H2') 'AutoFilter covers the exported range.'
    } finally { $zip.Dispose() }

    $attachmentPath = Join-Path $tempRoot '첨부문헌.txt'
    [IO.File]::WriteAllText($attachmentPath, '첨부 내용', (New-Object Text.UTF8Encoding($false)))
    $bundlePath = Export-CaseBundle -Path (Join-Path $tempRoot '묶음.zip') -WorkbookPath $workbookPath -Attachments @(@{ path = $attachmentPath; entryName = '첨부/문헌.txt' })
    $zip = Open-ExportTestZip $bundlePath
    try {
        Assert-ExportTest ($zip.Entries.Count -eq 2) 'Bundle includes workbook and attachment.'
        Assert-ExportTest ((Read-ExportZipEntry $zip '첨부/문헌.txt') -eq '첨부 내용') 'Attachment is copied without changing its contents.'
    } finally { $zip.Dispose() }
    foreach ($unsafeName in @('../outside.txt', 'C:\outside.txt', '/absolute.txt', 'a/./b.txt', 'a//b.txt', '.. /outside.txt', 'a/NUL.txt', 'a/file?.txt')) {
        Assert-ExportThrows { Export-CaseBundle -Path (Join-Path $tempRoot 'unsafe.zip') -WorkbookPath $workbookPath -Attachments @(@{ path = $attachmentPath; entryName = $unsafeName }) } "ZIP path traversal rejected: $unsafeName"
    }
    Assert-ExportThrows { Export-CaseBundle -Path (Join-Path $tempRoot 'duplicate.zip') -WorkbookPath $workbookPath -Attachments @(@{ path = $attachmentPath; entryName = 'A.txt' }, @{ path = $attachmentPath; entryName = 'a.txt' }) } 'Case-insensitive duplicate ZIP entry names are rejected.'
    Assert-ExportTest (-not [IO.File]::Exists((Join-Path $tempRoot 'unsafe.zip'))) 'Rejected archive was never published.'

    # Exercise the real storage/workflow modules, including demo role boundaries.
    Import-Module (Join-Path $PSScriptRoot '..\App\Modules\Workflow.psm1') -Force
    $storeRoot = Join-Path $tempRoot 'store'
    $adminContext = Initialize-Demo -Root $storeRoot
    $examinerMaster = @(Get-Users -Context $adminContext | Where-Object { $_.userId -eq 'U002' })[0]
    $examinerMaster | Add-Member -NotePropertyName bureau -NotePropertyValue '기계금속심사국' -Force
    $examinerMaster | Add-Member -NotePropertyName division -NotePropertyValue '로봇자동화심사과' -Force
    $null = Save-MasterRecord -Context $adminContext -Type Users -Id 'U002' -Data $examinerMaster -ExpectedVersion $examinerMaster.version -RequestId ([guid]::NewGuid().ToString('N'))
    $examinerContext = Get-Context -Root $storeRoot -UserId 'U002'
    $operatorContext = Get-Context -Root $storeRoot -UserId 'U003'
    $operationMonth = Get-Date -Format 'yyyy-MM'
    $realCase = New-Submission -Context $examinerContext -Period $operationMonth -ApplicationNumber ('10-' + (Get-Date -Format 'yyyy') + '-0000001') -Cpc 'G06F 16/00' -RequestId ([guid]::NewGuid().ToString('N'))
    $realAttachment = Import-DeliveryAttachment -Context $operatorContext -Case $realCase -SourceId 'S03' -Path $attachmentPath
    $realDelivery = Save-Delivery -Context $operatorContext -Case $realCase -SourceId 'S03' -ExpectedVersion 0 -RequestId ([guid]::NewGuid().ToString('N')) -Status Published -Documents @(
        [pscustomobject]@{ documentId = 'D001'; type = 'Patent'; identifier = '001122'; title = '=실제 납품'; description = '실제 모듈 연동'; attachments = @($realAttachment) },
        [pscustomobject]@{ documentId = 'D002'; identifier = '10-2020-0011223' },
        [pscustomobject]@{ documentId = 'D003'; identifier = '10-2021-0044556'; title = '' }
    )
    Assert-ExportTest ($realDelivery.data.documents[1].title -eq '' -and $realDelivery.data.documents[2].title -eq '') 'Identifier-only documents publish when title is omitted or empty.'
    $null = Save-Evaluation -Context $examinerContext -Case $realCase -ExpectedVersion 0 -RequestId ([guid]::NewGuid().ToString('N')) -DeliveryVersions @{ S01 = 0; S02 = 0; S03 = $realDelivery.version } -DocumentRatings @(
        [pscustomobject]@{ sourceId = 'S03'; documentId = 'D001'; score = 0; cited = $false; note = '=비고' },
        [pscustomobject]@{ sourceId = 'S03'; documentId = 'D002'; score = 5; cited = $true; note = '번호로만 등록한 문헌' }
    )
    $examinerMaster = @(Get-Users -Context $adminContext | Where-Object { $_.userId -eq 'U002' })[0]
    $examinerMaster.bureau = '변경한 심사국'
    $examinerMaster.division = '변경한 심사과'
    $null = Save-MasterRecord -Context $adminContext -Type Users -Id 'U002' -Data $examinerMaster -ExpectedVersion $examinerMaster.version -RequestId ([guid]::NewGuid().ToString('N'))
    $examinerContext = Get-Context -Root $storeRoot -UserId 'U002'
    $realOwnedExport = Export-ReviewSnapshot -Context $examinerContext -Period $operationMonth -Scope Mine -OutputDirectory $tempRoot -CaseId $realCase.caseId -IncludeAttachments
    Assert-ExportTest ($realOwnedExport.documentCount -eq 3 -and [IO.File]::Exists($realOwnedExport.bundlePath)) 'Real workflow exports legacy titled and identifier-only documents with their attachment ZIP.'
    $zip = Open-ExportTestZip $realOwnedExport.workbookPath
    try {
        Assert-ExportTest ((Read-ExportZipEntry $zip 'xl/worksheets/sheet1.xml').Contains('평가 중')) 'A partial evaluation is labelled as in progress.'
        Assert-ExportCaseFields -Archive $zip -Bureau '기계금속심사국' -Division '로봇자동화심사과' -Cpc 'G06F 16/00'
        [xml]$realDocumentSheet = Read-ExportZipEntry $zip 'xl/worksheets/sheet2.xml'
        Assert-ExportTest ($realDocumentSheet.DocumentElement.InnerText.Contains('=실제 납품')) 'Historical document titles remain available in the document list.'
        [xml]$realRatingSheet = Read-ExportZipEntry $zip 'xl/worksheets/sheet4.xml'
        $realRatingRows = @($realRatingSheet.GetElementsByTagName('row'))
        $ratingHeaders = @($realRatingRows[0].GetElementsByTagName('c') | ForEach-Object { $_.InnerText })
        $coreRatingHeaders = @('조사 방식', '문헌', '점수(0~5점)', '인용여부(O/X)', '비고')
        Assert-ExportTest (@($ratingHeaders | Where-Object { $coreRatingHeaders -contains $_ }).Count -eq 5 -and $ratingHeaders -notcontains '명칭') 'Rating export has the five requested evaluation columns and no title column.'
        $exportedRatings = @{}
        foreach ($row in @($realRatingRows | Select-Object -Skip 1)) {
            $values = @($row.GetElementsByTagName('c'))
            $ratingRow = @{}
            for ($columnIndex = 0; $columnIndex -lt $ratingHeaders.Count; $columnIndex++) { $ratingRow[$ratingHeaders[$columnIndex]] = $values[$columnIndex] }
            $exportedRatings[$ratingRow['문헌'].InnerText] = $ratingRow
        }
        Assert-ExportTest ($exportedRatings.ContainsKey('10-2020-0011223') -and $exportedRatings.ContainsKey('10-2021-0044556')) 'The document column shows the supplied numbers even without document titles.'
        Assert-ExportTest ($exportedRatings['10-2020-0011223']['점수(0~5점)'].InnerText -eq '5' -and $exportedRatings['10-2020-0011223']['인용여부(O/X)'].InnerText -eq 'O') 'Identifier-only document ratings retain the actual score and citation.'
        Assert-ExportTest ($exportedRatings['001122']['점수(0~5점)'].InnerText -eq '0' -and $exportedRatings['001122']['인용여부(O/X)'].InnerText -eq 'X') 'A zero score and false citation remain distinct from missing ratings.'
        Assert-ExportTest ($exportedRatings['10-2021-0044556']['점수(0~5점)'].InnerText -eq '' -and $exportedRatings['10-2021-0044556']['인용여부(O/X)'].InnerText -eq '') 'An unreviewed identifier-only document has blank score and citation.'
        Assert-ExportTest ($exportedRatings['001122']['비고'].GetAttribute('t') -eq 'inlineStr' -and $exportedRatings['001122']['비고'].InnerText -eq '=비고' -and $realRatingSheet.GetElementsByTagName('f').Count -eq 0) 'Formula-like rating notes remain literal text.'
    }
    finally { $zip.Dispose() }
    $realAssignedExport = Export-ReviewSnapshot -Context $operatorContext -Period $operationMonth -Scope Assigned -OutputDirectory $tempRoot -CaseId $realCase.caseId -IncludeAttachments
    Assert-ExportTest ($realAssignedExport.caseCount -eq 1 -and $realAssignedExport.documentCount -eq 3) 'Real assigned AI export succeeds without reading examiner evaluation.'
    $zip = Open-ExportTestZip $realAssignedExport.workbookPath
    try { Assert-ExportTest ((Read-ExportZipEntry $zip 'xl/worksheets/sheet1.xml').Contains('조회 권한 없음')) 'Real assigned export omits examiner evaluation.' }
    finally { $zip.Dispose() }
    Assert-ExportThrows { Export-ReviewSnapshot -Context $operatorContext -Period $operationMonth -Scope Mine -OutputDirectory $tempRoot -CaseId $realCase.caseId } 'Real per-case export respects Mine scope authorization.'

    # Mock the workflow boundary to verify stable snapshot rows and version labels.
    $global:ExportFixtureAttachmentPath = $attachmentPath
    $global:ExportFixtureEmptyCases = $false
    $global:ExportFixtureCaseFields = $false
    $global:ExportFixtureEvaluationReadCount = 0
    function global:Get-Cases {
        param($Context, $Period, $Scope)
        if ($global:ExportFixtureEmptyCases) { return }
        $record = [pscustomobject]@{ caseId = 'C1'; period = $Period; groupId = 'G1'; userId = 'U1'; applicationNumber = '001234567890'; status = 'Submitted'; version = 1; path = ''; submittedAt = '2026-10-22T00:00:00Z'; examinerName = '심사관'; duplicateWarning = $true }
        if ($global:ExportFixtureCaseFields) {
            $record | Add-Member -NotePropertyName examinerBureau -NotePropertyValue '=심사국'
            $record | Add-Member -NotePropertyName examinerDivision -NotePropertyValue '  심사과  '
            $record | Add-Member -NotePropertyName cpc -NotePropertyValue 'G06F 16/00'
        }
        return $record
    }
    function global:Get-Deliveries {
        param($Context, $Case)
        [pscustomobject]@{ version = 2; data = [pscustomobject]@{ sourceId = 'S02'; sourceName = 'ChatGPT'; noResults = $false; documents = @(
            [pscustomobject]@{ documentId = 'D1'; type = '특허'; identifier = '001122'; title = '=문헌'; description = '설명'; attachments = @([pscustomobject]@{ attachmentId = 'A1'; fileName = '문헌.txt'; relativePath = 'Data/a.txt' }) },
            [pscustomobject]@{ documentId = 'D2'; identifier = 'doi:sample'; attachments = @() }
        ) } }
        [pscustomobject]@{ version = 1; data = [pscustomobject]@{ sourceId = 'S01'; sourceName = '선행기술조사'; noResults = $true; noResultsReason = '관련 문헌 없음'; documents = @() } }
    }
    function global:Get-Evaluation {
        param($Context, $Case)
        $global:ExportFixtureEvaluationReadCount++
        return [pscustomobject]@{ version = 3; needsReconfirmation = $true; data = [pscustomobject]@{
            deliveryVersions = @{ S02 = 1 }; rankings = @([pscustomobject]@{ sourceId = 'S02'; rank = 1; status = 'Ranked' })
            documentRatings = @([pscustomobject]@{ sourceId = 'S02'; documentId = 'D1'; score = 0; cited = $false; note = '=비고' }, [pscustomobject]@{ sourceId = 'S02'; documentId = 'old'; score = 4; cited = $true; note = '이전 버전' })
        } }
    }
    function global:Resolve-DeliveryAttachment {
        param($Context, $Case, $SourceId, $Attachment)
        return $global:ExportFixtureAttachmentPath
    }
    $fixtureContext = [pscustomobject]@{ Root = $tempRoot; user = [pscustomobject]@{ userId = 'U1'; roles = @('examiner') } }
    $snapshot = Export-ReviewSnapshot -Context $fixtureContext -Period '2026-10' -Scope Mine -OutputDirectory $tempRoot -CaseId 'C1' -IncludeAttachments
    Assert-ExportTest ($snapshot.caseCount -eq 1 -and $snapshot.documentCount -eq 2) 'Snapshot returns correct case/document counts.'
    Assert-ExportTest ([IO.File]::Exists($snapshot.bundlePath)) 'Snapshot attachment bundle was created.'
    $zip = Open-ExportTestZip $snapshot.workbookPath
    try {
        [xml]$workbook = Read-ExportZipEntry $zip 'xl/workbook.xml'
        Assert-ExportTest ($workbook.GetElementsByTagName('sheet').Count -eq 4) 'Snapshot has four Korean sheets.'
        $caseXml = Read-ExportZipEntry $zip 'xl/worksheets/sheet1.xml'
        Assert-ExportTest ($caseXml.Contains('다른 심사관 중복 제출') -and $caseXml.Contains('확인 필요')) 'Cross-examiner duplicate submission warning is included.'
        $ratingXml = Read-ExportZipEntry $zip 'xl/worksheets/sheet4.xml'
        Assert-ExportTest ($ratingXml.Contains('현재 납품에 없는 문헌') -and $ratingXml.Contains('필요')) 'Stale and removed-document evaluations remain visible.'
        [xml]$ratingSheet = $ratingXml
        Assert-ExportTest ($ratingSheet.GetElementsByTagName('row').Count -eq 4) 'Current, unreviewed, and removed-document ratings are included.'
        Assert-ExportCaseFields -Archive $zip -Bureau '' -Division '' -Cpc ''
    } finally { $zip.Dispose() }
    $global:ExportFixtureCaseFields = $true
    $populatedSnapshot = Export-ReviewSnapshot -Context $fixtureContext -Period '2026-10' -Scope Mine -OutputDirectory $tempRoot -CaseId 'C1'
    $zip = Open-ExportTestZip $populatedSnapshot.workbookPath
    try {
        Assert-ExportCaseFields -Archive $zip -Bureau '=심사국' -Division '  심사과  ' -Cpc 'G06F 16/00'
        [xml]$documentSheet = Read-ExportZipEntry $zip 'xl/worksheets/sheet2.xml'
        Assert-ExportTest ($documentSheet.DocumentElement.InnerText.Contains('결과 없음')) 'No-results rows are included when checking the case fields.'
        Assert-ExportTest ($documentSheet.GetElementsByTagName('f').Count -eq 0) 'Organization fields cannot turn into Excel formulas.'
    } finally { $zip.Dispose() }
    Assert-ExportThrows { Export-ReviewSnapshot -Context $fixtureContext -Period '2026-10' -Scope Mine -OutputDirectory $tempRoot -CaseId 'not-authorized' } 'Case filter cannot bypass the authorized case list.'
    $global:ExportFixtureEmptyCases = $true
    Assert-ExportThrows { Export-ReviewSnapshot -Context $fixtureContext -Period '2026-10' -Scope Mine -OutputDirectory $tempRoot -CaseId 'C1' } 'Case export rejects an empty authorized list.'
    $emptySnapshot = Export-ReviewSnapshot -Context $fixtureContext -Period '2026-10' -Scope Mine -OutputDirectory $tempRoot
    Assert-ExportTest ($emptySnapshot.caseCount -eq 0 -and $emptySnapshot.documentCount -eq 0) 'Empty monthly export still produces a valid workbook.'
    $global:ExportFixtureEmptyCases = $false
    $readCount = $global:ExportFixtureEvaluationReadCount
    $assignedContext = [pscustomobject]@{ Root = $tempRoot; user = [pscustomobject]@{ userId = 'U2'; roles = @('examiner', 'ai') } }
    $assignedSnapshot = Export-ReviewSnapshot -Context $assignedContext -Period '2026-10' -Scope Assigned -OutputDirectory $tempRoot
    Assert-ExportTest ($global:ExportFixtureEvaluationReadCount -eq $readCount) 'Assigned AI export does not read another examiner evaluation.'
    $zip = Open-ExportTestZip $assignedSnapshot.workbookPath
    try { Assert-ExportTest ((Read-ExportZipEntry $zip 'xl/worksheets/sheet1.xml').Contains('조회 권한 없음')) 'Assigned export labels evaluation visibility correctly.' }
    finally { $zip.Dispose() }
    Write-Output 'PASS: Export workbook, XML text safety, ZIP path validation, and review snapshot tests.'
} finally {
    foreach ($functionName in @('Get-Cases', 'Get-Deliveries', 'Get-Evaluation', 'Resolve-DeliveryAttachment')) {
        Remove-Item -LiteralPath ('Function:\global:' + $functionName) -ErrorAction SilentlyContinue
    }
    Remove-Variable -Name ExportFixtureAttachmentPath -Scope Global -ErrorAction SilentlyContinue
    Remove-Variable -Name ExportFixtureEmptyCases -Scope Global -ErrorAction SilentlyContinue
    Remove-Variable -Name ExportFixtureCaseFields -Scope Global -ErrorAction SilentlyContinue
    Remove-Variable -Name ExportFixtureEvaluationReadCount -Scope Global -ErrorAction SilentlyContinue
    $resolvedTemp = [IO.Path]::GetFullPath($tempRoot)
    $expectedRoot = [IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\') + '\'
    if ($resolvedTemp.StartsWith($expectedRoot, [StringComparison]::OrdinalIgnoreCase) -and [IO.Path]::GetFileName($resolvedTemp).StartsWith('AiReviewExportTests-')) {
        Remove-Item -LiteralPath $resolvedTemp -Recurse -Force
    }
}
