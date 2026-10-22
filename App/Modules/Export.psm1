Set-StrictMode -Version 2.0

# No Office installation or third-party spreadsheet package is required.
Add-Type -AssemblyName System.IO.Compression

function Get-ExportProperty {
    param($Value, [string]$Name, $Default = $null)
    if ($null -eq $Value) { return $Default }
    if ($Value -is [System.Collections.IDictionary]) {
        if ($Value.Contains($Name)) { return $Value[$Name] }
        return $Default
    }
    $property = $Value.PSObject.Properties[$Name]
    if ($null -ne $property) { return $property.Value }
    return $Default
}

function ConvertTo-ExportXmlText {
    param($Value)
    if ($null -eq $Value) { return '' }
    $inputText = [string]$Value
    $clean = New-Object System.Text.StringBuilder
    for ($i = 0; $i -lt $inputText.Length; $i++) {
        $character = $inputText[$i]
        $code = [int]$character
        if ([char]::IsHighSurrogate($character)) {
            if (($i + 1) -lt $inputText.Length -and [char]::IsLowSurrogate($inputText[$i + 1])) {
                [void]$clean.Append($character)
                $i++
                [void]$clean.Append($inputText[$i])
            }
        } elseif ($code -eq 9 -or $code -eq 10 -or $code -eq 13 -or
            ($code -ge 32 -and $code -le 55295) -or ($code -ge 57344 -and $code -le 65533)) {
            [void]$clean.Append($character)
        }
    }
    return [System.Security.SecurityElement]::Escape($clean.ToString())
}

function Get-ExportColumnReference {
    param([int]$Number)
    $result = ''
    while ($Number -gt 0) {
        $Number--
        $result = ([string][char](65 + ($Number % 26))) + $result
        $Number = [math]::Floor($Number / 26)
    }
    return $result
}

function Get-ExportCellXml {
    param([string]$Reference, $Value, [int]$Style = 0)
    $styleText = if ($Style -gt 0) { ' s="{0}"' -f $Style } else { '' }
    if ($null -eq $Value) { return '<c r="{0}"{1}/>' -f $Reference, $styleText }
    $isNumber = $Value -is [byte] -or $Value -is [sbyte] -or $Value -is [int16] -or
        $Value -is [uint16] -or $Value -is [int32] -or $Value -is [uint32] -or
        $Value -is [int64] -or $Value -is [uint64] -or $Value -is [single] -or
        $Value -is [double] -or $Value -is [decimal]
    if ($isNumber) {
        $numericText = $Value.ToString([System.Globalization.CultureInfo]::InvariantCulture)
        if ($numericText -ne 'NaN' -and $numericText -ne 'Infinity' -and $numericText -ne '-Infinity') {
            return '<c r="{0}"{1}><v>{2}</v></c>' -f $Reference, $styleText, $numericText
        }
    }
    if ($Value -is [bool]) { $Value = if ($Value) { '예' } else { '아니오' } }
    elseif ($Value -is [datetime] -or $Value -is [datetimeoffset]) { $Value = $Value.ToString('o') }
    elseif ($Value -isnot [string] -and -not $isNumber) { $Value = ConvertTo-Json -InputObject $Value -Depth 15 -Compress }
    if (([string]$Value).Length -gt 32767) { throw "Excel 셀의 최대 글자 수(32,767)를 초과했습니다: $Reference" }
    # Every string is explicitly inline text: =, +, -, @ and leading zeros stay literal.
    return '<c r="{0}"{1} t="inlineStr"><is><t xml:space="preserve">{2}</t></is></c>' -f $Reference, $styleText, (ConvertTo-ExportXmlText $Value)
}

function Add-ExportZipText {
    param([System.IO.Compression.ZipArchive]$Archive, [string]$Name, [string]$Text)
    $entry = $Archive.CreateEntry($Name, [System.IO.Compression.CompressionLevel]::Optimal)
    $stream = $entry.Open()
    $writer = $null
    try {
        $writer = New-Object System.IO.StreamWriter($stream, (New-Object System.Text.UTF8Encoding($false)))
        $writer.Write($Text)
    } finally {
        if ($null -ne $writer) { $writer.Dispose() } else { $stream.Dispose() }
    }
}

function Get-ExportTarget {
    param([string]$Path)
    $fullPath = [System.IO.Path]::GetFullPath($Path)
    if (Test-Path -LiteralPath $fullPath) { throw "이미 존재하는 파일에는 내보낼 수 없습니다: $fullPath" }
    $directory = [System.IO.Path]::GetDirectoryName($fullPath)
    [void][System.IO.Directory]::CreateDirectory($directory)
    return [pscustomobject]@{ Path = $fullPath; TemporaryPath = (Join-Path $directory ('.export-' + [guid]::NewGuid().ToString('N') + '.tmp')) }
}

function Export-Workbook {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][System.Collections.IDictionary]$Sheets,
        [System.Collections.IDictionary]$Columns = @{}
    )
    if ($Sheets.Count -eq 0) { throw '통합문서에는 시트가 하나 이상 필요합니다.' }
    $target = Get-ExportTarget $Path
    $archive = $null
    $file = $null
    try {
        $file = [System.IO.File]::Open($target.TemporaryPath, [System.IO.FileMode]::CreateNew, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
        $archive = New-Object System.IO.Compression.ZipArchive($file, [System.IO.Compression.ZipArchiveMode]::Create, $false)
        $workbook = New-Object System.Text.StringBuilder
        [void]$workbook.Append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><bookViews><workbookView/></bookViews><sheets>')
        $relationships = New-Object System.Text.StringBuilder
        [void]$relationships.Append('<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">')
        $contentTypes = New-Object System.Text.StringBuilder
        [void]$contentTypes.Append('<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>')
        $usedNames = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
        $sheetIndex = 0
        foreach ($sheetKey in $Sheets.Keys) {
            $sheetIndex++
            $name = [regex]::Replace([string]$sheetKey, '[\x00-\x1F\[\]:*?/\\]', ' ').Trim().Trim([char]39)
            if ([string]::IsNullOrWhiteSpace($name)) { $name = '시트' }
            if ($name.Length -gt 31) { $name = $name.Substring(0, 31).TrimEnd([char[]]@([char]39)) }
            if ([char]::IsHighSurrogate($name[$name.Length - 1])) { $name = $name.Substring(0, $name.Length - 1) }
            $baseName = $name
            $suffix = 1
            while (-not $usedNames.Add($name)) {
                $suffix++
                $tail = ' ({0})' -f $suffix
                $prefixLength = [math]::Min($baseName.Length, 31 - $tail.Length)
                $name = $baseName.Substring(0, $prefixLength) + $tail
            }
            $rows = @($Sheets[$sheetKey])
            if ($rows.Count -gt 1048575) { throw "Excel의 시트별 최대 행 수를 초과했습니다: $name" }
            $columnMap = [ordered]@{}
            if ($Columns.Contains($sheetKey) -and $null -ne $Columns[$sheetKey]) {
                foreach ($key in $Columns[$sheetKey].Keys) { $columnMap[[string]$key] = [string]$Columns[$sheetKey][$key] }
            } else {
                foreach ($row in $rows) {
                    if ($null -eq $row) { continue }
                    $keys = if ($row -is [System.Collections.IDictionary]) { $row.Keys } else { $row.PSObject.Properties.Name }
                    foreach ($key in $keys) { if (-not $columnMap.Contains([string]$key)) { $columnMap[[string]$key] = [string]$key } }
                }
            }
            if ($columnMap.Count -eq 0) { $columnMap['info'] = '안내' }
            if ($columnMap.Count -gt 16384) { throw "Excel의 시트별 최대 열 수를 초과했습니다: $name" }
            $lastColumn = Get-ExportColumnReference $columnMap.Count
            $lastRow = $rows.Count + 1
            $sheetXml = New-Object System.Text.StringBuilder
            [void]$sheetXml.Append(('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:{0}{1}"/><sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft" activeCell="A2" sqref="A2"/></sheetView></sheetViews><sheetFormatPr defaultRowHeight="18"/><cols>' -f $lastColumn, $lastRow))
            $columnIndex = 0
            foreach ($key in $columnMap.Keys) {
                $columnIndex++
                $width = [math]::Max(12, [math]::Min(55, ([string]$columnMap[$key]).Length * 2 + 3))
                foreach ($row in ($rows | Select-Object -First 150)) {
                    $sample = [string](Get-ExportProperty $row $key)
                    $sampleWidth = [regex]::Replace($sample, '[^\x00-\x7F]', 'XX').Length + 2
                    $width = [math]::Max($width, [math]::Min(55, $sampleWidth))
                }
                [void]$sheetXml.Append(('<col min="{0}" max="{0}" width="{1}" customWidth="1"/>' -f $columnIndex, $width))
            }
            [void]$sheetXml.Append('</cols><sheetData><row r="1" ht="28" customHeight="1">')
            $columnIndex = 0
            foreach ($key in $columnMap.Keys) {
                $columnIndex++
                [void]$sheetXml.Append((Get-ExportCellXml ((Get-ExportColumnReference $columnIndex) + '1') $columnMap[$key] 1))
            }
            [void]$sheetXml.Append('</row>')
            $rowIndex = 1
            foreach ($row in $rows) {
                $rowIndex++
                [void]$sheetXml.Append(('<row r="{0}">' -f $rowIndex))
                $columnIndex = 0
                foreach ($key in $columnMap.Keys) {
                    $columnIndex++
                    [void]$sheetXml.Append((Get-ExportCellXml ((Get-ExportColumnReference $columnIndex) + $rowIndex) (Get-ExportProperty $row $key)))
                }
                [void]$sheetXml.Append('</row>')
            }
            [void]$sheetXml.Append(('</sheetData><autoFilter ref="A1:{0}{1}"/><pageMargins left="0.3" right="0.3" top="0.5" bottom="0.5" header="0.2" footer="0.2"/></worksheet>' -f $lastColumn, $lastRow))
            Add-ExportZipText $archive ('xl/worksheets/sheet{0}.xml' -f $sheetIndex) $sheetXml.ToString()
            [void]$workbook.Append(('<sheet name="{0}" sheetId="{1}" r:id="rId{1}"/>' -f (ConvertTo-ExportXmlText $name), $sheetIndex))
            [void]$relationships.Append(('<Relationship Id="rId{0}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{0}.xml"/>' -f $sheetIndex))
            [void]$contentTypes.Append(('<Override PartName="/xl/worksheets/sheet{0}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' -f $sheetIndex))
        }
        [void]$workbook.Append('</sheets><calcPr calcId="0" fullCalcOnLoad="0"/></workbook>')
        [void]$relationships.Append(('<Relationship Id="rId{0}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>' -f ($sheetIndex + 1)))
        [void]$contentTypes.Append('</Types>')
        Add-ExportZipText $archive '[Content_Types].xml' $contentTypes.ToString()
        Add-ExportZipText $archive '_rels/.rels' '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'
        Add-ExportZipText $archive 'xl/workbook.xml' $workbook.ToString()
        Add-ExportZipText $archive 'xl/_rels/workbook.xml.rels' $relationships.ToString()
        Add-ExportZipText $archive 'xl/styles.xml' '<?xml version="1.0" encoding="UTF-8"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2"><font><sz val="11"/><name val="맑은 고딕"/></font><font><b/><color rgb="FFFFFFFF"/><sz val="11"/><name val="맑은 고딕"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF243B53"/><bgColor indexed="64"/></patternFill></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"><alignment vertical="top" wrapText="1"/></xf><xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment vertical="center" wrapText="1"/></xf></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>'
        $archive.Dispose()
        $archive = $null
        $file = $null
        [System.IO.File]::Move($target.TemporaryPath, $target.Path)
        return $target.Path
    } finally {
        if ($null -ne $archive) { $archive.Dispose() }
        elseif ($null -ne $file) { $file.Dispose() }
        if ([System.IO.File]::Exists($target.TemporaryPath)) { [System.IO.File]::Delete($target.TemporaryPath) }
    }
}

function Get-SafeExportZipName {
    param([string]$Name)
    $normalized = $Name.Replace('\', '/')
    if ([string]::IsNullOrWhiteSpace($normalized) -or $normalized.StartsWith('/') -or
        $normalized -match '[:\x00-\x1F]' -or $normalized.EndsWith('/')) {
        throw "안전하지 않은 ZIP 내부 경로입니다: $Name"
    }
    foreach ($segment in $normalized.Split('/')) {
        if ($segment -eq '' -or $segment -eq '.' -or $segment -eq '..' -or
            $segment.EndsWith('.') -or $segment.EndsWith(' ') -or $segment -match '[<>"|?*]' -or
            $segment -match '^(?i:CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])(?:\.|$)') {
            throw "안전하지 않은 ZIP 내부 경로입니다: $Name"
        }
    }
    return $normalized
}

function Export-CaseBundle {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$WorkbookPath,
        [object[]]$Attachments = @()
    )
    $target = Get-ExportTarget $Path
    $files = New-Object 'System.Collections.Generic.List[object]'
    $names = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    $items = @([pscustomobject]@{ path = $WorkbookPath; entryName = [System.IO.Path]::GetFileName($WorkbookPath) }) + @($Attachments)
    foreach ($item in $items) {
        $sourcePath = [System.IO.Path]::GetFullPath([string](Get-ExportProperty $item 'path'))
        if (-not [System.IO.File]::Exists($sourcePath)) { throw "내보낼 파일을 찾을 수 없습니다: $sourcePath" }
        $entryName = Get-SafeExportZipName ([string](Get-ExportProperty $item 'entryName'))
        if (-not $names.Add($entryName)) { throw "ZIP 내부 파일명이 중복됩니다: $entryName" }
        $files.Add([pscustomobject]@{ path = $sourcePath; entryName = $entryName })
    }
    $archive = $null
    $file = $null
    try {
        $file = [System.IO.File]::Open($target.TemporaryPath, [System.IO.FileMode]::CreateNew, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
        $archive = New-Object System.IO.Compression.ZipArchive($file, [System.IO.Compression.ZipArchiveMode]::Create, $false)
        foreach ($item in $files) {
            $entry = $archive.CreateEntry($item.entryName, [System.IO.Compression.CompressionLevel]::Optimal)
            $inputStream = $null
            $outputStream = $null
            try {
                $inputStream = [System.IO.File]::Open($item.path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::Read)
                $outputStream = $entry.Open()
                $inputStream.CopyTo($outputStream)
            } finally {
                if ($null -ne $outputStream) { $outputStream.Dispose() }
                if ($null -ne $inputStream) { $inputStream.Dispose() }
            }
        }
        $archive.Dispose()
        $archive = $null
        $file = $null
        [System.IO.File]::Move($target.TemporaryPath, $target.Path)
        return $target.Path
    } finally {
        if ($null -ne $archive) { $archive.Dispose() }
        elseif ($null -ne $file) { $file.Dispose() }
        if ([System.IO.File]::Exists($target.TemporaryPath)) { [System.IO.File]::Delete($target.TemporaryPath) }
    }
}

function ConvertTo-ExportPathSegment {
    param([string]$Value)
    $safe = [regex]::Replace($Value, '[\x00-\x1F<>:"/\\|?*]', '_').Trim().TrimEnd([char]'.')
    if ([string]::IsNullOrWhiteSpace($safe) -or $safe -eq '..') { return 'unnamed' }
    return $safe
}

function Export-ReviewSnapshot {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]$Context,
        [Parameter(Mandatory = $true)][ValidatePattern('^\d{4}-\d{2}$')][string]$Period,
        [ValidateSet('Mine', 'Assigned', 'All')][string]$Scope = 'Mine',
        [Parameter(Mandatory = $true)][string]$OutputDirectory,
        [string]$CaseId,
        [switch]$IncludeAttachments
    )
    # Read the selected immutable revisions once, then build every sheet from these snapshots.
    $selectedCases = @(Get-Cases -Context $Context -Period $Period -Scope $Scope)
    if ($PSBoundParameters.ContainsKey('CaseId')) {
        if ([string]::IsNullOrWhiteSpace($CaseId)) { throw '내보낼 업무 건 ID를 지정하십시오.' }
        $selectedCases = @($selectedCases | Where-Object { [string]$_.caseId -eq $CaseId })
        if ($selectedCases.Count -ne 1) { throw '선택한 건을 현재 조회 범위에서 찾을 수 없습니다.' }
    }
    $currentUser = Get-ExportProperty $Context 'user'
    $currentRoles = @(Get-ExportProperty $currentUser 'roles' @())
    $currentUserId = [string](Get-ExportProperty $currentUser 'userId')
    if (-not (Get-Command Invoke-FileLock -ErrorAction SilentlyContinue)) {
        Import-Module (Join-Path $PSScriptRoot 'Storage.psm1') -ErrorAction Stop
    }
    $snapshots = New-Object 'System.Collections.Generic.List[object]'
    foreach ($case in $selectedCases) {
        $evaluationVisible = $currentRoles -contains 'admin' -or ($currentRoles -contains 'examiner' -and $currentUserId -eq [string]$case.userId)
        # Writers use this same short case lock, so delivery and evaluation versions
        # cannot change between the reads. ZIP/XLSX generation happens after release.
        $lockPath = Join-Path ([string](Get-ExportProperty $Context 'root')) ('Locks\{0}\{1}\{2}.lock' -f $Period, $case.groupId, $case.caseId)
        $captured = Invoke-FileLock -Path $lockPath -Action {
            $evaluationSnapshot = $null
            $deliverySnapshots = @(Get-Deliveries -Context $Context -Case $case)
            if ($evaluationVisible) { $evaluationSnapshot = Get-Evaluation -Context $Context -Case $case }
            [pscustomobject]@{
                Case = $case
                Deliveries = $deliverySnapshots
                Evaluation = $evaluationSnapshot
                EvaluationVisible = $evaluationVisible
            }
        }
        $snapshots.Add($captured)
    }
    $caseRows = New-Object 'System.Collections.Generic.List[object]'
    $documentRows = New-Object 'System.Collections.Generic.List[object]'
    $rankingRows = New-Object 'System.Collections.Generic.List[object]'
    $ratingRows = New-Object 'System.Collections.Generic.List[object]'
    $attachmentFiles = New-Object 'System.Collections.Generic.List[object]'
    $documentCount = 0
    foreach ($snapshot in $snapshots) {
        $case = $snapshot.Case
        # These values belong to the submitted case, not the current user profile.
        # Older submissions have no organization/CPC fields and export as blanks.
        $examinerBureau = [string](Get-ExportProperty $case 'examinerBureau' '')
        $examinerDivision = [string](Get-ExportProperty $case 'examinerDivision' '')
        $cpc = [string](Get-ExportProperty $case 'cpc' '')
        $evaluation = $snapshot.Evaluation
        $evaluationData = Get-ExportProperty $evaluation 'data'
        $ratings = @(Get-ExportProperty $evaluationData 'documentRatings' @())
        $rankings = @(Get-ExportProperty $evaluationData 'rankings' @())
        $evaluatedVersions = Get-ExportProperty $evaluationData 'deliveryVersions' @{}
        $stale = [bool](Get-ExportProperty $evaluation 'needsReconfirmation' $false)
        $seenRatings = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::Ordinal)
        $caseDocumentCount = 0
        foreach ($delivery in $snapshot.Deliveries) { $caseDocumentCount += @(Get-ExportProperty (Get-ExportProperty $delivery 'data') 'documents' @()).Count }
        $caseRows.Add([pscustomobject][ordered]@{
            period = $Period; groupId = $case.groupId; examiner = $case.examinerName; examinerBureau = $examinerBureau; examinerDivision = $examinerDivision; applicationNumber = [string]$case.applicationNumber; cpc = $cpc
            caseId = $case.caseId; status = $case.status; submissionVersion = $case.version; submittedAt = $case.submittedAt
            duplicateWarning = $(if ([bool](Get-ExportProperty $case 'duplicateWarning' $false)) { '확인 필요' } else { '' })
            publishedSources = $snapshot.Deliveries.Count; documentCount = $caseDocumentCount
            evaluationVersion = Get-ExportProperty $evaluation 'version'
            evaluationStatus = $(if (-not $snapshot.EvaluationVisible) { '조회 권한 없음' } elseif ($null -eq $evaluation) { '미평가' } elseif ($stale) { '재확인 필요' } elseif ([bool](Get-ExportProperty $evaluationData 'complete' $false)) { '평가 완료' } else { '평가 중' })
        })
        foreach ($delivery in $snapshot.Deliveries) {
            $data = $delivery.data
            $sourceId = [string]$data.sourceId
            $sourceName = [string]$data.sourceName
            $sourceVersion = $delivery.version
            $evaluatedVersion = Get-ExportProperty $evaluatedVersions $sourceId
            $sourceStale = $null -ne $evaluation -and ($null -eq $evaluatedVersion -or [string]$evaluatedVersion -ne [string]$sourceVersion)
            $ranking = @($rankings | Where-Object { [string](Get-ExportProperty $_ 'sourceId') -eq $sourceId } | Select-Object -First 1)
            $rankingValue = if ($ranking.Count -gt 0) { Get-ExportProperty $ranking[0] 'rank' } else { $null }
            $rankingStatus = if ($ranking.Count -gt 0) { Get-ExportProperty $ranking[0] 'status' } else { '미평가' }
            $rankingRows.Add([pscustomobject][ordered]@{
                period = $Period; groupId = $case.groupId; examiner = $case.examinerName; examinerBureau = $examinerBureau; examinerDivision = $examinerDivision; applicationNumber = [string]$case.applicationNumber; cpc = $cpc; caseId = $case.caseId
                sourceId = $sourceId; sourceName = $sourceName; deliveryVersion = $sourceVersion; evaluatedVersion = $evaluatedVersion
                rank = $rankingValue; status = $rankingStatus; needsReconfirmation = $(if ($stale -or $sourceStale) { '필요' } else { '' })
            })
            $documents = @(Get-ExportProperty $data 'documents' @())
            if ($documents.Count -eq 0 -and [bool](Get-ExportProperty $data 'noResults' $false)) {
                $documentRows.Add([pscustomobject][ordered]@{
                    period = $Period; groupId = $case.groupId; examiner = $case.examinerName; examinerBureau = $examinerBureau; examinerDivision = $examinerDivision; applicationNumber = [string]$case.applicationNumber; cpc = $cpc; caseId = $case.caseId
                    sourceId = $sourceId; sourceName = $sourceName; deliveryVersion = $sourceVersion; documentId = ''; type = '결과 없음'; identifier = ''; title = ''
                    description = Get-ExportProperty $data 'noResultsReason'; attachments = ''
                })
            }
            foreach ($document in $documents) {
                $documentCount++
                $documentId = [string]$document.documentId
                $attachmentNames = New-Object 'System.Collections.Generic.List[string]'
                foreach ($attachment in @(Get-ExportProperty $document 'attachments' @())) {
                    $attachmentNames.Add([string]$attachment.fileName)
                    if ($IncludeAttachments) {
                        $attachmentPath = Resolve-DeliveryAttachment -Context $Context -Case $case -SourceId $sourceId -Attachment $attachment
                        $entryName = '첨부문헌/{0}/{1}/{2}/{3}_{4}' -f (ConvertTo-ExportPathSegment $case.caseId), (ConvertTo-ExportPathSegment $sourceId), (ConvertTo-ExportPathSegment $documentId), (ConvertTo-ExportPathSegment ([string]$attachment.attachmentId)), (ConvertTo-ExportPathSegment ([string]$attachment.fileName))
                        $attachmentFiles.Add([pscustomobject]@{ path = $attachmentPath; entryName = $entryName })
                    }
                }
                $documentRows.Add([pscustomobject][ordered]@{
                    period = $Period; groupId = $case.groupId; examiner = $case.examinerName; examinerBureau = $examinerBureau; examinerDivision = $examinerDivision; applicationNumber = [string]$case.applicationNumber; cpc = $cpc; caseId = $case.caseId
                    sourceId = $sourceId; sourceName = $sourceName; deliveryVersion = $sourceVersion; documentId = $documentId; type = Get-ExportProperty $document 'type' ''
                    identifier = [string]$document.identifier; title = Get-ExportProperty $document 'title' ''; description = Get-ExportProperty $document 'description' ''; attachments = ($attachmentNames -join "`n")
                })
                $ratingKey = $sourceId + ':' + $documentId
                $rating = @($ratings | Where-Object { [string](Get-ExportProperty $_ 'sourceId') -eq $sourceId -and [string](Get-ExportProperty $_ 'documentId') -eq $documentId } | Select-Object -First 1)
                $ratingData = if ($rating.Count -gt 0) { $rating[0] } else { $null }
                if ($null -ne $ratingData) { [void]$seenRatings.Add($ratingKey) }
                $cited = Get-ExportProperty $ratingData 'cited'
                $ratingRows.Add([pscustomobject][ordered]@{
                    period = $Period; groupId = $case.groupId; examiner = $case.examinerName; examinerBureau = $examinerBureau; examinerDivision = $examinerDivision; applicationNumber = [string]$case.applicationNumber; cpc = $cpc; caseId = $case.caseId
                    sourceId = $sourceId; sourceName = $sourceName; deliveryVersion = $sourceVersion; evaluatedVersion = $evaluatedVersion; documentId = $documentId
                    identifier = [string]$document.identifier; score = Get-ExportProperty $ratingData 'score'
                    cited = $(if ($null -eq $cited) { '' } elseif ([bool]$cited) { 'O' } else { 'X' }); note = Get-ExportProperty $ratingData 'note'
                    needsReconfirmation = $(if ($sourceStale) { '필요' } else { '' })
                })
            }
        }
        # Keep evaluations of removed documents visible, linked to their evaluated revision.
        foreach ($rating in $ratings) {
            $sourceId = [string](Get-ExportProperty $rating 'sourceId')
            $documentId = [string](Get-ExportProperty $rating 'documentId')
            if ($seenRatings.Contains($sourceId + ':' + $documentId)) { continue }
            $cited = Get-ExportProperty $rating 'cited'
            $ratingRows.Add([pscustomobject][ordered]@{
                period = $Period; groupId = $case.groupId; examiner = $case.examinerName; examinerBureau = $examinerBureau; examinerDivision = $examinerDivision; applicationNumber = [string]$case.applicationNumber; cpc = $cpc; caseId = $case.caseId
                sourceId = $sourceId; sourceName = $sourceId; deliveryVersion = $null; evaluatedVersion = Get-ExportProperty $evaluatedVersions $sourceId; documentId = $documentId
                identifier = '현재 납품에 없는 문헌'; score = Get-ExportProperty $rating 'score'; cited = $(if ($null -eq $cited) { '' } elseif ([bool]$cited) { 'O' } else { 'X' })
                note = Get-ExportProperty $rating 'note'; needsReconfirmation = '필요'
            })
        }
    }
    $common = [ordered]@{ period = '운영월'; groupId = '그룹 ID'; examiner = '심사관'; examinerBureau = '심사국'; examinerDivision = '심사과'; applicationNumber = '출원번호'; cpc = 'CPC'; caseId = '업무 건 ID' }
    $caseColumns = [ordered]@{}; foreach ($key in $common.Keys) { $caseColumns[$key] = $common[$key] }
    $caseColumns['status'] = '제출 상태'; $caseColumns['submittedAt'] = '제출 시각'; $caseColumns['duplicateWarning'] = '다른 심사관 중복 제출'; $caseColumns['publishedSources'] = '납품 방식 수'; $caseColumns['documentCount'] = '납품 문헌 수'; $caseColumns['evaluationStatus'] = '평가 상태'
    $documentColumns = [ordered]@{}; foreach ($key in $common.Keys) { $documentColumns[$key] = $common[$key] }
    $documentColumns['sourceId'] = '조사 방식 ID'; $documentColumns['sourceName'] = '조사 방식'; $documentColumns['documentId'] = '문헌 ID'; $documentColumns['type'] = '문헌 유형'; $documentColumns['identifier'] = '문헌 식별정보'; $documentColumns['title'] = '명칭'; $documentColumns['description'] = '관련 설명'; $documentColumns['attachments'] = '첨부파일'
    $rankingColumns = [ordered]@{}; foreach ($key in $common.Keys) { $rankingColumns[$key] = $common[$key] }
    $rankingColumns['sourceId'] = '조사 방식 ID'; $rankingColumns['sourceName'] = '조사 방식'; $rankingColumns['rank'] = '품질 순위'; $rankingColumns['status'] = '순위 상태'; $rankingColumns['needsReconfirmation'] = '재확인'
    $ratingColumns = [ordered]@{}; foreach ($key in $common.Keys) { $ratingColumns[$key] = $common[$key] }
    $ratingColumns['sourceId'] = '조사 방식 ID'; $ratingColumns['sourceName'] = '조사 방식'; $ratingColumns['documentId'] = '문헌 ID'; $ratingColumns['identifier'] = '문헌'; $ratingColumns['score'] = '점수(0~5점)'; $ratingColumns['cited'] = '인용여부(O/X)'; $ratingColumns['note'] = '비고'; $ratingColumns['needsReconfirmation'] = '재확인'
    $sheets = [ordered]@{ '출원별 현황' = $caseRows.ToArray(); '선행기술 목록' = $documentRows.ToArray(); '품질 순위' = $rankingRows.ToArray(); '문헌별 평가' = $ratingRows.ToArray() }
    $columns = @{ '출원별 현황' = $caseColumns; '선행기술 목록' = $documentColumns; '품질 순위' = $rankingColumns; '문헌별 평가' = $ratingColumns }
    $outputRoot = [System.IO.Path]::GetFullPath($OutputDirectory)
    [void][System.IO.Directory]::CreateDirectory($outputRoot)
    $baseName = 'AI조사_{0}_{1}_{2}_{3}' -f $Period, $Scope, (Get-Date -Format 'yyyyMMdd_HHmmss_fff'), [guid]::NewGuid().ToString('N').Substring(0, 6)
    $workbookPath = Export-Workbook -Path (Join-Path $outputRoot ($baseName + '.xlsx')) -Sheets $sheets -Columns $columns
    $bundlePath = $null
    if ($IncludeAttachments) { $bundlePath = Export-CaseBundle -Path (Join-Path $outputRoot ($baseName + '.zip')) -WorkbookPath $workbookPath -Attachments $attachmentFiles.ToArray() }
    return [pscustomobject]@{ workbookPath = $workbookPath; bundlePath = $bundlePath; caseCount = $snapshots.Count; documentCount = $documentCount; exportedAt = [datetime]::UtcNow.ToString('o') }
}

Export-ModuleMember -Function Export-Workbook, Export-CaseBundle, Export-ReviewSnapshot
