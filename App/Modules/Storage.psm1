Set-StrictMode -Version 2.0

# UTF-8 JSON revisions are immutable. The lock file is deliberately never removed.
$script:Utf8 = New-Object System.Text.UTF8Encoding($false, $true)
$script:JsonHasDateKind = (Get-Command ConvertFrom-Json).Parameters.ContainsKey('DateKind')

function ConvertFrom-StorageJson {
    param([Parameter(Mandatory = $true)][string]$Json)
    # PowerShell 7.5+ otherwise turns ISO text fields into DateTime instances.
    if ($script:JsonHasDateKind) { return (ConvertFrom-Json -InputObject $Json -DateKind String -ErrorAction Stop) }
    return (ConvertFrom-Json -InputObject $Json -ErrorAction Stop)
}

function Get-SafePath {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Root,
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][string[]]$Segments
    )
    $fullRoot = [System.IO.Path]::GetFullPath($Root)
    $result = $fullRoot
    foreach ($segment in $Segments) {
        if ([string]::IsNullOrWhiteSpace($segment) -or $segment -eq '.' -or $segment -eq '..' -or
            $segment.IndexOfAny([System.IO.Path]::GetInvalidFileNameChars()) -ge 0 -or
            $segment.Contains('/') -or $segment.Contains('\') -or
            $segment.TrimEnd(' ', '.') -cne $segment -or
            $segment -match '^(?i:CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])(?:\.|$)') {
            throw "Unsafe path segment: '$segment'."
        }
        $result = [System.IO.Path]::Combine($result, $segment)
    }
    $result = [System.IO.Path]::GetFullPath($result)
    $prefix = $fullRoot.TrimEnd('\', '/') + [System.IO.Path]::DirectorySeparatorChar
    if (-not $result.Equals($fullRoot, [System.StringComparison]::OrdinalIgnoreCase) -and
        -not $result.StartsWith($prefix, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw 'The resolved path is outside the data root.'
    }
    return $result
}

function Invoke-FileLock {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][scriptblock]$Action,
        [ValidateRange(0, 600000)][int]$TimeoutMilliseconds = 5000
    )
    $fullPath = [System.IO.Path]::GetFullPath($Path)
    $parent = [System.IO.Path]::GetDirectoryName($fullPath)
    [void][System.IO.Directory]::CreateDirectory($parent)
    $timer = [System.Diagnostics.Stopwatch]::StartNew()
    $stream = $null
    $lastError = $null
    while ($null -eq $stream) {
        try {
            $stream = New-Object System.IO.FileStream($fullPath, [System.IO.FileMode]::OpenOrCreate,
                [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
        }
        catch [System.IO.IOException] {
            $lastError = $_.Exception
            if ($timer.ElapsedMilliseconds -ge $TimeoutMilliseconds) {
                throw (New-Object System.TimeoutException("Timed out acquiring file lock '$fullPath'.", $lastError))
            }
            $remaining = $TimeoutMilliseconds - [int]$timer.ElapsedMilliseconds
            Start-Sleep -Milliseconds ([Math]::Max(1, [Math]::Min(40, $remaining)))
        }
    }
    try { & $Action }
    finally { $stream.Dispose(); $timer.Stop() }
}

function Assert-RevisionEnvelope {
    param([object]$Envelope, [int]$Version, [string]$FilePath)
    if ($null -eq $Envelope -or $Envelope -is [array] -or $Envelope -isnot [pscustomobject]) {
        throw "Invalid revision object in '$FilePath'."
    }
    foreach ($name in @('schemaVersion', 'kind', 'targetId', 'version', 'actorId', 'writtenAt', 'requestId', 'reason', 'data')) {
        if ($null -eq $Envelope.PSObject.Properties[$name]) { throw "Missing '$name' in '$FilePath'." }
    }
    if ($Envelope.schemaVersion -isnot [int] -and $Envelope.schemaVersion -isnot [long]) {
        throw "Invalid schema version in '$FilePath'."
    }
    if ($Envelope.schemaVersion -ne 1) { throw "Unsupported schema version in '$FilePath'." }
    if (($Envelope.version -isnot [int] -and $Envelope.version -isnot [long]) -or $Envelope.version -ne $Version) {
        throw "Revision number does not match filename '$FilePath'."
    }
    foreach ($name in @('kind', 'targetId', 'actorId', 'requestId', 'reason', 'writtenAt')) {
        if ($Envelope.$name -isnot [string]) { throw "Invalid '$name' in '$FilePath'." }
    }
    if ([string]::IsNullOrWhiteSpace($Envelope.actorId) -or [string]::IsNullOrWhiteSpace($Envelope.requestId)) {
        throw "Missing actor or request identity in '$FilePath'."
    }
    $timestamp = [DateTimeOffset]::MinValue
    if (-not [DateTimeOffset]::TryParse($Envelope.writtenAt, [Globalization.CultureInfo]::InvariantCulture,
        [Globalization.DateTimeStyles]::RoundtripKind, [ref]$timestamp)) {
        throw "Invalid timestamp in '$FilePath'."
    }
}

function Get-Revisions {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$Path)
    $revisionPath = [System.IO.Path]::Combine([System.IO.Path]::GetFullPath($Path), 'revisions')
    if (-not [System.IO.Directory]::Exists($revisionPath)) {
        if ([System.IO.File]::Exists($revisionPath)) { throw "Revision directory is a file: '$revisionPath'." }
        return
    }
    $files = @([System.IO.Directory]::GetFiles($revisionPath, '*.json', [System.IO.SearchOption]::TopDirectoryOnly))
    $items = New-Object 'System.Collections.Generic.List[object]'
    foreach ($file in $files) {
        $name = [System.IO.Path]::GetFileName($file)
        if ($name -cnotmatch '^[0-9]{6,}\.json$') { throw "Unexpected committed revision filename '$file'." }
        $number = 0
        if (-not [int]::TryParse([System.IO.Path]::GetFileNameWithoutExtension($file), [ref]$number) -or
            $number -lt 1 -or $name -cne ('{0:D6}.json' -f $number)) {
            throw "Invalid committed revision filename '$file'."
        }
        $items.Add([pscustomobject]@{ Number = $number; Path = $file })
    }
    $expected = 1
    $requestIds = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::Ordinal)
    foreach ($item in @($items | Sort-Object Number)) {
        if ($item.Number -ne $expected) { throw "Revision sequence is incomplete at version $expected in '$revisionPath'." }
        try {
            $json = [System.IO.File]::ReadAllText($item.Path, $script:Utf8)
            $envelope = ConvertFrom-StorageJson -Json $json
        }
        catch { throw "Cannot read committed revision '$($item.Path)': $($_.Exception.Message)" }
        Assert-RevisionEnvelope -Envelope $envelope -Version $expected -FilePath $item.Path
        if (-not $requestIds.Add($envelope.requestId)) { throw "Duplicate request ID in revision history '$revisionPath'." }
        # The envelope is a scalar PSCustomObject. -NoEnumerate on newer
        # PowerShell versions can wrap it in a List[object] pipeline item.
        Write-Output $envelope
        $expected++
    }
}

function Read-Revision {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$Path)
    # Read all committed revisions, so a broken history cannot appear healthy.
    $all = @(Get-Revisions -Path $Path)
    if ($all.Count -eq 0) { return $null }
    return $all[$all.Count - 1]
}

function ConvertTo-CanonicalValue {
    param([AllowNull()][object]$Value)
    if ($null -eq $Value) { return 'null' }
    if ($Value -is [string]) { return (ConvertTo-Json -InputObject $Value -Compress -Depth 100) }
    if ($Value -is [bool]) { if ($Value) { return 'true' } else { return 'false' } }
    if ($Value -is [double] -or $Value -is [single]) {
        if ($Value -eq 0) { return '0' }
        return $Value.ToString('R', [Globalization.CultureInfo]::InvariantCulture)
    }
    if ($Value -is [decimal]) {
        return $Value.ToString('G29', [Globalization.CultureInfo]::InvariantCulture)
    }
    if ($Value -is [array]) {
        $parts = New-Object 'System.Collections.Generic.List[string]'
        foreach ($entry in $Value) { $parts.Add((ConvertTo-CanonicalValue -Value $entry)) }
        return '[' + [string]::Join(',', $parts.ToArray()) + ']'
    }
    if ($Value -is [pscustomobject]) {
        $names = New-Object 'System.Collections.Generic.List[string]'
        foreach ($property in $Value.PSObject.Properties) { $names.Add($property.Name) }
        $names.Sort([System.StringComparer]::Ordinal)
        $parts = New-Object 'System.Collections.Generic.List[string]'
        foreach ($name in $names) {
            $key = ConvertTo-Json -InputObject $name -Compress
            $parts.Add($key + ':' + (ConvertTo-CanonicalValue -Value $Value.$name))
        }
        return '{' + [string]::Join(',', $parts.ToArray()) + '}'
    }
    return (ConvertTo-Json -InputObject $Value -Compress -Depth 100)
}

function Get-SemanticPayload {
    param([object]$Value)
    # Round-trip first to normalize the JSON representation of dictionaries, dates,
    # numbers and PowerShell objects before comparing property order independently.
    $payload = [ordered]@{ kind = $Value.kind; targetId = $Value.targetId; actorId = $Value.actorId;
        reason = $Value.reason; data = $Value.data }
    $json = ConvertTo-Json -InputObject $payload -Depth 100 -Compress -ErrorAction Stop
    $normalized = ConvertFrom-StorageJson -Json $json
    return (ConvertTo-CanonicalValue -Value $normalized)
}

function Write-Revision {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][ValidateRange(0, 2147483646)][int]$ExpectedVersion,
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$RequestId,
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$ActorId,
        [AllowEmptyString()][string]$Reason = '',
        [Parameter(Mandatory = $true)][AllowNull()][object]$Data,
        [scriptblock]$Validate,
        [AllowEmptyString()][string]$Kind = '',
        [AllowEmptyString()][string]$TargetId = ''
    )
    if ([string]::IsNullOrWhiteSpace($RequestId) -or [string]::IsNullOrWhiteSpace($ActorId)) {
        throw 'ActorId and RequestId must not be whitespace.'
    }
    $entityPath = [System.IO.Path]::GetFullPath($Path)
    $lockPath = [System.IO.Path]::Combine($entityPath, 'write.lock')
    Invoke-FileLock -Path $lockPath -Action {
        $history = @(Get-Revisions -Path $entityPath)
        $payload = [pscustomobject]@{ kind = $Kind; targetId = $TargetId; actorId = $ActorId; reason = $Reason; data = $Data }
        $semantic = Get-SemanticPayload -Value $payload
        foreach ($existing in $history) {
            if ($existing.requestId -ceq $RequestId) {
                if ((Get-SemanticPayload -Value $existing) -cne $semantic) {
                    throw "Request ID '$RequestId' was already used for different content."
                }
                return $existing
            }
        }
        $current = 0
        if ($history.Count -gt 0) { $current = [int]$history[$history.Count - 1].version }
        if ($current -ne $ExpectedVersion) {
            throw "Version conflict: expected $ExpectedVersion, current $current. Reload before saving."
        }
        if ($null -ne $Validate) {
            $validationOutput = @(& $Validate $Data)
            foreach ($validationResult in $validationOutput) {
                if ($validationResult -is [bool] -and -not $validationResult) { throw 'Revision validation rejected the data.' }
            }
        }
        $next = $current + 1
        $envelope = [ordered]@{ schemaVersion = 1; kind = $Kind; targetId = $TargetId; version = $next;
            actorId = $ActorId; writtenAt = [DateTimeOffset]::UtcNow.ToString('o', [Globalization.CultureInfo]::InvariantCulture);
            requestId = $RequestId; reason = $Reason; data = $Data }
        $json = ConvertTo-Json -InputObject $envelope -Depth 100 -ErrorAction Stop
        # Check serialized JSON before creating any committed file.
        $snapshot = ConvertFrom-StorageJson -Json $json
        Assert-RevisionEnvelope -Envelope $snapshot -Version $next -FilePath $entityPath
        if ((Get-SemanticPayload -Value $snapshot) -cne $semantic) {
            throw 'The data changed during validation or could not be serialized faithfully.'
        }
        $revisionPath = [System.IO.Path]::Combine($entityPath, 'revisions')
        [void][System.IO.Directory]::CreateDirectory($revisionPath)
        $finalPath = [System.IO.Path]::Combine($revisionPath, ('{0:D6}.json' -f $next))
        $temporaryPath = [System.IO.Path]::Combine($revisionPath, ('.pending-' + [Guid]::NewGuid().ToString('N') + '.tmp'))
        try {
            $bytes = $script:Utf8.GetBytes($json)
            $output = New-Object System.IO.FileStream($temporaryPath, [System.IO.FileMode]::CreateNew,
                [System.IO.FileAccess]::Write, [System.IO.FileShare]::None)
            try { $output.Write($bytes, 0, $bytes.Length); $output.Flush($true) }
            finally { $output.Dispose() }
            $testJson = [System.IO.File]::ReadAllText($temporaryPath, $script:Utf8)
            $testEnvelope = ConvertFrom-StorageJson -Json $testJson
            Assert-RevisionEnvelope -Envelope $testEnvelope -Version $next -FilePath $temporaryPath
            if ((Get-SemanticPayload -Value $testEnvelope) -cne $semantic) { throw 'Temporary revision verification failed.' }
            # This overload cannot overwrite an existing committed revision.
            [System.IO.File]::Move($temporaryPath, $finalPath)
            $committed = Read-Revision -Path $entityPath
            if ($null -eq $committed -or $committed.version -ne $next -or $committed.requestId -cne $RequestId -or
                (Get-SemanticPayload -Value $committed) -cne $semantic) {
                throw "The save outcome is uncertain. Retry request '$RequestId' with the same content."
            }
            return $committed
        }
        finally {
            if ([System.IO.File]::Exists($temporaryPath)) { [System.IO.File]::Delete($temporaryPath) }
        }
    }
}

Export-ModuleMember -Function Get-SafePath, Invoke-FileLock, Get-Revisions, Read-Revision, Write-Revision
