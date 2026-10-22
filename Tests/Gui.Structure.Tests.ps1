$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0
if ([Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') { throw 'Run this test with powershell.exe -STA or pwsh.exe -STA.' }
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase
$appPath = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../App'))
$source = [IO.File]::ReadAllText((Join-Path $appPath 'Start.ps1'))
$tokens = $null
$parseErrors = $null
$null = [Management.Automation.Language.Parser]::ParseInput($source, [ref]$tokens, [ref]$parseErrors)
if (@($parseErrors).Count) { throw ('GUI PowerShell syntax errors: ' + ($parseErrors -join '; ')) }
$reader = [Xml.XmlReader]::Create((Join-Path $appPath 'Views/Main.xaml'))
try { $window = [Windows.Markup.XamlReader]::Load($reader) }
finally { $reader.Close() }
try {
    $controlReferences = @([regex]::Matches($source, '\$ui\.([A-Za-z][A-Za-z0-9_]*)') | ForEach-Object { $_.Groups[1].Value } | Sort-Object -Unique)
    foreach ($name in $controlReferences) {
        if ($null -eq $window.FindName($name)) { throw "Start.ps1 refers to missing XAML control '$name'." }
    }
    $eventReferences = @([regex]::Matches($source, '\$ui\.([A-Za-z][A-Za-z0-9_]*)\.Add_([A-Za-z][A-Za-z0-9_]*)\('))
    foreach ($reference in $eventReferences) {
        $name = $reference.Groups[1].Value
        $eventName = $reference.Groups[2].Value
        $control = $window.FindName($name)
        if ($null -eq $control.GetType().GetEvent($eventName)) {
            throw "Control '$name' has no event '$eventName'."
        }
    }
    foreach ($gridName in @('DocumentsGrid', 'RankGrid', 'RatingsGrid', 'SubmissionRowsGrid')) {
        $grid = $window.FindName($gridName)
        if ($grid.CanUserAddRows -or $grid.CanUserDeleteRows -or $grid.AutoGenerateColumns) {
            throw "Grid '$gridName' permits uncontrolled row/column creation."
        }
    }
    Write-Host ('GUI structure passed: {0} referenced controls, {1} event bindings, valid WPF XAML (PowerShell {2}).' -f $controlReferences.Count, $eventReferences.Count, $PSVersionTable.PSVersion)
}
finally { $window.Close() }
