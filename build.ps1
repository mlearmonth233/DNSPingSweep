<#
.SYNOPSIS
    Builds the NetTools Excel add-in (NetTools.xlam) and, optionally, its MSI installer.

.DESCRIPTION
    1. Starts Excel in the background, imports src\NetTools.bas into a new
       workbook and saves it as an add-in (.xlam).
    2. Adds the ribbon (src\customUI.xml) and the add-in's title/description
       to the .xlam package.
    3. With -Msi: downloads the MAC vendor list (IEEE OUI registry) to
       dist\oui.txt and packages it with the .xlam into
       dist\NetTools-<version>.msi with the WiX Toolset (installed
       automatically as a .NET tool if missing).

    Needs Windows with desktop Excel (2010 or later). Step 1 needs
    "Trust access to the VBA project object model"; the script switches it
    on for the duration of the build and puts it back afterwards.

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\build.ps1
.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\build.ps1 -Msi
#>
[CmdletBinding()]
param(
    [switch]$Msi,
    [string]$Version
)

$ErrorActionPreference = 'Stop'

# Windows PowerShell 5.1 leaves $PSScriptRoot empty inside param() defaults,
# so work out the script folder here instead.
$root = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
if (-not $Version) { $Version = (Get-Content (Join-Path $root 'VERSION') -Raw).Trim() }
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

$srcDir   = Join-Path $root 'src'
$distDir  = Join-Path $root 'dist'
$basFile  = Join-Path $srcDir 'NetTools.bas'
$uiFile   = Join-Path $srcDir 'customUI.xml'
$xlamFile = Join-Path $distDir 'NetTools.xlam'
$wxsFile  = Join-Path $root 'installer\NetTools.wxs'
$msiFile  = Join-Path $distDir "NetTools-$Version.msi"
$ouiFile  = Join-Path $distDir 'oui.txt'
$plinkFile = Join-Path $distDir 'plink.exe'
$loginFile = Join-Path $srcDir 'NetToolsLogin.vba'
$wixVersion = '5.0.2'

New-Item -ItemType Directory -Force -Path $distDir | Out-Null
if (Test-Path $xlamFile) { Remove-Item $xlamFile -Force }

function Release-Com([object]$o) {
    if ($null -ne $o) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($o) }
}

# ---------------------------------------------------------------------------
# 1. Build the .xlam with Excel
# ---------------------------------------------------------------------------
Write-Host "Building NetTools.xlam $Version ..."

# Excel's version (e.g. 16.0) tells us where its Trust Center settings live
$probe = New-Object -ComObject Excel.Application
$officeVer = $probe.Version
$probe.Quit(); Release-Com $probe; $probe = $null
[GC]::Collect(); [GC]::WaitForPendingFinalizers()

$secKey = "HKCU:\Software\Microsoft\Office\$officeVer\Excel\Security"
if (-not (Test-Path $secKey)) { New-Item -Path $secKey -Force | Out-Null }
$oldVbom = (Get-ItemProperty -Path $secKey -Name AccessVBOM -ErrorAction SilentlyContinue).AccessVBOM
Set-ItemProperty -Path $secKey -Name AccessVBOM -Value 1 -Type DWord

$excel = $null; $wb = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $wb = $excel.Workbooks.Add()
    try {
        [void]$wb.VBProject.VBComponents.Import($basFile)
    } catch {
        throw "Excel refused access to the VBA project. In Excel: File > Options > Trust Center > " +
              "Trust Center Settings > Macro Settings, tick 'Trust access to the VBA project object model', " +
              "then run this script again. ($($_.Exception.Message))"
    }
    # SSH login box: an empty UserForm (3 = vbext_ct_MSForm) with code that
    # builds its controls at run time - so no binary .frx has to be kept in git
    $form = $wb.VBProject.VBComponents.Add(3)
    $form.Name = 'NetToolsLogin'
    $code = $form.CodeModule
    if ($code.CountOfLines -gt 0) { $code.DeleteLines(1, $code.CountOfLines) }
    $code.AddFromString(([IO.File]::ReadAllText($loginFile) -replace "`r?`n", "`r`n"))
    $wb.SaveAs($xlamFile, 55)   # 55 = xlOpenXMLAddIn
    $wb.Close($false)
} finally {
    if ($null -ne $excel) { $excel.Quit() }
    Release-Com $wb; Release-Com $excel
    $wb = $null; $excel = $null
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    if ($null -eq $oldVbom) {
        Remove-ItemProperty -Path $secKey -Name AccessVBOM -ErrorAction SilentlyContinue
    } else {
        Set-ItemProperty -Path $secKey -Name AccessVBOM -Value $oldVbom -Type DWord
    }
}

# ---------------------------------------------------------------------------
# 2. Add the ribbon XML and document properties to the package
# ---------------------------------------------------------------------------
function Read-ZipXml($zip, [string]$name) {
    $entry = $zip.GetEntry($name)
    if ($null -eq $entry) { return $null }
    $reader = New-Object IO.StreamReader($entry.Open())
    try { $xml = New-Object Xml.XmlDocument; $xml.LoadXml($reader.ReadToEnd()); return $xml }
    finally { $reader.Close() }
}

function Write-ZipText($zip, [string]$name, [string]$text) {
    $old = $zip.GetEntry($name)
    if ($null -ne $old) { $old.Delete() }
    $writer = New-Object IO.StreamWriter($zip.CreateEntry($name).Open(), (New-Object Text.UTF8Encoding($false)))
    try { $writer.Write($text) } finally { $writer.Close() }
}

Write-Host "Adding ribbon ..."
$zip = [IO.Compression.ZipFile]::Open($xlamFile, [IO.Compression.ZipArchiveMode]::Update)
try {
    Write-ZipText $zip 'customUI/customUI.xml' ([IO.File]::ReadAllText($uiFile))

    # Package relationship that points Office at the ribbon XML
    $relNs   = 'http://schemas.openxmlformats.org/package/2006/relationships'
    $relType = 'http://schemas.microsoft.com/office/2006/relationships/ui/extensibility'
    $rels = Read-ZipXml $zip '_rels/.rels'
    $exists = $false
    foreach ($r in $rels.DocumentElement.ChildNodes) { if ($r.GetAttribute('Type') -eq $relType) { $exists = $true } }
    if (-not $exists) {
        $node = $rels.CreateElement('Relationship', $relNs)
        $node.SetAttribute('Id', 'rIdNetToolsRibbon')
        $node.SetAttribute('Type', $relType)
        $node.SetAttribute('Target', 'customUI/customUI.xml')
        [void]$rels.DocumentElement.AppendChild($node)
    }
    Write-ZipText $zip '_rels/.rels' $rels.OuterXml

    # Title/description shown in Excel's Add-ins dialog
    $dcNs = 'http://purl.org/dc/elements/1.1/'
    $core = Read-ZipXml $zip 'docProps/core.xml'
    if ($null -ne $core) {
        $props = [ordered]@{ title = 'NetTools'; description = 'DNS lookup, reverse lookup and ping for a column of IP addresses or hostnames.' }
        foreach ($k in $props.Keys) {
            $n = $core.DocumentElement.GetElementsByTagName($k, $dcNs) | Select-Object -First 1
            if ($null -eq $n) { $n = $core.CreateElement('dc', $k, $dcNs); [void]$core.DocumentElement.AppendChild($n) }
            $n.InnerText = $props[$k]
        }
        Write-ZipText $zip 'docProps/core.xml' $core.OuterXml
    }
} finally {
    $zip.Dispose()
}
Write-Host "Built $xlamFile" -ForegroundColor Green

# ---------------------------------------------------------------------------
# 3. MSI (optional)
# ---------------------------------------------------------------------------
if (-not $Msi) { return }

# MSI versions must be numeric x.y.z
if ($Version -notmatch '^\d+\.\d+(\.\d+)?$') { throw "VERSION '$Version' must look like 1.0.0" }

$wix = Get-Command wix -ErrorAction SilentlyContinue
if ($null -eq $wix) {
    if ($null -eq (Get-Command dotnet -ErrorAction SilentlyContinue)) {
        throw "The MSI needs the WiX Toolset, which needs the .NET SDK (6 or later): https://dotnet.microsoft.com/download"
    }
    Write-Host "Installing WiX Toolset $wixVersion (dotnet tool) ..."
    dotnet tool install --global wix --version $wixVersion
    if ($LASTEXITCODE -ne 0) { throw "Could not install the WiX Toolset." }
    $env:PATH += ";$env:USERPROFILE\.dotnet\tools"
}

# MAC vendor list: IEEE registry, falling back to Wireshark's copy of it.
# Written as "AABBCC<tab>Vendor" lines. Downloaded once per build.
function Get-OuiLines {
    $ua = @{ 'User-Agent' = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) NetTools-build' }
    try {
        $r = Invoke-WebRequest -UseBasicParsing -Headers $ua -TimeoutSec 120 -Uri 'https://standards-oui.ieee.org/oui/oui.csv'
        $text = [Text.Encoding]::UTF8.GetString($r.RawContentStream.ToArray())
        $lines = $text | ConvertFrom-Csv | Where-Object { $_.Assignment -match '^[0-9A-Fa-f]{6}$' } |
            ForEach-Object { "{0}`t{1}" -f $_.Assignment.ToUpper(), (($_.'Organization Name' -replace '\s+', ' ').Trim()) }
        if ($lines.Count -gt 1000) { return $lines }
    } catch { Write-Warning "IEEE OUI download failed: $($_.Exception.Message)" }
    try {
        $r = Invoke-WebRequest -UseBasicParsing -Headers $ua -TimeoutSec 120 -Uri 'https://www.wireshark.org/download/automated/data/manuf'
        $text = [Text.Encoding]::UTF8.GetString($r.RawContentStream.ToArray())
        $lines = foreach ($l in ($text -split "`n")) {
            if ($l -match '^([0-9A-Fa-f]{2}):([0-9A-Fa-f]{2}):([0-9A-Fa-f]{2})\s+(\S+)(?:\s+(.+))?$') {
                $name = if ($matches[5]) { $matches[5].Trim() } else { $matches[4] }
                "{0}{1}{2}`t{3}" -f $matches[1].ToUpper(), $matches[2].ToUpper(), $matches[3].ToUpper(), $name
            }
        }
        if ($lines.Count -gt 1000) { return $lines }
    } catch { Write-Warning "Wireshark manuf download failed: $($_.Exception.Message)" }
    return @()
}

# PuTTY's plink.exe for SSH Logs - official build, Authenticode-checked
Write-Host "Downloading plink.exe (PuTTY) ..."
$hasPlink = 0
try {
    $tmp = "$plinkFile.download"
    Invoke-WebRequest -UseBasicParsing -TimeoutSec 120 -OutFile $tmp `
        -Headers @{ 'User-Agent' = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) NetTools-build' } `
        -Uri 'https://the.earth.li/~sgtatham/putty/latest/w64/plink.exe'
    $sig = Get-AuthenticodeSignature $tmp
    if ($sig.Status -ne 'Valid' -or $sig.SignerCertificate.Subject -notmatch 'Simon Tatham') {
        Remove-Item $tmp -Force
        throw "signature check failed ($($sig.Status), $($sig.SignerCertificate.Subject))"
    }
    Move-Item -Force $tmp $plinkFile
    Write-Host "  $((Get-Item $plinkFile).VersionInfo.ProductVersion)"
} catch {
    Write-Warning "plink.exe download failed: $($_.Exception.Message)"
}
if (Test-Path $plinkFile) { $hasPlink = 1 } else {
    Write-Warning "The MSI won't include plink.exe - SSH Logs will need PuTTY 0.77+ installed."
}

Write-Host "Downloading MAC vendor list ..."
[Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
$ouiLines = @(Get-OuiLines)
if ($ouiLines.Count -gt 0) {
    [IO.File]::WriteAllLines($ouiFile, [string[]]$ouiLines, (New-Object Text.UTF8Encoding($false)))
    Write-Host "  $($ouiLines.Count) vendors"
} elseif (Test-Path $ouiFile) {
    Write-Warning "Using the vendor list from the last build."
} else {
    Write-Warning "No MAC vendor list - the Vendor column will be empty. Re-run the build when online."
    [IO.File]::WriteAllText($ouiFile, "# vendor list unavailable`n")
}

Write-Host "Building MSI ..."
if (Test-Path $msiFile) { Remove-Item $msiFile -Force }
wix build $wxsFile -arch x86 -d "Version=$Version" -d "XlamPath=$xlamFile" -d "OuiPath=$ouiFile" -d "PlinkPath=$plinkFile" -d "HasPlink=$hasPlink" -o $msiFile
if ($LASTEXITCODE -ne 0) { throw "wix build failed." }
Remove-Item (Join-Path $distDir '*.wixpdb') -ErrorAction SilentlyContinue
Write-Host "Built $msiFile" -ForegroundColor Green
