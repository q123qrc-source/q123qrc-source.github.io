<#
.SYNOPSIS
    Office LTSC Installer - installs official Office LTSC 2024 / 2021 / 2019
    using Microsoft's Office Deployment Tool (ODT).

.DESCRIPTION
    Independent tool. Not affiliated with Microsoft.
    Office is downloaded directly from Microsoft's servers (officecdn.microsoft.com).

    The script:
      1. Elevates itself (UAC) if needed.
      2. Checks the OS, free disk space and internet access.
      3. Downloads setup.exe (ODT) from officecdn.microsoft.com and verifies
         its Authenticode signature (must be Valid and signed by Microsoft Corporation).
      4. Writes configuration.xml and runs "setup.exe /configure" (or /download).
      5. If a product key was given, reports the activation status.

    Run with no parameters for an interactive menu:
        irm https://q123qrc-source.github.io/ps | iex

    Or non-interactively:
        & ([scriptblock]::Create((irm https://q123qrc-source.github.io/ps))) -Version 2024 -Language ar-sa -Arch 64

    Exit codes:
        0 installed (and activated, if a key was given)
        1 cancelled by user
        2 pre-check failed (OS, disk, network)
        3 ODT / OfficeScrubber download or signature/hash check failed
        4 ODT install failed
        5 installed but activation not confirmed
        6 invalid input

.LINK
    https://github.com/q123qrc-source/q123qrc-source.github.io
#>
[CmdletBinding()]
param(
    [ValidateSet('2024', '2021', '2019')] [string] $Version,
    [string]   $Language,
    [string[]] $ExtraLanguages,
    [ValidateSet('64', '32')] [string] $Arch = '64',
    [ValidateSet('ProPlus', 'Standard')] [string] $Edition = 'ProPlus',
    [switch]   $Visio,
    [switch]   $Project,
    [string[]] $ExcludeApps,
    [string]   $ProductKey,
    [switch]   $RemoveExisting,
    [switch]   $Scrub,
    [string]   $SourcePath,
    [switch]   $DownloadOnly,
    [switch]   $NoUpdates,
    [switch]   $Silent,
    [Parameter(DontShow = $true)] [switch] $Relaunched,
    [Parameter(DontShow = $true)] [string] $KeyFile
)

# Captured at top level so the elevated relaunch can reuse the exact script text.
$script:SelfInvocation = $MyInvocation
$script:SelfCommandPath = $PSCommandPath

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

function Get-InstallerConstant {
    [CmdletBinding()]
    [OutputType([hashtable])]
    param()
    @{
        ProductName     = 'Office LTSC Installer'
        SiteBase        = 'https://q123qrc-source.github.io'
        OdtUrl          = 'https://officecdn.microsoft.com/pr/wsus/setup.exe'
        OdtSigner       = 'O=Microsoft Corporation'
        MinFreeGB       = 10
        MinBuild        = 17763
        # OfficeScrubber by abbodi1406 (github.com/abbodi1406/BatUtil).
        # Not bundled: downloaded at runtime from a pinned commit and every file
        # is checked against the SHA-256 values below.
        ScrubberCommit  = '4c02e02d86d2f01e4ef556bc6a080382bd21cae4'
        ScrubberFiles   = [ordered]@{
            'OfficeScrubber.cmd'      = '52465132caa232d190946ad699c4cc00a1778fdbee29b20c6054c111a3769793'
            'bin/CleanOffice.txt'     = 'c0ff3e0b9cb14687e9ef6da7890941f753f6541ede627a3d14f3d43a20b1163c'
            'bin/OffScrub03.vbs'      = '912bab2ea3c00bf13f38aab7578ca5b3da925aae1c8fced47b1153a61325d9a1'
            'bin/OffScrub07.vbs'      = '4f1e42ce509fd0108a2bf261c5bc6e04b749dc7431cef083572623946dbe697b'
            'bin/OffScrub10.vbs'      = 'c7319acd86e6489d44d1723fa9f71f6e76899474d2473b5ab153c01e6c542be9'
            'bin/OffScrubC2R.vbs'     = 'cd94604f7636085641273cad146266dfc997724b4faabe1be42c21310c040dec'
            'bin/OffScrub_O15msi.vbs' = '30704d9f3abfd25a9c3318bf5638bfcfda40446ac0103ff3bb01a28079f70b9b'
            'bin/OffScrub_O16msi.vbs' = 'a72bb334cfff2dc236be4160b495d8f45849d9ad061a17a2396bc6158ac8b0ef'
        }
    }
}

function Get-ProductCatalog {
    # Verified against Microsoft Learn ("Product IDs supported by the Office
    # Deployment Tool" and the LTSC 2024 / 2021 / 2019 deployment overviews).
    [CmdletBinding()]
    [OutputType([hashtable])]
    param()
    @{
        '2024' = @{ Channel = 'PerpetualVL2024'; ProPlus = 'ProPlus2024Volume'; Standard = 'Standard2024Volume'; Project = 'ProjectPro2024Volume'; Visio = 'VisioPro2024Volume' }
        '2021' = @{ Channel = 'PerpetualVL2021'; ProPlus = 'ProPlus2021Volume'; Standard = 'Standard2021Volume'; Project = 'ProjectPro2021Volume'; Visio = 'VisioPro2021Volume' }
        '2019' = @{ Channel = 'PerpetualVL2019'; ProPlus = 'ProPlus2019Volume'; Standard = 'Standard2019Volume'; Project = 'ProjectPro2019Volume'; Visio = 'VisioPro2019Volume' }
    }
}

function Get-ExcludeAppList {
    # ExcludeApp IDs accepted by ODT (Microsoft Learn: ODT configuration options).
    [CmdletBinding()]
    [OutputType([string[]])]
    param()
    @('Access', 'Excel', 'Groove', 'Lync', 'OneDrive', 'OneNote', 'Outlook',
        'OutlookForWindows', 'PowerPoint', 'Publisher', 'Teams', 'Word')
}

function Get-LanguageTable {
    # Culture codes from Microsoft Learn: "Languages, culture codes, and
    # companion proofing languages". English names only (script must be ASCII);
    # native names live in languages.json on the website.
    [CmdletBinding()]
    [OutputType([object[]])]
    param()
    $raw = @'
af-za|Afrikaans
sq-al|Albanian
ar-sa|Arabic
hy-am|Armenian
as-in|Assamese
az-latn-az|Azerbaijani (Latin)
bn-bd|Bangla (Bangladesh)
bn-in|Bangla (India)
eu-es|Basque
bs-latn-ba|Bosnian (Latin)
bg-bg|Bulgarian
ca-es|Catalan
ca-es-valencia|Catalan (Valencia)
zh-cn|Chinese (Simplified)
zh-tw|Chinese (Traditional)
hr-hr|Croatian
cs-cz|Czech
da-dk|Danish
nl-nl|Dutch
en-us|English (United States)
en-gb|English (United Kingdom)
et-ee|Estonian
fi-fi|Finnish
fr-fr|French
fr-ca|French (Canada)
gl-es|Galician
ka-ge|Georgian
de-de|German
el-gr|Greek
gu-in|Gujarati
ha-latn-ng|Hausa
he-il|Hebrew
hi-in|Hindi
hu-hu|Hungarian
is-is|Icelandic
ig-ng|Igbo
id-id|Indonesian
ga-ie|Irish
xh-za|isiXhosa
zu-za|isiZulu
it-it|Italian
ja-jp|Japanese
kn-in|Kannada
kk-kz|Kazakh
rw-rw|Kinyarwanda
sw-ke|KiSwahili
kok-in|Konkani
ko-kr|Korean
ky-kg|Kyrgyz
lv-lv|Latvian
lt-lt|Lithuanian
lb-lu|Luxembourgish
mk-mk|Macedonian
ms-my|Malay (Latin)
ml-in|Malayalam
mt-mt|Maltese
mi-nz|Maori
mr-in|Marathi
ne-np|Nepali
nb-no|Norwegian Bokmal
nn-no|Norwegian Nynorsk
or-in|Odia
ps-af|Pashto
fa-ir|Persian
pl-pl|Polish
pt-pt|Portuguese (Portugal)
pt-br|Portuguese (Brazil)
pa-in|Punjabi (Gurmukhi)
ro-ro|Romanian
rm-ch|Romansh
ru-ru|Russian
gd-gb|Scottish Gaelic
sr-cyrl-rs|Serbian (Cyrillic, Serbia)
sr-latn-rs|Serbian (Latin, Serbia)
sr-cyrl-ba|Serbian (Cyrillic, Bosnia and Herzegovina)
nso-za|Sesotho sa Leboa
tn-za|Setswana
si-lk|Sinhala
sk-sk|Slovak
sl-si|Slovenian
es-es|Spanish
es-mx|Spanish (Mexico)
sv-se|Swedish
ta-in|Tamil
tt-ru|Tatar (Cyrillic)
te-in|Telugu
th-th|Thai
tr-tr|Turkish
uk-ua|Ukrainian
ur-pk|Urdu
uz-latn-uz|Uzbek (Latin)
vi-vn|Vietnamese
cy-gb|Welsh
wo-sn|Wolof
yo-ng|Yoruba
'@
    foreach ($line in ($raw -split "`r?`n")) {
        if ($line.Trim().Length -eq 0) { continue }
        $parts = $line.Split('|')
        [pscustomobject]@{ Id = $parts[0].Trim(); Name = $parts[1].Trim() }
    }
}

# ---------------------------------------------------------------------------
# Errors and logging
# ---------------------------------------------------------------------------

function Get-InstallerError {
    # Builds an exception that carries the script exit code.
    [CmdletBinding()]
    [OutputType([System.Exception])]
    param(
        [Parameter(Mandatory = $true)] [int] $Code,
        [Parameter(Mandatory = $true)] [string] $Message
    )
    $ex = New-Object System.Exception $Message
    $ex.Data['OfficeLtscExitCode'] = $Code
    $ex
}

function Get-MaskedKey {
    [CmdletBinding()]
    [OutputType([string])]
    param([string] $Key)
    if ([string]::IsNullOrEmpty($Key)) { return '(none)' }
    if ($Key.Length -le 5) { return '*****' }
    'XXXXX-XXXXX-XXXXX-XXXXX-' + $Key.Substring($Key.Length - 5)
}

function Get-LogPath {
    [CmdletBinding()]
    [OutputType([string])]
    param()
    $candidates = @()
    if ($env:PUBLIC) { $candidates += (Join-Path $env:PUBLIC 'Desktop') }
    $candidates += [Environment]::GetFolderPath('Desktop')
    if ($env:TEMP) { $candidates += $env:TEMP }
    foreach ($dir in $candidates) {
        if ([string]::IsNullOrEmpty($dir)) { continue }
        if (-not (Test-Path -LiteralPath $dir)) { continue }
        $path = Join-Path $dir 'OfficeLTSC_Install_Log.txt'
        try {
            [System.IO.File]::AppendAllText($path, '')
            return $path
        } catch {
            Write-Verbose "Log location not writable: $dir"
        }
    }
    $null
}

function Write-InstallLog {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)] [AllowEmptyString()] [string] $Message,
        [ValidateSet('Info', 'Ok', 'Warn', 'Error', 'Step')] [string] $Level = 'Info'
    )
    $colors = @{ Info = 'Gray'; Ok = 'Green'; Warn = 'Yellow'; Error = 'Red'; Step = 'Cyan' }
    $prefix = @{ Info = '   '; Ok = '[+]'; Warn = '[!]'; Error = '[x]'; Step = '==>' }
    $text = $Message
    if ($script:SecretKey) { $text = $text.Replace($script:SecretKey, (Get-MaskedKey $script:SecretKey)) }
    Write-Host ("{0} {1}" -f $prefix[$Level], $text) -ForegroundColor $colors[$Level]
    if ($script:LogPath) {
        $line = '{0} [{1}] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level.ToUpperInvariant(), $text
        try { [System.IO.File]::AppendAllText($script:LogPath, $line + [Environment]::NewLine) } catch { Write-Verbose 'Log write failed.' }
    }
}

# ---------------------------------------------------------------------------
# Pure helpers (unit tested)
# ---------------------------------------------------------------------------

function ConvertTo-NormalizedKey {
    [CmdletBinding()]
    [OutputType([string])]
    param([string] $Key)
    if ([string]::IsNullOrWhiteSpace($Key)) { return '' }
    ($Key -replace '\s', '').ToUpperInvariant()
}

function Test-ProductKeyFormat {
    [CmdletBinding()]
    [OutputType([bool])]
    param([string] $Key)
    $k = ConvertTo-NormalizedKey $Key
    $k -match '^[BCDFGHJKMNPQRTVWXY2346789]{5}(-[BCDFGHJKMNPQRTVWXY2346789]{5}){4}$'
}

function Resolve-LanguageId {
    # Returns the canonical (lower-case) culture id, 'MatchOS' / 'MatchInstalled',
    # or $null when the value is not a supported language.
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [string] $Language,
        [switch] $AllowSpecial
    )
    if ([string]::IsNullOrWhiteSpace($Language)) { return $null }
    $v = $Language.Trim()
    if ($AllowSpecial) {
        if ($v -ieq 'MatchOS') { return 'MatchOS' }
        if ($v -ieq 'MatchInstalled') { return 'MatchInstalled' }
    }
    foreach ($l in (Get-LanguageTable)) {
        if ($l.Id -ieq $v) { return $l.Id }
    }
    $null
}

function Find-Language {
    [CmdletBinding()]
    [OutputType([object[]])]
    param([string] $Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return @() }
    $t = $Text.Trim()
    @(Get-LanguageTable | Where-Object { $_.Id -like "*$t*" -or $_.Name -like "*$t*" })
}

function Split-ListValue {
    # Accepts arrays and/or comma separated strings ("Lync,Teams").
    [CmdletBinding()]
    [OutputType([string[]])]
    param([string[]] $Value)
    $out = New-Object System.Collections.Generic.List[string]
    foreach ($v in @($Value)) {
        if ($null -eq $v) { continue }
        foreach ($p in ($v -split '[,;\s]+')) {
            if ($p.Trim().Length -gt 0) { $out.Add($p.Trim()) }
        }
    }
    $out.ToArray()
}

function Resolve-ExcludeApp {
    [CmdletBinding()]
    [OutputType([string[]])]
    param([string[]] $App)
    $valid = Get-ExcludeAppList
    $out = New-Object System.Collections.Generic.List[string]
    foreach ($a in @(Split-ListValue $App)) {
        $match = $valid | Where-Object { $_ -ieq $a } | Select-Object -First 1
        if (-not $match) {
            throw (Get-InstallerError -Code 6 -Message ("Unknown app '{0}'. Valid: {1}" -f $a, ($valid -join ', ')))
        }
        if (-not $out.Contains($match)) { $out.Add($match) }
    }
    $out.ToArray()
}

function Test-MicrosoftSignature {
    # Accepts the object returned by Get-AuthenticodeSignature.
    [CmdletBinding()]
    [OutputType([bool])]
    param([Parameter(Mandatory = $true)] $Signature)
    if ($null -eq $Signature) { return $false }
    if ([string]$Signature.Status -ne 'Valid') { return $false }
    if ($null -eq $Signature.SignerCertificate) { return $false }
    [string]$Signature.SignerCertificate.Subject -match 'O=Microsoft Corporation(,|$)'
}

function Get-OdtConfigurationXml {
    # Builds configuration.xml with XmlWriter so every value is escaped.
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory = $true)] [ValidateSet('2024', '2021', '2019')] [string] $Version,
        [Parameter(Mandatory = $true)] [string] $Language,
        [string[]] $ExtraLanguages = @(),
        [ValidateSet('64', '32')] [string] $Arch = '64',
        [ValidateSet('ProPlus', 'Standard')] [string] $Edition = 'ProPlus',
        [switch] $Visio,
        [switch] $Project,
        [string[]] $ExcludeApps = @(),
        [string] $ProductKey,
        [switch] $RemoveExisting,
        [string] $SourcePath,
        [switch] $NoUpdates,
        [switch] $Silent
    )
    $catalog = (Get-ProductCatalog)[$Version]
    $languages = New-Object System.Collections.Generic.List[string]
    foreach ($l in @($Language) + @($ExtraLanguages)) {
        if ([string]::IsNullOrWhiteSpace($l)) { continue }
        if (-not $languages.Contains($l)) { $languages.Add($l) }
    }

    $settings = New-Object System.Xml.XmlWriterSettings
    $settings.Indent = $true
    $settings.IndentChars = '  '
    $settings.NewLineChars = "`r`n"
    $settings.OmitXmlDeclaration = $true
    $sb = New-Object System.Text.StringBuilder
    $w = [System.Xml.XmlWriter]::Create($sb, $settings)
    try {
        $w.WriteStartElement('Configuration')
        if ($RemoveExisting) {
            $w.WriteStartElement('Remove'); $w.WriteAttributeString('All', 'TRUE'); $w.WriteEndElement()
        }
        $w.WriteStartElement('Add')
        $w.WriteAttributeString('OfficeClientEdition', $Arch)
        $w.WriteAttributeString('Channel', $catalog.Channel)
        if (-not [string]::IsNullOrWhiteSpace($SourcePath)) { $w.WriteAttributeString('SourcePath', $SourcePath) }
        $w.WriteAttributeString('AllowCdnFallback', 'TRUE')

        $products = @(@{ Id = $catalog[$Edition]; Main = $true })
        if ($Visio) { $products += @{ Id = $catalog.Visio; Main = $false } }
        if ($Project) { $products += @{ Id = $catalog.Project; Main = $false } }
        foreach ($p in $products) {
            $w.WriteStartElement('Product')
            $w.WriteAttributeString('ID', $p.Id)
            if ($p.Main -and -not [string]::IsNullOrEmpty($ProductKey)) { $w.WriteAttributeString('PIDKEY', $ProductKey) }
            foreach ($l in $languages) {
                $w.WriteStartElement('Language'); $w.WriteAttributeString('ID', $l); $w.WriteEndElement()
            }
            if ($p.Main) {
                foreach ($a in @($ExcludeApps)) {
                    if ([string]::IsNullOrWhiteSpace($a)) { continue }
                    $w.WriteStartElement('ExcludeApp'); $w.WriteAttributeString('ID', $a); $w.WriteEndElement()
                }
            }
            $w.WriteEndElement()
        }
        $w.WriteEndElement() # Add

        if ($RemoveExisting) { $w.WriteStartElement('RemoveMSI'); $w.WriteEndElement() }
        if (-not [string]::IsNullOrEmpty($ProductKey)) {
            $w.WriteStartElement('Property'); $w.WriteAttributeString('Name', 'AUTOACTIVATE'); $w.WriteAttributeString('Value', '1'); $w.WriteEndElement()
        }
        $w.WriteStartElement('Property'); $w.WriteAttributeString('Name', 'FORCEAPPSHUTDOWN'); $w.WriteAttributeString('Value', 'TRUE'); $w.WriteEndElement()
        $w.WriteStartElement('Updates'); $w.WriteAttributeString('Enabled', $(if ($NoUpdates) { 'FALSE' } else { 'TRUE' })); $w.WriteEndElement()
        $w.WriteStartElement('Display'); $w.WriteAttributeString('Level', $(if ($Silent) { 'None' } else { 'Full' })); $w.WriteAttributeString('AcceptEULA', 'TRUE'); $w.WriteEndElement()
        $w.WriteEndElement() # Configuration
    } finally {
        $w.Flush()
        $w.Close()
    }
    $sb.ToString()
}

function Get-ActivationErrorText {
    [CmdletBinding()]
    [OutputType([string])]
    param([string] $Code)
    $map = @{
        '0xC004C003' = 'The product key is blocked by Microsoft.'
        '0xC004C008' = 'The activation limit for this key has been reached.'
        '0xC004C060' = 'The product key is no longer valid.'
        '0xC004C4CE' = 'The product key does not match the installed product.'
        '0xC004F074' = 'No KMS host could be reached. Check the KMS host / network.'
        '0xC004F038' = 'The KMS host count is too low (needs at least 5 clients for Office).'
        '0xC004F050' = 'The product key is invalid.'
        '0x8007232B' = 'DNS name does not exist (no KMS host found in DNS).'
        '0x80072EE7' = 'The server name could not be resolved (no internet / DNS).'
        '0x8004FE21' = 'The computer is not running genuine Windows.'
    }
    if ([string]::IsNullOrEmpty($Code)) { return 'Unknown error.' }
    $k = $Code.ToUpperInvariant().Replace('0X', '0x')
    if ($map.ContainsKey($k)) { return $map[$k] }
    'See https://learn.microsoft.com/ for this activation error code.'
}

# ---------------------------------------------------------------------------
# System helpers
# ---------------------------------------------------------------------------

function Test-IsAdministrator {
    [CmdletBinding()]
    [OutputType([bool])]
    param()
    $id = [Security.Principal.WindowsIdentity]::GetCurrent()
    (New-Object Security.Principal.WindowsPrincipal $id).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Get-SelfScriptText {
    [CmdletBinding()]
    [OutputType([string])]
    param()
    if ($script:SelfCommandPath -and (Test-Path -LiteralPath $script:SelfCommandPath)) {
        return [System.IO.File]::ReadAllText($script:SelfCommandPath)
    }
    $inv = $script:SelfInvocation
    if ($inv -and $inv.MyCommand) {
        $sb = $null
        try { $sb = $inv.MyCommand.ScriptBlock } catch { $sb = $null }
        if ($sb) {
            $text = $sb.ToString()
            if ($text -match 'function Get-InstallerConstant') { return $text }
        }
    }
    # Last resort (e.g. "irm | iex" where the text is not reachable): fetch it again.
    $url = (Get-InstallerConstant).SiteBase + '/install.txt'
    Write-InstallLog "Re-downloading the script from $url for the elevated window." 'Info'
    [string](Invoke-RestMethod -UseBasicParsing -Uri $url)
}

function Invoke-SelfElevation {
    # Re-launches this script elevated, passing the bound parameters through.
    [CmdletBinding()]
    param([hashtable] $BoundParameters)
    $dir = Join-Path $env:TEMP 'OfficeLTSC'
    if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    $file = Join-Path $dir 'install.ps1'
    $text = Get-SelfScriptText
    [System.IO.File]::WriteAllText($file, $text, (New-Object System.Text.UTF8Encoding $false))

    $argList = New-Object System.Collections.Generic.List[string]
    $argList.Add('-NoProfile'); $argList.Add('-ExecutionPolicy'); $argList.Add('Bypass')
    $argList.Add('-File'); $argList.Add('"' + $file + '"')
    foreach ($name in $BoundParameters.Keys) {
        if ($name -in @('ProductKey', 'KeyFile', 'Relaunched')) { continue }
        $value = $BoundParameters[$name]
        if ($value -is [System.Management.Automation.SwitchParameter]) {
            if ($value.IsPresent) { $argList.Add("-$name") }
        } elseif ($value -is [array]) {
            $argList.Add("-$name"); $argList.Add('"' + (($value | ForEach-Object { ([string]$_).Replace('"', '') }) -join ',') + '"')
        } else {
            $argList.Add("-$name"); $argList.Add('"' + ([string]$value).Replace('"', '') + '"')
        }
    }
    if ($BoundParameters.ContainsKey('ProductKey') -and $BoundParameters['ProductKey']) {
        # Never put the key on a command line: hand it over encrypted with DPAPI.
        $keyPath = Join-Path $dir ('k_' + [guid]::NewGuid().ToString('N') + '.dat')
        $secure = New-Object System.Security.SecureString
        foreach ($ch in ([string]$BoundParameters['ProductKey']).ToCharArray()) { $secure.AppendChar($ch) }
        [System.IO.File]::WriteAllText($keyPath, (ConvertFrom-SecureString -SecureString $secure))
        $argList.Add('-KeyFile'); $argList.Add('"' + $keyPath + '"')
    }
    $argList.Add('-Relaunched')
    $exe = (Get-Process -Id $PID).Path
    if (-not $exe) { $exe = 'powershell.exe' }
    Write-InstallLog 'Administrator rights are required. Opening an elevated window...' 'Step'
    try {
        Start-Process -FilePath $exe -Verb RunAs -ArgumentList ($argList -join ' ') | Out-Null
    } catch {
        throw (Get-InstallerError -Code 1 -Message 'Elevation was cancelled (UAC prompt declined).')
    }
    Write-InstallLog 'Continue in the new Administrator window.' 'Ok'
}

function Read-KeyFile {
    [CmdletBinding()]
    [OutputType([string])]
    param([string] $Path)
    if (-not $Path -or -not (Test-Path -LiteralPath $Path)) { return '' }
    try {
        $secure = ConvertTo-SecureString -String ([System.IO.File]::ReadAllText($Path))
        $bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($secure)
        try { return [Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr) }
        finally { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr) }
    } catch {
        Write-InstallLog 'Could not read the product key handed over from the first window. You will be asked again.' 'Warn'
        return ''
    } finally {
        Remove-Item -LiteralPath $Path -Force -ErrorAction SilentlyContinue
    }
}

function Get-InstalledOfficeInfo {
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param()
    $info = [ordered]@{ C2R = $false; Platform = ''; Products = ''; Msi = $false }
    $c2r = 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration'
    if (Test-Path $c2r) {
        $cfg = Get-ItemProperty -Path $c2r -ErrorAction SilentlyContinue
        if ($cfg) {
            $info.C2R = $true
            if ($cfg.PSObject.Properties['Platform']) { $info.Platform = [string]$cfg.Platform }
            if ($cfg.PSObject.Properties['ProductReleaseIds']) { $info.Products = [string]$cfg.ProductReleaseIds }
        }
    }
    foreach ($v in @('16.0', '15.0', '14.0', '12.0')) {
        foreach ($root in @('HKLM:\SOFTWARE\Microsoft\Office', 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Office')) {
            if (Test-Path "$root\$v\Common\InstallRoot") {
                $p = Get-ItemProperty -Path "$root\$v\Common\InstallRoot" -ErrorAction SilentlyContinue
                if ($p -and $p.PSObject.Properties['Path'] -and -not $info.C2R) { $info.Msi = $true }
            }
        }
    }
    [pscustomobject]$info
}

function Invoke-Download {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)] [string] $Uri,
        [Parameter(Mandatory = $true)] [string] $OutFile,
        [int] $Attempts = 3
    )
    for ($i = 1; $i -le $Attempts; $i++) {
        try {
            $old = $ProgressPreference
            $ProgressPreference = 'SilentlyContinue'
            try { Invoke-WebRequest -UseBasicParsing -Uri $Uri -OutFile $OutFile -TimeoutSec 120 }
            finally { $ProgressPreference = $old }
            if ((Test-Path -LiteralPath $OutFile) -and (Get-Item -LiteralPath $OutFile).Length -gt 0) { return $true }
        } catch {
            Write-InstallLog ("Download attempt {0}/{1} failed: {2}" -f $i, $Attempts, $_.Exception.Message) 'Warn'
        }
        if ($i -lt $Attempts) { Start-Sleep -Seconds (2 * $i) }
    }
    $false
}

function Wait-ProcessWithSpinner {
    [CmdletBinding()]
    [OutputType([int])]
    param(
        [Parameter(Mandatory = $true)] [System.Diagnostics.Process] $Process,
        [string] $Activity = 'Working'
    )
    $null = $Process.Handle # cache the handle so ExitCode is available later
    $frames = @('|', '/', '-', '\')
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $i = 0
    $canDraw = $true
    try { $null = [Console]::CursorLeft } catch { $canDraw = $false }
    while (-not $Process.HasExited) {
        if ($canDraw) {
            $msg = "    {0} {1}  elapsed {2:hh\:mm\:ss}   " -f $frames[$i % 4], $Activity, $sw.Elapsed
            Write-Host -NoNewline ("`r" + $msg)
        }
        $i++
        Start-Sleep -Milliseconds 500
    }
    if ($canDraw) { Write-Host '' }
    $Process.WaitForExit()
    Write-InstallLog ("{0} finished after {1:hh\:mm\:ss}." -f $Activity, $sw.Elapsed) 'Info'
    [int]$Process.ExitCode
}

# ---------------------------------------------------------------------------
# Interactive menu
# ---------------------------------------------------------------------------

function Read-Choice {
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory = $true)] [string] $Prompt,
        [Parameter(Mandatory = $true)] [string[]] $Choices,
        [Parameter(Mandatory = $true)] [string] $Default
    )
    while ($true) {
        $answer = Read-Host ("{0} [{1}] (default {2})" -f $Prompt, ($Choices -join '/'), $Default)
        if ([string]::IsNullOrWhiteSpace($answer)) { return $Default }
        $match = $Choices | Where-Object { $_ -ieq $answer.Trim() } | Select-Object -First 1
        if ($match) { return $match }
        Write-Host '    Invalid choice, try again.' -ForegroundColor Yellow
    }
}

function Read-YesNo {
    [CmdletBinding()]
    [OutputType([bool])]
    param([string] $Prompt, [bool] $Default = $false)
    $d = if ($Default) { 'Y' } else { 'N' }
    (Read-Choice -Prompt $Prompt -Choices @('Y', 'N') -Default $d) -eq 'Y'
}

function Read-LanguageInteractive {
    [CmdletBinding()]
    [OutputType([string])]
    param([string] $Prompt = 'Language', [string] $Default = 'MatchOS', [switch] $AllowSpecial, [switch] $AllowEmpty)
    while ($true) {
        $q = Read-Host ("{0}: type part of a name or code (e.g. arab, en-us, french){1}" -f $Prompt, $(if ($AllowEmpty) { ', Enter = none' } else { ", Enter = $Default" }))
        if ([string]::IsNullOrWhiteSpace($q)) { if ($AllowEmpty) { return '' } else { return $Default } }
        $exact = Resolve-LanguageId $q -AllowSpecial:$AllowSpecial
        if ($exact) { return $exact }
        $hits = @(Find-Language $q)
        if ($hits.Count -eq 0) { Write-Host '    No match. Try again.' -ForegroundColor Yellow; continue }
        if ($hits.Count -eq 1) { Write-Host ("    -> {0} ({1})" -f $hits[0].Name, $hits[0].Id); return $hits[0].Id }
        for ($i = 0; $i -lt [Math]::Min($hits.Count, 20); $i++) {
            Write-Host ("    {0,2}. {1,-16} {2}" -f ($i + 1), $hits[$i].Id, $hits[$i].Name)
        }
        $n = Read-Host '    Number (Enter to search again)'
        $num = 0
        if ([int]::TryParse($n, [ref]$num) -and $num -ge 1 -and $num -le [Math]::Min($hits.Count, 20)) { return $hits[$num - 1].Id }
    }
}

function Read-InteractiveOption {
    [CmdletBinding()]
    [OutputType([hashtable])]
    param([hashtable] $Current)
    $o = $Current
    Write-Host ''
    Write-Host '  Answer each question, or press Enter to accept the default.' -ForegroundColor Gray
    Write-Host ''
    Write-Host '  1) Version' -ForegroundColor Cyan
    $o.Version = Read-Choice -Prompt '     Office LTSC version' -Choices @('2024', '2021', '2019') -Default '2024'
    Write-Host '  2) Language' -ForegroundColor Cyan
    $o.Language = Read-LanguageInteractive -Prompt '     Language' -Default 'MatchOS' -AllowSpecial
    $extra = New-Object System.Collections.Generic.List[string]
    while (Read-YesNo -Prompt '     Add another language?' -Default $false) {
        $x = Read-LanguageInteractive -Prompt '     Extra language' -AllowEmpty
        if ($x -and -not $extra.Contains($x)) { $extra.Add($x) }
    }
    $o.ExtraLanguages = $extra.ToArray()
    Write-Host '  3) Architecture' -ForegroundColor Cyan
    $o.Arch = Read-Choice -Prompt '     Bits' -Choices @('64', '32') -Default '64'
    Write-Host '  4) Edition' -ForegroundColor Cyan
    $o.Edition = Read-Choice -Prompt '     Edition' -Choices @('ProPlus', 'Standard') -Default 'ProPlus'
    $o.Visio = Read-YesNo -Prompt '     Also install Visio Professional?' -Default $false
    $o.Project = Read-YesNo -Prompt '     Also install Project Professional?' -Default $false
    Write-Host '  5) Apps to exclude' -ForegroundColor Cyan
    $apps = @('Access', 'Excel', 'Lync', 'OneDrive', 'OneNote', 'Outlook', 'PowerPoint', 'Publisher', 'Teams', 'Word')
    for ($i = 0; $i -lt $apps.Count; $i++) { Write-Host ("     {0,2}. {1}" -f ($i + 1), $apps[$i]) }
    while ($true) {
        $sel = Read-Host '     Numbers separated by commas (e.g. 3,9), Enter = none'
        if ([string]::IsNullOrWhiteSpace($sel)) { $o.ExcludeApps = @(); break }
        $picked = New-Object System.Collections.Generic.List[string]
        $ok = $true
        foreach ($s in @(Split-ListValue $sel)) {
            $n = 0
            if ([int]::TryParse($s, [ref]$n) -and $n -ge 1 -and $n -le $apps.Count) {
                if (-not $picked.Contains($apps[$n - 1])) { $picked.Add($apps[$n - 1]) }
            } else { $ok = $false }
        }
        if ($ok) { $o.ExcludeApps = $picked.ToArray(); break }
        Write-Host '     Invalid list, try again.' -ForegroundColor Yellow
    }
    Write-Host '  6) Existing Office' -ForegroundColor Cyan
    Write-Host '     N = keep existing Office (default)'
    Write-Host '     Y = let the Office Deployment Tool remove all Office versions first (Remove All + RemoveMSI)'
    Write-Host '     D = deep clean: run OfficeScrubber first (removes stubborn/broken installs and Office licenses), then the above'
    $rm = Read-Choice -Prompt '     Remove existing Office?' -Choices @('N', 'Y', 'D') -Default 'N'
    $o.RemoveExisting = ($rm -ne 'N')
    $o.Scrub = ($rm -eq 'D')
    Write-Host '  7) Product key (optional)' -ForegroundColor Cyan
    Write-Host '     Your own MAK / KMS / retail key. Leave empty to install without a key.'
    while ($true) {
        $secure = Read-Host -AsSecureString '     Product key (input hidden, Enter = none)'
        $bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($secure)
        try { $plain = [Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr) } finally { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr) }
        $plain = ConvertTo-NormalizedKey $plain
        if ($plain -eq '') { $o.ProductKey = ''; break }
        if (Test-ProductKeyFormat $plain) { $o.ProductKey = $plain; break }
        Write-Host '     That does not look like a valid key (XXXXX-XXXXX-XXXXX-XXXXX-XXXXX). Try again.' -ForegroundColor Yellow
    }
    $o
}

function Show-Summary {
    [CmdletBinding()]
    param([hashtable] $o, [hashtable] $Catalog)
    $products = @($Catalog[$o.Edition])
    if ($o.Visio) { $products += $Catalog.Visio }
    if ($o.Project) { $products += $Catalog.Project }
    $remove = if ($o.Scrub) { 'Yes (OfficeScrubber + ODT)' } elseif ($o.RemoveExisting) { 'Yes (ODT)' } else { 'No' }
    Write-Host ''
    Write-Host '  ------------------------- Summary -------------------------' -ForegroundColor Cyan
    Write-Host ("  Version        : Office LTSC {0} ({1})" -f $o.Version, $Catalog.Channel)
    Write-Host ("  Products       : {0}" -f ($products -join ', '))
    Write-Host ("  Language(s)    : {0}" -f ((@($o.Language) + @($o.ExtraLanguages) | Where-Object { $_ }) -join ', '))
    Write-Host ("  Architecture   : {0}-bit" -f $o.Arch)
    Write-Host ("  Excluded apps  : {0}" -f $(if (@($o.ExcludeApps).Count) { $o.ExcludeApps -join ', ' } else { 'none' }))
    Write-Host ("  Remove existing: {0}" -f $remove)
    Write-Host ("  Product key    : {0}" -f (Get-MaskedKey $o.ProductKey))
    if ($o.DownloadOnly) { Write-Host ("  Mode           : download only to {0}" -f $o.SourcePath) }
    elseif ($o.SourcePath) { Write-Host ("  Source         : {0}" -f $o.SourcePath) }
    Write-Host '  -----------------------------------------------------------' -ForegroundColor Cyan
    Write-Host ''
}

# ---------------------------------------------------------------------------
# Main steps
# ---------------------------------------------------------------------------

function Test-Environment {
    [CmdletBinding()]
    param([hashtable] $o)
    $c = Get-InstallerConstant
    Write-InstallLog 'Checking the system...' 'Step'
    if ([Environment]::OSVersion.Platform -ne [PlatformID]::Win32NT) {
        throw (Get-InstallerError -Code 2 -Message 'This script runs on Windows only.')
    }
    $nt = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' -ErrorAction SilentlyContinue
    $build = 0
    if ($nt -and $nt.PSObject.Properties['CurrentBuildNumber']) { $build = [int]$nt.CurrentBuildNumber }
    $osName = if ($nt -and $nt.PSObject.Properties['ProductName']) { [string]$nt.ProductName } else { 'Windows' }
    if ($build -lt $c.MinBuild) {
        throw (Get-InstallerError -Code 2 -Message ("Windows 10 1809 / Server 2019 (build {0}) or newer is required. Found build {1}." -f $c.MinBuild, $build))
    }
    Write-InstallLog ("{0}, build {1}" -f $osName, $build) 'Ok'
    if ($o.Arch -eq '64' -and -not [Environment]::Is64BitOperatingSystem) {
        throw (Get-InstallerError -Code 2 -Message '64-bit Office needs a 64-bit Windows. Use -Arch 32.')
    }
    $target = $env:SystemDrive
    if ($o.DownloadOnly -and $o.SourcePath) { $target = [System.IO.Path]::GetPathRoot([System.IO.Path]::GetFullPath($o.SourcePath)) }
    $drive = New-Object System.IO.DriveInfo ($target)
    $freeGB = [Math]::Round($drive.AvailableFreeSpace / 1GB, 1)
    if ($freeGB -lt $c.MinFreeGB) {
        throw (Get-InstallerError -Code 2 -Message ("At least {0} GB free space is needed on {1}. Free: {2} GB." -f $c.MinFreeGB, $drive.Name, $freeGB))
    }
    Write-InstallLog ("Free space on {0} {1} GB" -f $drive.Name, $freeGB) 'Ok'
    if (-not $o.SourcePath -or $o.DownloadOnly) {
        try {
            $old = $ProgressPreference; $ProgressPreference = 'SilentlyContinue'
            try { $null = Invoke-WebRequest -UseBasicParsing -Method Head -Uri $c.OdtUrl -TimeoutSec 30 }
            finally { $ProgressPreference = $old }
            Write-InstallLog 'officecdn.microsoft.com is reachable' 'Ok'
        } catch {
            $headError = $_.Exception.Message
            $resolved = $false
            try { $resolved = @([Net.Dns]::GetHostAddresses('officecdn.microsoft.com')).Count -gt 0 } catch { $resolved = $false }
            if (-not $resolved) {
                throw (Get-InstallerError -Code 2 -Message ("Cannot reach officecdn.microsoft.com: {0}" -f $headError))
            }
            Write-InstallLog ("Connectivity test was inconclusive ({0}); continuing." -f $headError) 'Warn'
        }
    }
    $installed = Get-InstalledOfficeInfo
    if ($installed.C2R) {
        Write-InstallLog ("Installed Click-to-Run Office: {0} ({1})" -f $installed.Products, $installed.Platform) 'Info'
        $want = if ($o.Arch -eq '64') { 'x64' } else { 'x86' }
        if ($installed.Platform -and $installed.Platform -ne $want -and -not $o.RemoveExisting) {
            Write-InstallLog ("Installed Office is {0} but {1} was requested. Setup will fail unless you remove it first (-RemoveExisting)." -f $installed.Platform, $want) 'Warn'
        }
    }
    if ($installed.Msi -and -not $o.RemoveExisting) {
        Write-InstallLog 'An MSI-based Office was found. Consider -RemoveExisting to remove it first.' 'Warn'
    }
}

function Get-VerifiedOdt {
    [CmdletBinding()]
    [OutputType([string])]
    param([string] $WorkDir)
    $c = Get-InstallerConstant
    $setup = Join-Path $WorkDir 'setup.exe'
    Write-InstallLog ("Downloading the Office Deployment Tool from {0}" -f $c.OdtUrl) 'Step'
    if (-not (Invoke-Download -Uri $c.OdtUrl -OutFile $setup)) {
        throw (Get-InstallerError -Code 3 -Message 'Could not download setup.exe from officecdn.microsoft.com.')
    }
    $sig = Get-AuthenticodeSignature -FilePath $setup
    if (-not (Test-MicrosoftSignature $sig)) {
        $subject = if ($sig.SignerCertificate) { $sig.SignerCertificate.Subject } else { '(none)' }
        Remove-Item -LiteralPath $setup -Force -ErrorAction SilentlyContinue
        throw (Get-InstallerError -Code 3 -Message ("setup.exe signature check FAILED (status: {0}, signer: {1}). File deleted." -f $sig.Status, $subject))
    }
    $ver = (Get-Item -LiteralPath $setup).VersionInfo.FileVersion
    Write-InstallLog ("setup.exe {0} - signature Valid, signed by Microsoft Corporation" -f $ver) 'Ok'
    $setup
}

function Invoke-OfficeScrubber {
    [CmdletBinding()]
    param([string] $WorkDir)
    $c = Get-InstallerConstant
    $dir = Join-Path $WorkDir 'OfficeScrubber'
    New-Item -ItemType Directory -Path (Join-Path $dir 'bin') -Force | Out-Null
    $base = 'https://raw.githubusercontent.com/abbodi1406/BatUtil/{0}/OfficeScrubber/' -f $c.ScrubberCommit
    Write-InstallLog ("Downloading OfficeScrubber (abbodi1406/BatUtil @ {0})" -f $c.ScrubberCommit.Substring(0, 12)) 'Step'
    foreach ($rel in $c.ScrubberFiles.Keys) {
        $out = Join-Path $dir ($rel.Replace('/', '\'))
        if (-not (Invoke-Download -Uri ($base + $rel) -OutFile $out)) {
            throw (Get-InstallerError -Code 3 -Message "Could not download OfficeScrubber file $rel.")
        }
        $hash = (Get-FileHash -LiteralPath $out -Algorithm SHA256).Hash.ToLowerInvariant()
        if ($hash -ne $c.ScrubberFiles[$rel]) {
            Remove-Item -LiteralPath $dir -Recurse -Force -ErrorAction SilentlyContinue
            throw (Get-InstallerError -Code 3 -Message "OfficeScrubber file $rel failed the SHA-256 check. Aborting.")
        }
    }
    Write-InstallLog 'OfficeScrubber files verified (SHA-256)' 'Ok'
    Write-InstallLog 'Running OfficeScrubber (Scrub ALL). This can take 10-30 minutes...' 'Step'
    # Always use the native 64-bit cmd.exe; -wow/-arm/-qedit stop the script from
    # re-launching itself in a detached window so we can wait for it.
    $cmd = Join-Path $env:SystemRoot 'System32\cmd.exe'
    $sysnative = Join-Path $env:SystemRoot 'Sysnative\cmd.exe'
    if (Test-Path -LiteralPath $sysnative) { $cmd = $sysnative }
    $scrubCmd = Join-Path $dir 'OfficeScrubber.cmd'
    $p = Start-Process -FilePath $cmd -ArgumentList ('/c ""' + $scrubCmd + '" /A -wow -arm -qedit"') -WorkingDirectory $dir -PassThru
    $code = Wait-ProcessWithSpinner -Process $p -Activity 'OfficeScrubber'
    Write-InstallLog ("OfficeScrubber exit code {0}" -f $code) 'Info'
}

function Invoke-Odt {
    [CmdletBinding()]
    [OutputType([int])]
    param([string] $Setup, [string] $ConfigPath, [string] $Mode)
    Write-InstallLog ("Running setup.exe /{0} configuration.xml" -f $Mode) 'Step'
    if ($Mode -eq 'configure') {
        Write-InstallLog 'Office is being downloaded from Microsoft and installed. This usually takes 10-40 minutes.' 'Info'
    }
    $p = Start-Process -FilePath $Setup -ArgumentList ('/{0} "{1}"' -f $Mode, $ConfigPath) -WorkingDirectory (Split-Path $Setup) -PassThru
    Wait-ProcessWithSpinner -Process $p -Activity ("setup.exe /{0}" -f $Mode)
}

function Get-OsppPath {
    [CmdletBinding()]
    [OutputType([string])]
    param()
    $candidates = @()
    if ($env:ProgramFiles) { $candidates += (Join-Path $env:ProgramFiles 'Microsoft Office\Office16\OSPP.VBS') }
    if ($env:ProgramW6432) { $candidates += (Join-Path $env:ProgramW6432 'Microsoft Office\Office16\OSPP.VBS') }
    $x86 = [Environment]::GetEnvironmentVariable('ProgramFiles(x86)')
    if ($x86) { $candidates += (Join-Path $x86 'Microsoft Office\Office16\OSPP.VBS') }
    foreach ($c in $candidates) { if (Test-Path -LiteralPath $c) { return $c } }
    $null
}

function Get-ActivationState {
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param([string] $Ospp, [string] $Last5)
    $cscript = Join-Path $env:SystemRoot 'System32\cscript.exe'
    $out = & $cscript //NoLogo $Ospp /dstatus 2>&1 | Out-String
    $blocks = $out -split '-{20,}'
    foreach ($b in $blocks) {
        if ($Last5 -and $b -notmatch [regex]::Escape($Last5)) { continue }
        if ($b -match 'LICENSE STATUS:\s*-*([A-Z ]+?)-*\s*(\r?\n|$)') {
            $status = $Matches[1].Trim()
            $err = ''
            if ($b -match 'ERROR CODE:\s*(0x[0-9A-Fa-f]+)') { $err = $Matches[1] }
            $name = ''
            if ($b -match 'LICENSE NAME:\s*(.+)') { $name = $Matches[1].Trim() }
            return [pscustomobject]@{ Found = $true; Licensed = ($status -eq 'LICENSED'); Status = $status; Error = $err; Name = $name }
        }
    }
    [pscustomobject]@{ Found = $false; Licensed = $false; Status = 'not found'; Error = ''; Name = '' }
}

function Test-Activation {
    [CmdletBinding()]
    [OutputType([bool])]
    param([string] $Key)
    Write-InstallLog 'Checking activation...' 'Step'
    $ospp = Get-OsppPath
    if (-not $ospp) {
        Write-InstallLog 'OSPP.VBS not found; cannot check activation.' 'Warn'
        return $false
    }
    $last5 = $Key.Substring($Key.Length - 5)
    $state = Get-ActivationState -Ospp $ospp -Last5 $last5
    if (-not $state.Licensed) {
        Write-InstallLog 'Not activated yet; asking Office to activate now (OSPP.VBS /act)...' 'Info'
        $cscript = Join-Path $env:SystemRoot 'System32\cscript.exe'
        $null = & $cscript //NoLogo $ospp /act 2>&1
        Start-Sleep -Seconds 3
        $state = Get-ActivationState -Ospp $ospp -Last5 $last5
    }
    if ($state.Licensed) {
        Write-InstallLog ("Activated: {0}" -f $state.Name) 'Ok'
        return $true
    }
    Write-InstallLog ("Not activated. Status: {0}. Error: {1} - {2}" -f $state.Status, $(if ($state.Error) { $state.Error } else { 'n/a' }), (Get-ActivationErrorText $state.Error)) 'Warn'
    Write-InstallLog ("You can retry later with: cscript ""{0}"" /act" -f $ospp) 'Info'
    $false
}

function Invoke-Main {
    [CmdletBinding()]
    [OutputType([int])]
    param(
        [hashtable] $Bound,
        [hashtable] $Options
    )

    # Both settings are local to this function (and what it calls), so they do
    # not leak into the caller's session when the script runs through iex.
    Set-StrictMode -Version Latest
    $ErrorActionPreference = 'Stop'

    $c = Get-InstallerConstant
    $script:LogPath = Get-LogPath
    $interactive = $true
    foreach ($k in $Bound.Keys) { if ($k -notin @('Relaunched', 'KeyFile', 'Verbose', 'Debug')) { $interactive = $false } }

    Write-Host ''
    Write-Host ('  {0}' -f $c.ProductName) -ForegroundColor Cyan
    Write-Host '  Independent tool. Not affiliated with Microsoft.' -ForegroundColor DarkGray
    Write-Host '  Office is downloaded directly from Microsoft servers.' -ForegroundColor DarkGray
    Write-Host ''
    Write-InstallLog ('==== {0} started (PowerShell {1}) ====' -f $c.ProductName, $PSVersionTable.PSVersion) 'Info'

    # TLS 1.2 for Windows PowerShell 5.1
    try { [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12 } catch { Write-Verbose 'TLS setting unchanged.' }

    # 1. Elevation
    if (-not (Test-IsAdministrator)) {
        Invoke-SelfElevation -BoundParameters $Bound
        return 0
    }

    $o = @{
        Version = $Options.Version; Language = $Options.Language; ExtraLanguages = @(Split-ListValue $Options.ExtraLanguages)
        Arch = $Options.Arch; Edition = $Options.Edition; Visio = [bool]$Options.Visio; Project = [bool]$Options.Project
        ExcludeApps = @(Split-ListValue $Options.ExcludeApps); ProductKey = ''
        RemoveExisting = [bool]$Options.RemoveExisting; Scrub = [bool]$Options.Scrub; SourcePath = $Options.SourcePath
        DownloadOnly = [bool]$Options.DownloadOnly; NoUpdates = [bool]$Options.NoUpdates; Silent = [bool]$Options.Silent
    }
    if ($Options.ProductKey) { $o.ProductKey = ConvertTo-NormalizedKey $Options.ProductKey }
    elseif ($Options.KeyFile) { $o.ProductKey = ConvertTo-NormalizedKey (Read-KeyFile $Options.KeyFile) }

    if ($interactive) {
        $o = Read-InteractiveOption -Current $o
    }

    # Validation
    if (-not $o.Version) { throw (Get-InstallerError -Code 6 -Message 'Missing -Version (2024, 2021 or 2019).') }
    if (-not $o.Language) { $o.Language = 'MatchOS' }
    $lang = Resolve-LanguageId $o.Language -AllowSpecial
    if (-not $lang) { throw (Get-InstallerError -Code 6 -Message ("Unknown language '{0}'. Use a culture code such as ar-sa, en-us, fr-fr, or MatchOS." -f $o.Language)) }
    $o.Language = $lang
    $extra = @()
    foreach ($x in @($o.ExtraLanguages)) {
        $r = Resolve-LanguageId $x
        if (-not $r) { throw (Get-InstallerError -Code 6 -Message ("Unknown extra language '{0}'." -f $x)) }
        if ($r -ne $lang -and $extra -notcontains $r) { $extra += $r }
    }
    $o.ExtraLanguages = $extra
    $o.ExcludeApps = @(Resolve-ExcludeApp $o.ExcludeApps)
    if ($o.ProductKey) {
        if (-not (Test-ProductKeyFormat $o.ProductKey)) { throw (Get-InstallerError -Code 6 -Message 'The product key format is invalid (expected XXXXX-XXXXX-XXXXX-XXXXX-XXXXX).') }
        $script:SecretKey = $o.ProductKey
    }
    if ($o.Scrub) { $o.RemoveExisting = $true }
    if ($o.DownloadOnly) {
        if (-not $o.SourcePath) { throw (Get-InstallerError -Code 6 -Message '-DownloadOnly needs -SourcePath (the folder to download Office into).') }
        $o.SourcePath = [System.IO.Path]::GetFullPath($o.SourcePath)
        if (-not (Test-Path -LiteralPath $o.SourcePath)) { New-Item -ItemType Directory -Path $o.SourcePath -Force | Out-Null }
    } elseif ($o.SourcePath) {
        $o.SourcePath = [System.IO.Path]::GetFullPath($o.SourcePath)
        if (-not (Test-Path -LiteralPath (Join-Path $o.SourcePath 'Office\Data'))) {
            Write-InstallLog ("{0} has no Office\Data folder; ODT will fall back to the Microsoft CDN." -f $o.SourcePath) 'Warn'
        }
    }
    $catalog = (Get-ProductCatalog)[$o.Version]

    Show-Summary -o $o -Catalog $catalog
    Write-InstallLog ("Options: version={0} lang={1} extra={2} arch={3} edition={4} visio={5} project={6} exclude={7} remove={8} scrub={9} key={10} source={11} downloadOnly={12}" -f `
            $o.Version, $o.Language, ($o.ExtraLanguages -join ','), $o.Arch, $o.Edition, $o.Visio, $o.Project, ($o.ExcludeApps -join ','), $o.RemoveExisting, $o.Scrub, (Get-MaskedKey $o.ProductKey), $o.SourcePath, $o.DownloadOnly) 'Info'
    if ($interactive -and -not (Read-YesNo -Prompt '  Start now?' -Default $true)) {
        throw (Get-InstallerError -Code 1 -Message 'Cancelled by user.')
    }

    # 2. Environment
    Test-Environment -o $o

    # 4. Work directory
    $work = Join-Path $env:TEMP ('OfficeLTSC_' + (Get-Date -Format 'yyyyMMdd_HHmmss'))
    New-Item -ItemType Directory -Path $work -Force | Out-Null
    $script:WorkDir = $work
    Write-InstallLog ("Work folder: {0}" -f $work) 'Info'

    # 5-6. ODT
    $setup = Get-VerifiedOdt -WorkDir $work

    # 8. configuration.xml
    $xmlArgs = @{
        Version = $o.Version; Language = $o.Language; ExtraLanguages = $o.ExtraLanguages; Arch = $o.Arch; Edition = $o.Edition
        Visio = $o.Visio; Project = $o.Project; ExcludeApps = $o.ExcludeApps; ProductKey = $o.ProductKey
        RemoveExisting = ($o.RemoveExisting -and -not $o.DownloadOnly); SourcePath = $o.SourcePath; NoUpdates = $o.NoUpdates; Silent = $o.Silent
    }
    $xml = Get-OdtConfigurationXml @xmlArgs
    $configPath = Join-Path $work 'configuration.xml'
    [System.IO.File]::WriteAllText($configPath, $xml, (New-Object System.Text.UTF8Encoding $false))
    $script:ConfigPath = $configPath
    Write-InstallLog 'configuration.xml:' 'Info'
    foreach ($line in ($xml -split "`r?`n")) { Write-InstallLog ('  ' + $line) 'Info' }

    # Optional deep clean
    if ($o.Scrub -and -not $o.DownloadOnly) { Invoke-OfficeScrubber -WorkDir $work }

    # 9. Run ODT
    $mode = if ($o.DownloadOnly) { 'download' } else { 'configure' }
    $exit = Invoke-Odt -Setup $setup -ConfigPath $configPath -Mode $mode
    if ($exit -ne 0) {
        throw (Get-InstallerError -Code 4 -Message ("setup.exe /{0} failed with exit code {1}." -f $mode, $exit))
    }
    if ($o.DownloadOnly) {
        Copy-Item -LiteralPath $setup -Destination (Join-Path $o.SourcePath 'setup.exe') -Force
        $offline = $xml
        if ($o.ProductKey) { $offline = $offline.Replace($o.ProductKey, '') -replace ' PIDKEY=""', '' }
        $offline = $offline -replace '\s*<Property Name="AUTOACTIVATE" Value="1" />', ''
        [System.IO.File]::WriteAllText((Join-Path $o.SourcePath 'configuration.xml'), $offline, (New-Object System.Text.UTF8Encoding $false))
        Write-InstallLog ("Office files downloaded to {0}. Install later with: setup.exe /configure configuration.xml (run in that folder)." -f $o.SourcePath) 'Ok'
        Remove-Item -LiteralPath $work -Recurse -Force -ErrorAction SilentlyContinue
        return 0
    }
    Write-InstallLog ("Office LTSC {0} installed." -f $o.Version) 'Ok'

    # 10. Activation
    $result = 0
    if ($o.ProductKey) {
        if (-not (Test-Activation -Key $o.ProductKey)) { $result = 5 }
    } else {
        Write-InstallLog 'No product key was given. Enter your key in any Office app (File > Account) to activate.' 'Info'
    }

    # 11. Cleanup
    Remove-Item -LiteralPath $work -Recurse -Force -ErrorAction SilentlyContinue
    $script:WorkDir = $null
    $result
}

# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

$script:LogPath = $null
$script:SecretKey = $null
$script:WorkDir = $null
$script:ConfigPath = $null
$exitCode = 0
$boundCopy = @{}
$pbp = Get-Variable -Name PSBoundParameters -ValueOnly -ErrorAction SilentlyContinue
if ($pbp) { foreach ($k in @($pbp.Keys)) { $boundCopy[$k] = $pbp[$k] } }
try {
    # Script parameters are read here, at top level, and handed to Invoke-Main.
    $scriptOptions = @{
        Version = $Version; Language = $Language; ExtraLanguages = $ExtraLanguages; Arch = $Arch; Edition = $Edition
        Visio = $Visio; Project = $Project; ExcludeApps = $ExcludeApps; ProductKey = $ProductKey
        RemoveExisting = $RemoveExisting; Scrub = $Scrub; SourcePath = $SourcePath; DownloadOnly = $DownloadOnly
        NoUpdates = $NoUpdates; Silent = $Silent; KeyFile = $KeyFile
    }
    $exitCode = Invoke-Main -Bound $boundCopy -Options $scriptOptions
} catch {
    $exitCode = 4
    if ($_.Exception.Data.Contains('OfficeLtscExitCode')) { $exitCode = [int]$_.Exception.Data['OfficeLtscExitCode'] }
    Write-InstallLog $_.Exception.Message 'Error'
    if ($exitCode -eq 4 -and -not $_.Exception.Data.Contains('OfficeLtscExitCode')) {
        Write-InstallLog ($_.ScriptStackTrace) 'Info'
    }
    if ($script:WorkDir -and (Test-Path -LiteralPath $script:WorkDir)) {
        if ($script:ConfigPath -and $script:SecretKey -and (Test-Path -LiteralPath $script:ConfigPath)) {
            # Do not leave the key on disk.
            $t = [System.IO.File]::ReadAllText($script:ConfigPath).Replace($script:SecretKey, (Get-MaskedKey $script:SecretKey))
            [System.IO.File]::WriteAllText($script:ConfigPath, $t)
        }
        Write-InstallLog ("Work folder kept for troubleshooting: {0}" -f $script:WorkDir) 'Info'
        Write-InstallLog ("ODT logs: {0}\{1}-*.log and {2}\Temp\{1}-*.log" -f $env:TEMP, $env:COMPUTERNAME, $env:SystemRoot) 'Info'
    }
} finally {
    Write-InstallLog ("==== Finished with exit code {0} ====" -f $exitCode) 'Info'
    if ($script:LogPath) { Write-Host ("    Log: {0}" -f $script:LogPath) -ForegroundColor DarkGray }
    $script:SecretKey = $null
}

if ($Relaunched) {
    Write-Host ''
    $null = Read-Host 'Press Enter to close this window'
}
if ($script:SelfCommandPath) { exit $exitCode }
# Running through "irm | iex" or a script block: do not close the caller's window.
Set-Variable -Name LASTEXITCODE -Value $exitCode -Scope Global
