#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }
# Unit tests for the pure functions in src/install.ps1.
# The script is never executed here: its function definitions are extracted
# from the AST and loaded, so nothing is downloaded or installed.

BeforeAll {
    $script:Root = Split-Path -Parent $PSScriptRoot
    $script:ScriptPath = Join-Path $Root 'src/install.ps1'
    $tokens = $null
    $parseErrors = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($ScriptPath, [ref]$tokens, [ref]$parseErrors)
    $functions = $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $false)
    foreach ($f in $functions) { . ([scriptblock]::Create($f.Extent.Text)) }

    function script:Get-Golden([string] $Name) {
        [xml](Get-Content -LiteralPath (Join-Path $Root "tests/golden/$Name") -Raw)
    }
}

Describe 'Script file' {
    It 'parses without errors' {
        $text = Get-Content -LiteralPath $ScriptPath -Raw
        { [scriptblock]::Create($text) } | Should -Not -Throw
    }

    It 'is pure ASCII without a BOM' {
        $bytes = [System.IO.File]::ReadAllBytes($ScriptPath)
        ($bytes | Where-Object { $_ -gt 0x7F }).Count | Should -Be 0
    }

    It 'exposes the documented parameters' {
        $params = (Get-Command -Name $ScriptPath).Parameters.Keys
        foreach ($p in 'Version', 'Language', 'ExtraLanguages', 'Arch', 'Edition', 'Visio', 'Project', 'ExcludeApps',
            'ProductKey', 'RemoveExisting', 'Scrub', 'SourcePath', 'DownloadOnly', 'NoUpdates', 'Silent') {
            $params | Should -Contain $p
        }
    }

    It 'rejects an unknown -Version at binding time' {
        { & $ScriptPath -Version 2018 } | Should -Throw '*does not belong to the set*'
    }

    It 'rejects an unknown -Arch at binding time' {
        { & $ScriptPath -Arch 128 } | Should -Throw '*does not belong to the set*'
    }

    It 'does not use aliases for common cmdlets' {
        $tokens = $null; $errs = $null
        $ast = [System.Management.Automation.Language.Parser]::ParseFile($ScriptPath, [ref]$tokens, [ref]$errs)
        $commands = $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.CommandAst] }, $true) |
            ForEach-Object { $_.GetCommandName() } | Where-Object { $_ }
        $commands | Should -Not -Contain 'irm'
        $commands | Should -Not -Contain 'iwr'
        $commands | Should -Not -Contain 'iex'
        $commands | Should -Not -Contain '%'
        $commands | Should -Not -Contain '?'
    }
}

Describe 'Product key validation' {
    It 'accepts a well-formed key: <key>' -TestCases @(
        @{ key = 'NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP' }
        @{ key = 'nmmkj-6rk4f-kmjvx-8d9mj-6mwkp' }
        @{ key = ' NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP ' }
    ) {
        Test-ProductKeyFormat $key | Should -BeTrue
    }

    It 'rejects a malformed key: <key>' -TestCases @(
        @{ key = '' }
        @{ key = 'NMMKJ-6RK4F-KMJVX-8D9MJ' }
        @{ key = 'NMMKJ6RK4FKMJVX8D9MJ6MWKP' }
        @{ key = 'AAAAA-6RK4F-KMJVX-8D9MJ-6MWKP' }
        @{ key = 'NMMKJ-6RK4F-KMJVX-8D9MJ-6MWK0' }
        @{ key = 'NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP-XXXXX' }
    ) {
        Test-ProductKeyFormat $key | Should -BeFalse
    }

    It 'masks all but the last 5 characters' {
        $masked = Get-MaskedKey 'NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP'
        $masked | Should -Be 'XXXXX-XXXXX-XXXXX-XXXXX-6MWKP'
        $masked | Should -Not -Match 'NMMKJ'
        Get-MaskedKey '' | Should -Be '(none)'
    }
}

Describe 'Language validation' {
    It 'normalises <in> to <out>' -TestCases @(
        @{ in = 'ar-sa'; out = 'ar-sa' }
        @{ in = 'AR-SA'; out = 'ar-sa' }
        @{ in = 'sr-Latn-RS'; out = 'sr-latn-rs' }
        @{ in = 'ca-ES-valencia'; out = 'ca-es-valencia' }
    ) {
        Resolve-LanguageId $in | Should -Be $out
    }

    It 'allows MatchOS only with -AllowSpecial' {
        Resolve-LanguageId 'matchos' -AllowSpecial | Should -Be 'MatchOS'
        Resolve-LanguageId 'MatchInstalled' -AllowSpecial | Should -Be 'MatchInstalled'
        Resolve-LanguageId 'MatchOS' | Should -BeNullOrEmpty
    }

    It 'rejects unknown languages' {
        Resolve-LanguageId 'xx-yy' | Should -BeNullOrEmpty
        Resolve-LanguageId '' | Should -BeNullOrEmpty
        Resolve-LanguageId 'ar-sa"><x' | Should -BeNullOrEmpty
    }

    It 'finds languages by partial name' {
        (Find-Language 'arab').Id | Should -Contain 'ar-sa'
        (Find-Language 'french').Id | Should -Contain 'fr-ca'
    }

    It 'has unique ids' {
        $ids = (Get-LanguageTable).Id
        ($ids | Sort-Object -Unique).Count | Should -Be $ids.Count
    }

    It 'matches site/languages.json exactly' {
        $json = Get-Content -LiteralPath (Join-Path $Root 'site/languages.json') -Raw -Encoding UTF8 | ConvertFrom-Json
        $siteIds = @($json | ForEach-Object { $_.id.ToLowerInvariant() } | Where-Object { $_ -notin 'matchos' }) | Sort-Object
        $scriptIds = @((Get-LanguageTable).Id) | Sort-Object
        Compare-Object $scriptIds $siteIds | Should -BeNullOrEmpty
    }
}

Describe 'Products catalog' {
    It 'matches site/products.json' {
        $json = Get-Content -LiteralPath (Join-Path $Root 'site/products.json') -Raw -Encoding UTF8 | ConvertFrom-Json
        $catalog = Get-ProductCatalog
        foreach ($v in '2024', '2021', '2019') {
            $site = $json.versions | Where-Object { $_.version -eq $v }
            $site | Should -Not -BeNullOrEmpty
            $site.channel | Should -Be $catalog[$v].Channel
            $site.products.ProPlus | Should -Be $catalog[$v].ProPlus
            $site.products.Standard | Should -Be $catalog[$v].Standard
            $site.products.Visio | Should -Be $catalog[$v].Visio
            $site.products.Project | Should -Be $catalog[$v].Project
        }
    }
}

Describe 'Parameter parsing helpers' {
    It 'splits comma separated lists' {
        @(Split-ListValue 'Lync,Teams') | Should -Be @('Lync', 'Teams')
        @(Split-ListValue @('Lync', 'Teams, Word')) | Should -Be @('Lync', 'Teams', 'Word')
        @(Split-ListValue $null).Count | Should -Be 0
    }

    It 'normalises ExcludeApp ids' {
        @(Resolve-ExcludeApp @('lync', 'TEAMS', 'Lync')) | Should -Be @('Lync', 'Teams')
    }

    It 'rejects unknown ExcludeApp ids with exit code 6' {
        $err = $null
        try { Resolve-ExcludeApp @('Paint') } catch { $err = $_ }
        $err | Should -Not -BeNullOrEmpty
        $err.Exception.Data['OfficeLtscExitCode'] | Should -Be 6
    }
}

Describe 'Signature check' {
    It 'accepts a valid Microsoft signature' {
        $sig = [pscustomobject]@{ Status = 'Valid'; SignerCertificate = [pscustomobject]@{ Subject = 'CN=Microsoft Corporation, O=Microsoft Corporation, L=Redmond, S=Washington, C=US' } }
        Test-MicrosoftSignature $sig | Should -BeTrue
    }

    It 'rejects <name>' -TestCases @(
        @{ name = 'a HashMismatch'; status = 'HashMismatch'; subject = 'CN=Microsoft Corporation, O=Microsoft Corporation, C=US' }
        @{ name = 'an unsigned file'; status = 'NotSigned'; subject = $null }
        @{ name = 'another signer'; status = 'Valid'; subject = 'CN=Evil, O=Not Microsoft Corporation Ltd, C=US' }
        @{ name = 'a lookalike org'; status = 'Valid'; subject = 'CN=x, O=Microsoft Corporation Fake, C=US' }
    ) {
        $cert = if ($subject) { [pscustomobject]@{ Subject = $subject } } else { $null }
        $sig = [pscustomobject]@{ Status = $status; SignerCertificate = $cert }
        Test-MicrosoftSignature $sig | Should -BeFalse
    }
}

Describe 'configuration.xml generation' {
    It 'matches golden file 2024 ar-sa x64 with key and remove' {
        $xml = Get-OdtConfigurationXml -Version 2024 -Language 'ar-sa' -Arch 64 -Edition ProPlus -ExcludeApps @('Lync') `
            -ProductKey 'NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP' -RemoveExisting
        ([xml]$xml).OuterXml | Should -Be (Get-Golden '2024-ar-sa-x64-key-remove.xml').OuterXml
    }

    It 'matches golden file 2021 fr-fr x86 Standard + Visio + Project' {
        $xml = Get-OdtConfigurationXml -Version 2021 -Language 'fr-fr' -ExtraLanguages @('en-us') -Arch 32 -Edition Standard `
            -Visio -Project -ExcludeApps @('Teams', 'OneDrive') -NoUpdates -Silent
        ([xml]$xml).OuterXml | Should -Be (Get-Golden '2021-fr-fr-x86-standard-visio-project.xml').OuterXml
    }

    It 'matches golden file 2019 MatchOS with an escaped SourcePath' {
        $xml = Get-OdtConfigurationXml -Version 2019 -Language 'MatchOS' -SourcePath 'D:\Office & Data'
        ([xml]$xml).OuterXml | Should -Be (Get-Golden '2019-matchos-source.xml').OuterXml
    }

    It 'never adds PIDKEY or AUTOACTIVATE without a key' {
        $xml = Get-OdtConfigurationXml -Version 2024 -Language 'en-us'
        $xml | Should -Not -Match 'PIDKEY'
        $xml | Should -Not -Match 'AUTOACTIVATE'
    }

    It 'escapes hostile values instead of injecting markup' {
        $xml = Get-OdtConfigurationXml -Version 2024 -Language 'en-us' -SourcePath 'C:\x" /><Remove All="TRUE'
        $doc = [xml]$xml
        $doc.Configuration.SelectNodes('Remove').Count | Should -Be 0
        $doc.Configuration.Add.SourcePath | Should -Be 'C:\x" /><Remove All="TRUE'
    }
}

Describe 'Activation error text' {
    It 'explains common codes' {
        Get-ActivationErrorText '0xC004F074' | Should -Match 'KMS'
        Get-ActivationErrorText '0xc004c008' | Should -Match 'limit'
        Get-ActivationErrorText '0xC004C003' | Should -Match 'blocked'
    }
}
