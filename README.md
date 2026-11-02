# Office LTSC Installer

Installs **official Microsoft Office LTSC 2024 / 2021 / Office 2019 (volume)** in any language Office supports, using only Microsoft's Office Deployment Tool and Microsoft's CDN.

> Independent tool. Not affiliated with Microsoft. Office is downloaded directly from Microsoft's servers.

- Website (Arabic / English): **https://q123qrc-source.github.io/**
- Interactive install, from PowerShell (Admin):

```powershell
irm https://q123qrc-source.github.io/ps | iex
```

- Non-interactive:

```powershell
& ([scriptblock]::Create((irm https://q123qrc-source.github.io/ps))) -Version 2024 -Language ar-sa -Arch 64
```

## Parameters

| Parameter | Values | Notes |
|---|---|---|
| `-Version` | `2024`, `2021`, `2019` | Required in non-interactive mode |
| `-Language` | culture code (`ar-sa`, `en-us`, `fr-fr`, ...), `MatchOS`, `MatchInstalled` | Default `MatchOS` |
| `-ExtraLanguages` | culture codes | `-ExtraLanguages en-us,fr-fr` |
| `-Arch` | `64`, `32` | Default `64` |
| `-Edition` | `ProPlus`, `Standard` | Default `ProPlus` |
| `-Visio`, `-Project` | switch | Adds VisioPro / ProjectPro volume products |
| `-ExcludeApps` | `Access, Excel, Groove, Lync, OneDrive, OneNote, Outlook, OutlookForWindows, PowerPoint, Publisher, Teams, Word` | |
| `-ProductKey` | `XXXXX-XXXXX-XXXXX-XXXXX-XXXXX` | Your own key. Never logged in full. With a key, ODT activates automatically when online |
| `-RemoveExisting` | switch | Adds `<Remove All="TRUE"/>` and `<RemoveMSI/>` |
| `-Scrub` | switch | Runs OfficeScrubber before installing (implies `-RemoveExisting`), see SECURITY.md |
| `-SourcePath` | folder | Install from an offline source (`Office\Data` inside), with CDN fallback |
| `-DownloadOnly` | switch | `setup.exe /download` into `-SourcePath`; also writes `setup.exe` + `configuration.xml` there |
| `-NoUpdates` | switch | `<Updates Enabled="FALSE"/>` |
| `-Silent` | switch | `<Display Level="None"/>` |

Exit codes: `0` installed, `1` cancelled, `2` pre-check failed, `3` download / signature / hash check failed, `4` ODT failed, `5` installed but activation not confirmed, `6` invalid input.

When run through `irm ... | iex` the script does not close your window; the exit code is left in `$LASTEXITCODE`.

## Repository layout

```
src/install.ps1          the only source of the script (pure ASCII, CI-enforced)
site/                    landing page, languages.json, products.json, assets
tests/                   Pester tests + golden configuration.xml files
.github/workflows/ci.yml     ASCII check, parse test, PSScriptAnalyzer, Pester (PS 5.1 + 7)
.github/workflows/pages.yml  build (+ optional signing) -> deploy -> smoke test of irm /ps
```

Published URLs: `/` (page), `/ps` (script, used by `irm`), `/install.txt` (same bytes), `/install.ps1`, `/install.ps1.sha256`, `/languages.json`, `/products.json`.

## GitHub Pages setup (one time)

1. The repository must be **public** and named exactly `q123qrc-source.github.io`.
2. Merge the work into the **`main`** branch (the deploy workflow runs on pushes to `main`) and make `main` the default branch (Settings → General → Default branch).
3. Settings → Pages → Build and deployment → Source: **GitHub Actions**.
4. Settings → Pages → tick **Enforce HTTPS**. Do not add a custom domain / `CNAME`.
5. Push to `main` (or run the "Deploy to GitHub Pages" workflow manually). The `smoke` job then fetches `/ps` with Windows PowerShell 5.1 and PowerShell 7 and checks it is identical to the published hash and parses.

GitHub Pages caches for up to 10 minutes (`Cache-Control: max-age=600`), so a new release may take that long to reach users.

### Code signing (optional, recommended)

The deploy workflow signs `install.ps1` with Authenticode when these repository secrets exist (Settings → Secrets and variables → Actions):

| Secret | Value |
|---|---|
| `CODESIGN_PFX_BASE64` | your code-signing certificate (`.pfx`) as base64: `[Convert]::ToBase64String([IO.File]::ReadAllBytes('cert.pfx')) \| Set-Clipboard` |
| `CODESIGN_PFX_PASSWORD` | the `.pfx` password |

Optional variable `CODESIGN_TIMESTAMP_URL` (default `http://timestamp.digicert.com`). Without the secrets the script is published unsigned and the workflow shows a warning. A self-signed certificate works technically, but only machines that trust it will show the signature as `Valid`; a certificate from a public CA (or Azure Trusted Signing) is needed for everyone else.

## Why `/ps` and pure ASCII

GitHub Pages cannot vary on User-Agent, so `/` is always HTML and the script is served from `/ps/index.html`. That file is sent as `text/html` with no charset, which Windows PowerShell 5.1 decodes as ISO-8859-1, so the script must be pure ASCII (no BOM, no Arabic, no smart quotes). Arabic text lives only in the web page and `languages.json`.

## Testing status

| Check | Status |
|---|---|
| PSScriptAnalyzer, ASCII, parse, Pester (PS 5.1 and 7) | Runs in CI on every push (`ci.yml`) |
| `irm /ps` and `/install.txt` equal the published script in PS 5.1 and 7 | Runs after every deploy (`pages.yml` → `smoke`) |
| What `irm /install.ps1` returns | Printed by the smoke job; `/ps` and `/install.txt` are the supported URLs |
| Real installs in Windows Sandbox (2024 en-us / ar-sa x64, 2021 fr-fr, 32-bit, MatchOS), tampered `setup.exe`, no-admin relaunch, `-RemoveExisting` over Office 2016 MSI + Microsoft 365 | **Not run yet** - must be done manually on Windows (see checklist below) |
| Web page: Arabic / English, light / dark, 360 px, Copy button | Checked in a headless browser |

Manual checklist (Windows Sandbox):

- [ ] 2024 `en-us` x64, 2024 `ar-sa` x64, 2021 `fr-fr` x64, one `-Arch 32`, one `MatchOS`
- [ ] Flip a byte in the downloaded `setup.exe` (e.g. break in the debugger, or test `Test-MicrosoftSignature` with a modified copy) → exit code 3
- [ ] Invalid key / unknown language → exit code 6 with a clear message
- [ ] Non-admin window → UAC prompt → continues with the same parameters
- [ ] `-RemoveExisting` and `-Scrub` with Office 2016 MSI + Microsoft 365 installed
- [ ] Product key never appears in full in the log or console

## License

MIT for the files in this repository (see `LICENSE`). Microsoft Office, the Office Deployment Tool and OfficeScrubber keep their own licenses and are never redistributed here.
