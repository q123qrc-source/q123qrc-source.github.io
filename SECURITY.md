# Security

## What the script downloads

| What | From | How it is verified |
|---|---|---|
| Office Deployment Tool (`setup.exe`) | `https://officecdn.microsoft.com/pr/wsus/setup.exe` | `Get-AuthenticodeSignature` must return `Status = Valid` and a signer subject containing `O=Microsoft Corporation`. Otherwise the file is deleted and the script stops with exit code 3. |
| Office itself | Microsoft's CDN (`officecdn.microsoft.com`), downloaded by `setup.exe` | Done by Microsoft's own setup. Office files are Microsoft-signed `.cab` / `.dat` files. |
| OfficeScrubber (only with `-Scrub` / "deep clean") | `https://raw.githubusercontent.com/abbodi1406/BatUtil/<commit>/OfficeScrubber/...`, pinned to commit `4c02e02d86d2f01e4ef556bc6a080382bd21cae4` | Each of the 8 files is checked against a SHA-256 value hard-coded in `install.ps1`. Any mismatch deletes the folder and stops the script (exit code 3). |
| `install.txt` (only if the elevated relaunch cannot read the running script text) | this site | Same file as `/ps`. |

Nothing else is downloaded. The site does not host, mirror or modify any Microsoft file, and it contains no activation bypass of any kind. Product keys are the user's own (MAK / KMS client / retail).

OfficeScrubber is a third-party open source tool by abbodi1406. Its repository has no license file, so it is **not** redistributed here; it is fetched at run time from its original location. It runs Microsoft's OffScrub VBS scripts and, in "Scrub ALL" mode, also removes installed Office licenses and product keys.

## Product keys

- Never put in a URL or in the command the website generates.
- Entered with hidden input (`Read-Host -AsSecureString`) or passed by the user as `-ProductKey` in their own PowerShell window.
- When the script has to relaunch itself elevated, the key is handed over in a temporary file encrypted with Windows DPAPI (current user) and deleted as soon as it is read. It is never put on a command line.
- Logs and console output show only the last 5 characters.
- `configuration.xml` (which contains the key) is deleted with the work folder after a successful install. If the install fails and the folder is kept for troubleshooting, the key in it is masked first.

## Verifying the script

```powershell
irm https://q123qrc-source.github.io/install.txt -OutFile install.ps1
Get-FileHash .\install.ps1 -Algorithm SHA256
Get-AuthenticodeSignature .\install.ps1 | Format-List Status, SignerCertificate
```

Compare the hash with `https://q123qrc-source.github.io/install.ps1.sha256` and the value on the home page. Both are written by the deploy workflow from the exact published bytes. When the code-signing secrets are set, the published script carries an Authenticode signature from the owner's certificate.

## Reporting a problem

Open an issue (or a private security advisory) in this repository.
