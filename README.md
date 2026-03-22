# BingSpotlight

A PowerShell project that brings a Bing Spotlight-like lock screen experience to Windows 11-10 LTSC. Should also work with non-LTSC Windows and probably not with Home versions.

The project downloads the current Bing image, renders a text banner on top of it, applies the generated image to the Windows lock screen through the registry, keeps a local image history, and falls back to the latest valid image if Bing is unavailable.

## Why this project

Windows Spotlight is not natively available on some LTSC editions. This project provides a simple, local, and installable alternative based on Windows PowerShell 5.1.

## Features

- interactive installation
- Bing market selection during setup
- configurable retention for generated images
- automatic fallback to the latest valid image
- local logging
- scheduled task running as `SYSTEM`
- clean uninstall with explicit confirmation

## Target environment

- Windows 11 LTSC
- Windows PowerShell 5.1
- Administrator rights

## Repository contents

- [Install-BingSpotlight.ps1](Install-BingSpotlight.ps1): installs the solution and creates the scheduled task
- [BingSpotlight.ps1](BingSpotlight.ps1): main script executed by the scheduled task
- [Uninstall-BingSpotlight.ps1](Uninstall-BingSpotlight.ps1): clean uninstaller
- [plan.md](plan.md): design notes and technical details

## Installation

Open PowerShell as Administrator in the project directory.

If script execution is blocked on the machine, temporarily allow it for the current session:

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
```

Then run the installer:

```powershell
.\Install-BingSpotlight.ps1
```

The installer asks for:

1. how many days generated images should be kept
2. which Bing market to use, for example `fr-FR` or `en-US`

After installation, the files are copied to:

```text
C:\ProgramData\BingSpotlight
```

## Installed structure

```text
C:\ProgramData\BingSpotlight\
├── BingSpotlight.ps1
├── Uninstall-BingSpotlight.ps1
├── config.json
├── logs\
│   └── BingSpotlight.log
├── source\
│   └── bing_source.jpg
└── rendered\
    ├── lockscreen_2026-03-22.jpg
    └── ...
```

## Configuration

The configuration file is created automatically here:

```text
C:\ProgramData\BingSpotlight\config.json
```

Example:

```json
{
  "Market": "fr-FR",
  "RetentionDays": 14,
  "RetryCount": 5,
  "RetryDelaySeconds": 15
}
```

## How it works

At each run, the main script:

1. reads `config.json`
2. calls the Bing API for the configured market
3. downloads the current image
4. renders the title and copyright on the image
5. saves a dated image in `rendered`
6. applies that image to the lock screen through `HKLM`
7. removes older images according to the configured retention period

Using a dated filename helps reduce Windows lock screen caching issues.

## Network failure behavior

The project does not blindly overwrite the current state when Bing is unavailable.

- network calls use configurable retry logic
- if Bing remains unavailable, the script tries to reapply the latest valid image stored in `rendered`
- if no valid image exists yet, the run fails cleanly and writes the error to the log

## Post-installation test

To trigger an immediate test run:

```powershell
Start-ScheduledTask -TaskPath "\Custom\" -TaskName "BingSpotlight_LockScreen"
```

Then verify:

1. that a `lockscreen_*.jpg` file appears in `C:\ProgramData\BingSpotlight\rendered`
2. that the log contains a complete successful run
3. that the lock screen shows the generated image

Quick log check:

```powershell
Get-Content "C:\ProgramData\BingSpotlight\logs\BingSpotlight.log" -Tail 20
```

Expected lines usually include:

```text
[INFO] Execution started. RetentionDays=...
[INFO] Metadata retrieved: ...
[INFO] Source image downloaded.
[INFO] Rendered image created: ...
[INFO] Registry updated successfully.
[INFO] Cleanup finished.
```

## Uninstall

To uninstall:

```powershell
C:\ProgramData\BingSpotlight\Uninstall-BingSpotlight.ps1
```

The script asks the user to type `OUI` to confirm, then it:

1. removes the `\Custom\BingSpotlight_LockScreen` scheduled task
2. clears `PersonalizationCSP` values if they still point to the installation directory
3. removes `C:\ProgramData\BingSpotlight`

## Limitations and notes

- the project targets Windows PowerShell 5.1, not PowerShell 7 as the primary runtime
- `System.Drawing` is used for image rendering
- the Bing endpoint used here is an unofficial `HPImageArchive` API
- this project is not affiliated with Microsoft

## License

This repository is licensed under the GNU Affero General Public License v3.0.

See [LICENSE](LICENSE).
