# FreeMixKit

![Platform](https://img.shields.io/badge/platform-Windows-blue)
![License: MIT](https://img.shields.io/badge/license-MIT-blue)
[![GitHub stars](https://img.shields.io/github/stars/catsmoker/FreeMixKit?style=social)](https://github.com/catsmoker/FreeMixKit/stargazers)

FreeMixKit is a Windows PowerShell utility suite with a keyboard-driven grid menu for maintenance, networking, developer setup, and system tweaks.

## What It Does

`w.ps1` provides grouped modules such as:

- `Developer`: full dev environment setup (runtime and tool installs via Winget)
- `Maintenance`: cleanup, system repair (SFC/DISM), malware scan, system report
- `Tweaks`: DNS profiles, network reset, telemetry and service toggles, registry optimization
- `Utilities`: registry backup, shortcut creation, resolution helper, external utility launchers
- `Activation / third-party patching`: includes modules that execute external activation scripts

## Important Safety Notes

- Run only if you understand each module you execute.
- Several modules make system-level changes (services, Defender state, registry, DNS, Winsock).
- Some modules download and run remote scripts/tools. Review `w.ps1` before use.
- Activation-related modules may violate software terms or laws depending on your jurisdiction.

## Requirements

- Windows 10/11
- PowerShell (Windows PowerShell 5.1 or PowerShell 7)
- Administrator privileges (the script self-elevates)
- Internet connection for modules that download packages/scripts

## Run

(Recommended) type:
 ```powershell
 irm catsmoker.github.io/w | iex
 ```
 in the powershell terminal.

From this folder:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\w.ps1
```

PowerShell 7 alternative:

```powershell
pwsh -NoProfile -ExecutionPolicy Bypass -File .\w.ps1
```

## Controls

Inside the UI:

- `Up/Down`: move within a column
- `Left/Right`: switch columns
- `Enter`: run selected module
- `Q` or `Esc`: exit

## Output and Logs

On exit, session results are exported to:

- `%USERPROFILE%\Desktop\FreeMixKit_ModuleResults_<timestamp>.csv`

This includes module name, status, duration, message, and rollback hints (if available).

## Project Structure

- `w.ps1`: main application script
- `README.md`: project documentation

## Recommended Practice

- Test high-risk modules on a non-production machine first.
- Keep backups before registry/service/security changes.
- Run one change at a time and verify system behavior before continuing.

### Disclaimer

This script includes tools that can modify your system extensively,
It also contains modules that automate the download and execution of third-party software and activation scripts,
Please use these features responsibly and at your own risk,
The author is not responsible for any data loss or system instability.

##  License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.
