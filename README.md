# FreeMixKit

![Platform](https://img.shields.io/badge/platform-Windows-blue)
![License: MIT](https://img.shields.io/badge/license-MIT-blue)
[![GitHub stars](https://img.shields.io/github/stars/catsmoker/FreeMixKit?style=social)](https://github.com/catsmoker/FreeMixKit/stargazers)

FreeMixKit is a Windows PowerShell toolbox that bundles maintenance tasks, developer setup, common utilities, and activation helpers into one menu-driven script.

**Goal:** keep trusted free software (and vetted patches) in one place for easy setup and upkeep. [* also trusted cracked software 😈]

<img width="1214" height="668" alt="Screenshot 2026-03-10 171357" src="https://github.com/user-attachments/assets/9bbc7abc-c302-4e53-a662-79033f4afcca" />

(*): There is no such thing as “trusted cracked software” because cracking inherently involves modifying and bypassing a program’s licensing and security mechanisms, which makes the source unverifiable and legally unauthorized; as a result, these versions frequently contain hidden malware, backdoors, or spyware, cannot receive secure official updates, and expose users to data theft, system compromise, and legal consequences, whereas legitimate options such as open-source alternatives, student licenses, official discounts, or free trials provide verifiable integrity, ongoing security patches, and lawful use without the systemic risks associated with altered binaries.
What I mean by “trusted cracked software” is software that is widely used by many people, frequently discussed on forums, and considered popular within certain online communities.

## Contents

- [Highlights](#highlights)
- [Module Overview](#module-overview)
- [Requirements](#requirements)
- [Quick Start](#quick-start)
- [Run Locally](#run-locally)
- [Safety & Transparency](#safety--transparency)
- [Recommended Practice](#recommended-practice)
- [License](#license)

## Highlights

- One PowerShell entry point: `w.ps1` launches all modules.
- Covers dev setup, cleanup/repair, media tweaks, and activation helpers.
- Uses Winget and bundled scripts to automate most installs.
- Built for Windows 10/11; self-elevates when admin rights are required.

## Module Overview

- **Developer** – Installs runtimes, package managers, IDEs, and CLI tools via Winget.
- **Maintenance** – Disk cleanup, SFC/DISM repair, MRT + KVRT scans, and system reports.
- **Software** – Shortcuts to tools such as Adobe GenP, Winget Upgrade, Spicetify, Legcord.
- **Utilities** – Resolution helper, shortcut creator, WinUtil launcher, other small helpers.
- **Activation / third-party patching** – Runs external activation scripts and patches; review before use.

## Requirements

- Windows 10 or 11
- PowerShell 5.1 or PowerShell 7
- Administrator privileges (the script self-elevates)
- Internet access for downloads and remote scripts

## Quick Start

Run directly from PowerShell (recommended):

```powershell
irm https://raw.githubusercontent.com/catsmoker/FreeMixKit/main/w.ps1 | iex
```

## Run Locally

From the repository root:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\w.ps1
```

Using PowerShell 7:

```powershell
pwsh -NoProfile -ExecutionPolicy Bypass -File .\w.ps1
```

## Safety & Transparency

- Some modules download and execute third-party tools or activation scripts; read `w.ps1` before running them.
- Activation modules may violate license terms or local law—use only where legal and at your own risk.
- Several actions modify system settings, services, and the registry. Ensure you understand each option you select.
- Source code is plain PowerShell; inspect or edit to suit your environment.

## Recommended Practice

- Try high-risk modules on a non-production machine first.
- Keep backups before registry or security changes.
- Run one change at a time and verify system behavior between steps.

## License

This project is licensed under the MIT License. See [LICENSE](LICENSE) for details.


