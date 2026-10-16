# FreeMixKit

![Platform](https://img.shields.io/badge/platform-Windows-blue)
![License: MIT](https://img.shields.io/badge/license-MIT-blue)
[![GitHub stars](https://img.shields.io/github/stars/catsmoker/FreeMixKit?style=social)](https://github.com/catsmoker/FreeMixKit/stargazers)

FreeMixKit is a Windows PowerShell toolbox that bundles maintenance tasks, developer setup, common utilities, and activation helpers into one desktop app.

**Goal:** keep trusted free software (and vetted patches) in one place for easy setup and upkeep. [* also trusted cracked software 😈]

<img width="1120" height="760" alt="FreeMixKit main window: sidebar navigation, module cards with risk badges and Run buttons, and the output console" src="screenshot.png" />

The Software tab lays every installable out as a logo card grid that reflows as the window is resized:

<img width="1120" height="760" alt="The Software tab: a responsive grid of logo cards, each with an Install button, with the search box filtering the page" src="screenshot-software.png" />

(*): There is no such thing as “trusted cracked software” because cracking inherently involves modifying and bypassing a program’s licensing and security mechanisms, which makes the source unverifiable and legally unauthorized; as a result, these versions frequently contain hidden malware, backdoors, or spyware, cannot receive secure official updates, and expose users to data theft, system compromise, and legal consequences, whereas legitimate options such as open-source alternatives, student licenses, official discounts, or free trials provide verifiable integrity, ongoing security patches, and lawful use without the systemic risks associated with altered binaries.
What I mean by “trusted cracked software” is software that is widely used by many people, frequently discussed on forums, and considered popular within certain online communities.

## Contents

- [Highlights](#highlights)
- [Module Overview](#module-overview)
- [Requirements](#requirements)
- [Quick Start](#quick-start)
- [Run Locally](#run-locally)
- [Building from source](#building-from-source)
- [Safety & Transparency](#safety--transparency)
- [Recommended Practice](#recommended-practice)
- [License](#license)

## Highlights

- One PowerShell entry point: `w.ps1` launches the whole app.
- WPF desktop window: sidebar navigation by area, compact cards with risk badges plus tooltips, group chips plus live search with a match count, a dark/light theme toggle, and a streaming log pane.
- Software tab: click a card to select it, click again to deselect; one **Install Selected (n)** button installs the queue sequentially in selection order, with step/total progress, per-item results, cancel for pending items, and a final summary. Search and chip filters never deselect.
- Exclusive tasks show a loading overlay with status over a dimmed, locked interface while real command output streams into the log pane: plain tool output in white, errors and warnings colored, framework status lines muted gray. The shared progress bar shows a measured percentage with step/total (e.g. `Installing X (2 of 5) — 40%`) whenever the output carries one, stays honestly indeterminate otherwise, and keeps a green 100% / red / amber end state until the next task. Background runspaces keep the window responsive through long operations such as SFC and DISM.
- Risk ratings explained in the UI: Low, Medium, and High badges each render in their own color — hover any badge for the reason and dimensions, or open **Risk ratings?** for the full criteria. Ordinary tools stay untagged.
- The Software tab is a responsive logo-card grid, and every developer tool installs on its own — one card per component, never a single bundle that drags the rest in.
- Covers dev setup, cleanup/repair, media tweaks, and activation helpers.
- Uses Winget and bundled scripts to automate most installs.
- Built for Windows 10/11; self-elevates when admin rights are required.

## Module Overview

Each tab is a sidebar page; the entries under it are grouped into sections.

- **System → Cleanup & Repair** – Disk cleanup, SFC/DISM repair, MRT + KVRT scans, and a System Report action that writes a full hardware/Windows diagnostic text file to the desktop.
- **Utilities → Fixes / Windows Tweaks** – Resolution and Game Bar fixes; Winget Upgrade, WinUtil, and Winhance. Tools for managing, optimising, or maintaining the machine.
- **Software → Browsers, Messaging, Media, File Sharing & Downloads, Development, Video & Image Tools, Gaming, System & Maintenance, Productivity** – Standalone applications as compact logo cards; **click cards to select them, then Install Selected** runs the queue in the order selected (no per-card Install button). A chip bar filters by group, combines with live search, and shows the match count. The queue survives filtering, reflows from the top-left, and resets the scroll to the top on each query change:
  - *Development* – Visual C++ Redistributables, .NET, Node.js, Python, Java, PowerShell 7, Git, FFmpeg, Bibata Cursor, Visual Studio Community, Android Studio, Godot Engine, and Unity Hub. **Each card installs only its own component.**
  - *Productivity* – NanaZip and Notepad++.
  - *Video & Image Tools* – File Converter and HandBrake.
  - *Gaming* – DirectX/XNA game runtimes and Steam.
  - *System & Maintenance* – Mem Reduct, WinMemoryCleaner, and Bulk Crap Uninstaller.
  - *Media* – Spicetify, SpotX, VLC, and Stremio.
  - *Messaging* – Legcord, Telegram Desktop, and ZapFast (unofficial WhatsApp client, installed from its official GitHub releases).
  - *Browsers* – Brave. *File Sharing & Downloads* – LocalSend and qBittorrent.
  - New catalog entries install via WinGet whenever a verified package exists, falling back to the vendor's official download page; every package ID was verified with `winget search --id <Id> --exact`.
- **Activation → Microsoft / Other / Adobe** – Activation and licensing tools (MAS, IDM, Adobe CC & GenP) plus Adobe's official uninstaller. Review before use.
- **About** – What Free Mix Kit is and what the current version can do, plus support actions: repository, website, and desktop shortcut links, a **Report an Issue** button opening the official GitHub Issues page, and **Share Logs** to save sanitized diagnostics for attaching to a bug report.

## Requirements

- Windows 10 or 11
- PowerShell 5.1 or PowerShell 7
- Administrator privileges (the script self-elevates)
- Internet access for downloads and remote scripts

Windows PowerShell 5.1 already runs on a single-threaded apartment, which WPF
requires. PowerShell 7 starts multi-threaded, so the script relaunches itself
once with `-STA` to satisfy both that and elevation in a single restart. If the
host cannot be relaunched, it falls back to Windows PowerShell.

## Quick Start

Run directly from PowerShell (recommended):

```powershell
irm https://raw.githubusercontent.com/catsmoker/FreeMixKit/main/w.ps1 | iex
```

The window opens once the script has elevated itself.

## Run Locally

From the repository root:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\w.ps1
```

Using PowerShell 7:

```powershell
pwsh -NoProfile -ExecutionPolicy Bypass -File .\w.ps1
```

## Building from source

`w.ps1` at the repository root is a **generated** file — it is committed so the
`irm | iex` one-liner above keeps working, but it is not where you edit. The
source of truth lives in `src/`:

```
src/
├── w.template.ps1     entry point: setup, elevation/STA relaunch, boot
├── xaml/ui.xaml       the WPF window
├── theme/theme.ps1    colour tokens and Set-FmkUITheme
├── assets/logos.ps1   vendored logo path data for the Software grid
├── functions/         UI plumbing (runspace, logging, dialogs, cards, boot)
├── modules/           the module bodies, grouped by area
└── registry.ps1       module registration: label, risk, description, UI placement
```

Rebuild after any change under `src/`:

```powershell
.\Compile.ps1            # writes .\w.ps1
.\Compile.ps1 -Verify    # validates the sources, writes nothing
powershell -NoProfile -ExecutionPolicy Bypass -File .\tests\Test-Fmk.ps1
```

`Compile.ps1` parses every source file with the real PowerShell parser and
cross-checks the wiring, so a module cannot silently vanish from the interface or
fail the first time it is clicked: every `Category` in `registry.ps1` must have a
matching page in `ui.xaml`, every `Register-Module` must point at a module that
exists, every `%%Token%%` must have a theme key, every `Logo` must resolve to an
asset entry, and those asset entries must be shaped correctly.

## Risk ratings

Risk ratings explain the potential consequences of using a tool — technical safety, system changes, privacy, source trustworthiness, and licensing or service-policy concerns. They are guidance, not a guarantee of safety.

- **LOW** – Limited or well-understood risk. Reputable source, predictable changes, no significant security concern. Policy caveats may still apply.
- **MEDIUM** – Meaningful caveats: significant modifications, admin operations, stability or privacy implications, or uncertain compatibility/licensing.
- **HIGH** – Significant risk or serious consequences: unverified executables, disabled security controls, destructive or deceptive behavior. Licensing alone never makes High.
- **Untagged** – No warning warranted (ordinary verified installs).

Key reassessments (rationales in `src/registry.ps1`): **SpotX Medium** (patches the Spotify client; ad-blocking may conflict with Spotify terms), **Legcord Medium** (unofficial client handling account tokens; Discord ToS caveats), **Winhance Low** (established utility; review options before applying), **WinUtil Low** overall (established open utility; inner operations vary and need admin), **GenP High** (unverified mirror executables plus Adobe-terms impact, documented separately), MAS/IAS **High** with explicit confirmations. Ordinary Winget catalog entries such as .NET, Notepad++, and NanaZip are untagged.

## Safety & Transparency

- Some modules download and execute third-party tools or activation scripts; read `w.ps1` before running them.
- Activation modules may violate license terms or local law—use only where legal and at your own risk. High-risk modules require an explicit confirmation; the Software queue shows one combined Medium/High review before installing.
- The queue never bypasses UAC, signature checks, or Windows security controls, never retries forever, and never swaps in an unrelated installer after a failure.
- Several actions modify system settings, services, and the registry. Ensure you understand each option you select.
- Source code is plain PowerShell; inspect or edit to suit your environment.
- Verify before you run: `irm https://raw.githubusercontent.com/catsmoker/FreeMixKit/main/w.ps1` downloads the file — save and review it first when practical, and check the release channel rather than executing an unverified replacement.

## Recommended Practice

- Try high-risk modules on a non-production machine first.
- Keep backups before registry or security changes.
- Run one change at a time and verify system behavior between steps.

## License

This project is licensed under the MIT License. See [LICENSE](LICENSE) for details.


