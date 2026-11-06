# FreeMixKit

![Platform](https://img.shields.io/badge/platform-Windows-blue)
![License: MIT](https://img.shields.io/badge/license-MIT-blue)
![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-blue)
[![GitHub stars](https://img.shields.io/github/stars/catsmoker/FreeMixKit?style=social)](https://github.com/catsmoker/FreeMixKit/stargazers)
[![GitHub issues](https://img.shields.io/github/issues/catsmoker/FreeMixKit)](https://github.com/catsmoker/FreeMixKit/issues)

FreeMixKit is a **Windows PowerShell toolbox** with a WPF desktop interface that bundles system maintenance, developer setup, common utilities, and activation helpers into a single, self-contained application.

**Goal:** Keep trusted free software (and vetted patches) in one place for easy setup and upkeep.

---

## Table of Contents

- [Screenshots](#screenshots)
- [Highlights](#highlights)
- [Module Overview](#module-overview)
- [Requirements](#requirements)
- [Quick Start](#quick-start)
- [Run Locally](#run-locally)
- [Building from Source](#building-from-source)
- [Project Structure](#project-structure)
- [Architecture Overview](#architecture-overview)
- [Adding New Modules](#adding-new-modules)
- [Risk Ratings](#risk-ratings)
- [Safety & Transparency](#safety--transparency)
- [Recommended Practice](#recommended-practice)
- [Troubleshooting](#troubleshooting)
- [Contributing](#contributing)
- [License](#license)

---

## Screenshots

| Main Window | Software Tab |
|-------------|--------------|
| ![FreeMixKit main window](screenshot.png) | ![Software tab](screenshot-software.png) |

*Screenshots show the sidebar navigation, module cards with risk badges, streaming log pane, and the responsive logo-card grid on the Software tab.*

---

## Highlights

- **Single entry point:** `w.ps1` launches the entire application
- **WPF desktop UI:** Sidebar navigation, compact cards with risk badges & tooltips, group chips with live search & match count, dark/light theme toggle, streaming log pane
- **Software queue:** Click cards to select/deselect; one **Install Selected (n)** button runs the queue in selection order with step/total progress, per-item results, cancel for pending items, and a final summary. Filters never deselect.
- **Exclusive task overlay:** Loading overlay with status over a dimmed, locked interface while real command output streams to the log (white=stdout, colored=warnings/errors, gray=framework status). Shared progress bar shows measured percentage with step/total (e.g. `Installing X (2 of 5) — 40%`) when available, stays honestly indeterminate otherwise, and retains green 100% / red / amber end state until the next task.
- **Background runspaces:** Keeps the window responsive during long operations (SFC, DISM, large installs)
- **Risk ratings in UI:** Low/Medium/High badges with distinct colors — hover for reason & dimensions, or open **Risk ratings?** for full criteria. Ordinary verified installs stay untagged.
- **Responsive logo-card grid:** Every developer tool installs independently — one card per component, never a monolithic bundle.
- **WinGet-first installs:** Verified package IDs (`winget search --id <Id> --exact`), fallback to vendor download pages
- **Self-elevation:** Automatically requests admin rights when needed
- **Windows 10/11 support:** PowerShell 5.1 (native STA) and PowerShell 7 (auto-relaunches with `-STA`)

---

## Module Overview

Each tab is a sidebar page; entries are grouped into sections.

| Category | Groups | Description |
|----------|--------|-------------|
| **System** | Cleanup & Repair | Disk cleanup, SFC/DISM repair, MRT + KVRT scans, System Report (hardware/Windows diagnostics → Desktop) |
| **Utilities** | Fixes / Windows Tweaks | Resolution fix (CRU), ms-gamebar-fix, Winget Upgrade, WinUtil, Winhance |
| **Software** | 9 groups (see below) | Standalone apps as logo cards — click to select, then **Install Selected** runs queue in order |
| **Activation** | Microsoft / Other / Adobe | MAS, IAS, Adobe CC, GenP, Adobe Uninstaller |
| **About** | Project / Desktop | GitHub repo, website, desktop shortcut, Report Issue, Share Logs |

### Software Tab Groups

| Group | Applications |
|-------|--------------|
| **Development** | Visual C++ Redistributables, .NET (SDK 10 + runtimes 8/9), Node.js LTS, Python 3, Eclipse Temurin JDK, PowerShell 7, Git, FFmpeg, Bibata Cursor, Visual Studio Community, Android Studio, Godot Engine, Unity Hub |
| **Productivity** | NanaZip, Notepad++ |
| **Video & Image Tools** | File Converter, HandBrake |
| **Gaming** | DirectX/XNA Game Runtimes, Steam |
| **System & Maintenance** | Mem Reduct, WinMemoryCleaner, Bulk Crap Uninstaller |
| **Media** | Spicetify, SpotX, VLC, Stremio |
| **Messaging** | Vesktop (Vencord), Telegram Desktop, ZapFast (unofficial WhatsApp, from official GitHub releases) |
| **Browsers** | Brave, LibreWolf |
| **File Sharing & Downloads** | LocalSend, qBittorrent |

> **Note:** Every catalog entry installs via WinGet when a verified package exists, otherwise falls back to the vendor's official download page. All package IDs were verified with `winget search --id <Id> --exact`.

---

## Requirements

| Requirement | Details |
|-------------|---------|
| **OS** | Windows 10 or 11 |
| **PowerShell** | 5.1 (built-in) or 7.x |
| **Privileges** | Administrator (script self-elevates via UAC) |
| **Network** | Internet access for downloads & remote scripts |

**STA Note:** WPF requires a single-threaded apartment. Windows PowerShell 5.1 is STA by default. PowerShell 7 runs MTA, so the script relaunches itself once with `-STA` (combined with elevation in a single restart). If the host cannot be relaunched, it falls back to Windows PowerShell.

---

## Quick Start

Run directly from PowerShell (downloads & executes latest `w.ps1`):

```powershell
irm https://raw.githubusercontent.com/catsmoker/FreeMixKit/main/w.ps1 | iex
```

> ⚠️ **Security:** The `irm | iex` pattern executes remote code. For production use, [review the source first](#verify-before-running) or [run locally](#run-locally).

---

## Run Locally

From the repository root:

```powershell
# Windows PowerShell 5.1
powershell -NoProfile -ExecutionPolicy Bypass -File .\w.ps1

# PowerShell 7 (pwsh)
pwsh -NoProfile -ExecutionPolicy Bypass -File .\w.ps1
```

---

## Building from Source

`w.ps1` at the repository root is a **generated file** — committed so the `irm | iex` one-liner works, but **not where you edit**. The source of truth lives in `src/`:

```
src/
├── w.template.ps1      # Entry point: setup, elevation/STA relaunch, boot
├── xaml/ui.xaml        # WPF window definition (with %%Token%% placeholders)
├── theme/theme.ps1     # Colour tokens (Dark/Light) + Get-FmkThemedXaml
├── assets/logos.ps1    # Vendored SVG path data for Software grid icons
├── functions/          # UI plumbing (01-helpers → 11-sysreport, loaded alphabetically)
├── modules/            # Module scriptbodies ($Modules["Key"] = { ... })
└── registry.ps1        # Module registration: label, risk, Category, Group, Logo, etc.
```

### Rebuild Commands

```powershell
# Full rebuild (writes w.ps1)
.\Compile.ps1

# Validate only (no output written) — run before committing
.\Compile.ps1 -Verify

# Run tests (if available)
powershell -NoProfile -ExecutionPolicy Bypass -File .\tests\Test-Fmk.ps1
```

### What `Compile.ps1` Validates

- Every source file parses cleanly with the real PowerShell parser
- No assignment to read-only automatic variables (`$host`, `$HOME`, `$PSScriptRoot`, etc. — case-insensitive)
- Every `Category` in `registry.ps1` has a matching `x:Name="page_<Category>"` in `ui.xaml`
- Every `%%Token%%` in XAML has a matching key in `theme.ps1`
- XAML is well-formed XML (no `--` in comments, no bare `&`)
- Every logo entry is a valid brand mark (`Color`+`Data`), monogram (`Mono`), PNG (`Png`), or shape set (`Shapes`)
- Every `Logo = "X"` in registry resolves to an asset entry
- Every `Register-Module` points to an existing `$Modules["Key"]` definition
- No duplicate registration keys
- Risk metadata: Medium/High require `RiskReason`; High requires `ConfirmMessage`

Output is CRLF, UTF-8 without BOM. There are no external tests/linters/CI — **`-Verify` is the check**.

---

## Project Structure

```
FreeMixKit/
├── w.ps1                 # GENERATED — built by Compile.ps1 (do not edit)
├── Compile.ps1           # Build script: concatenates src/ → w.ps1 + validation
├── README.md             # This file
├── LICENSE               # MIT License
├── freemixkit_icon.ico   # Application icon
├── docs/
│   ├── BUILD.md          # Detailed build documentation
│   ├── DEVELOPMENT.md    # Development guidelines
│   └── SECURITY.md       # Security policy
├── src/
│   ├── w.template.ps1
│   ├── xaml/ui.xaml
│   ├── theme/theme.ps1
│   ├── assets/
│   │   ├── logos.ps1
│   │   └── icon.ps1      # Base64-encoded .ico for embedded use
│   ├── functions/
│   │   ├── 01-helpers.ps1
│   │   ├── 02-registry.ps1
│   │   ├── 03-ui-log.ps1
│   │   ├── 04-runspace.ps1
│   │   ├── 05-dialogs.ps1
│   │   ├── 06-cards.ps1
│   │   ├── 07-boot.ps1
│   │   ├── 08-risk.ps1
│   │   ├── 09-queue.ps1
│   │   ├── 10-support.ps1
│   │   └── 11-sysreport.ps1
│   ├── modules/
│   │   ├── software.ps1
│   │   ├── developer.ps1
│   │   ├── activation.ps1
│   │   ├── maintenance.ps1
│   │   └── utils.ps1
│   └── registry.ps1
└── tests/
    └── Test-Fmk.ps1
```

---

## Architecture Overview

### Load Order (matters)

`src/w.template.ps1` → **setup → elevation/STA relaunch → `Set-StrictMode 1.0` → helpers → theme → modules → registry → XAML → boot**

Runtime uses **StrictMode 1.0** (not `Latest`). Version 2.0+ throws on property access against missing keys/`$null`; the UI code legitimately touches missing keys (FindName lookups, optional ModuleMeta fields, keys absent from `$sync` before boot finishes). 1.0 keeps useful uninitialized-variable checking and drops the rest.

### Module System

- **`functions/01..11`** load alphabetically: helpers → registry fns → ui-log → runspace → dialogs → cards → boot → risk → queue → support → sysreport
- **`modules/*.ps1`** only assign `$Modules["Key"] = { … }` scriptblocks using `Write-Log` / `Invoke-ExternalCommand`
- **`registry.ps1`** gives each module: label, risk, `Category` (sidebar page), `Group` (section heading), options (`RequiresNetwork`, `ConfirmMessage`, `Verify`, `RollbackHint`, `Logo`, `ActionLabel`)
- **Module result contract** (`02-registry.ps1`): body runs as `& $Modules[key] | Out-Null` (stray output discarded); signal failure with `throw` or non-zero `$LASTEXITCODE`. Set `$Script:ModuleExecutionContext["PreviousState"]` to report what changed.

### UI & Theming

- **Sidebar order** lives in `$FmkCategories` (`06-cards.ps1`); a new category needs all three: registry `Category` + `page_<Category>` in XAML + entry in `$FmkCategories`
- **Software page** is a Grid (filter chips above inner `scroll_<Category>` viewer), not a bare `ScrollViewer` like the rest. Visibility toggles on the outer `page_Software` element — the inner scroller must **never** use the `PageScroll` style (its `Visibility=Collapsed` would hide cards permanently). Give it plain scrollbar/padding properties instead.
- **Grid categories** (`$FmkGridCategories`) render entries as logo cards in a `WrapPanel`. Tile size is fixed by `$FmkTileWidth`/`$FmkTileHeight` (height budget: logo + name + description + button must fit). **Never set `WrapPanel.ItemWidth`/`ItemHeight`** — they apply to collapsed children too, leaving holes. Size belongs on the tile (`Width`/`Height`) so `Collapsed` items measure to nothing and drop out of layout.
- **Logos:** Vendored SVG path data in `assets/logos.ps1`, drawn with `Path.Data`. Mini-language prefixed `F1 ` (Nonzero fill rule) because WPF defaults to EvenOdd and hollows marks like FFmpeg's. Missing/unreadable asset → monogram fallback (grid never depends on network). Brand colours failing `Test-FmkLowContrast` against card surface → theme foreground fallback.
- **Theme:** Palette keys in `theme.ps1` substituted into XAML text **before** `XamlReader.Parse`. Never recolor live window via `Window.Resources` (throws / frozen brushes) — rebuild and re-parse instead. Dark chrome set via `DwmSetWindowAttribute` on `SourceInitialized` (`Set-FmkTitleBar`) because setting before `Show` is silently dropped.

### Threading & Scope Traps (load-bearing)

| Rule | Why |
|------|-----|
| Never hand `Dispatcher.Invoke` a scriptblock created in a runspace | Carries that runspace's session state → deadlock |
| Park args on `$global:sync` (synchronized hashtable) | UI-thread delegates built in `New-FmkSyncState` (`ActLogFlush`, `ActStatus`, `ActBusy`, `ActShowDialog`) |
| Click handlers read `$this.Uid` and `$global:sync` | PowerShell scriptblocks don't close over locals |
| Prefix pipeline-polluting calls with `[void]` | `Add_Click`, `ArrayList.Add`, pump timer |
| Take **last** element of `EndInvoke` output | Module's stray output must not become the result object |
| `New-FmkSyncState` auto-exposes every `x:Name` as `$sync.<Name>` | No second list to update |

### Background Runspaces

- One fresh STA runspace per invocation
- Dot-sources the built script with `-FmkLibraryMode`
- Library/UI split is an `if (-not $FmkLibraryMode)` guard, **not** an early `return` (a `return` would abort the dot-sourcing caller)
- Bootstrap path: `$PSCommandPath`, else a copy fetched to `C:\FreeMixKit\temp` for the `irm | iex` case

### Runtime Notes

- **State dir:** `C:\FreeMixKit` (`temp/`, `logs/`, `downloads/`)
- **Session log:** Append-only (don't add `-Force` — library mode runs again for every module)
- **Module results:** Export to CSV on window close

---

## Adding New Modules

### 1. Create the Module Scriptblock

Add to the appropriate `src/modules/*.ps1` file:

```powershell
$Modules["My_NewTool"] = {
    Write-Log "Installing My New Tool..." "Info"
    # Use Install-FmkCatalogApp for WinGet packages:
    Install-FmkCatalogApp -Label "My New Tool" -WingetId "Publisher.Tool" -FallbackUrl "https://example.com/download"
    # Or custom logic with Invoke-ExternalCommand for streaming output
    Write-Log "My New Tool done." "Success"
}
```

### 2. Register the Module

Add to `src/registry.ps1`:

```powershell
Register-Module "My_NewTool" "My New Tool" "Description of what this does." $Modules["My_NewTool"] "None" @{
    Category        = "Software"      # Must match a page_<Category> in ui.xaml
    Group           = "Development"   # Section heading within the page
    ActionLabel     = "Install"       # Button text
    Logo            = "My_NewTool"    # Must exist in assets/logos.ps1
    RequiresNetwork = $true           # If it downloads anything
    # Risk metadata (for Low/Medium/High):
    # RiskReason      = "..."
    # RiskTechnical   = "..."
    # RiskSource      = "..."
    # RiskPrivacy     = "..."
    # RiskLicense     = "..."
    # RiskPrivileges  = "..."
    # RiskEvidence    = "..."
    # RiskReviewed    = "2026-XX-XX"
}
```

### 3. Add Logo Asset (for Software grid)

Add to `src/assets/logos.ps1`:

```powershell
'My_NewTool' = @{
    Color = '#HEXCOLOR'   # Brand colour
    Data  = 'F1 M...'     # SVG path data (F1 = Nonzero fill rule)
}
```

Or for a monogram fallback:

```powershell
'My_NewTool' = @{
    Mono = 'MT'           # 1-4 characters
}
```

### 4. Add XAML Page (if new Category)

Add to `src/xaml/ui.xaml` inside `<Grid x:Name="MainContent">`:

```xml
<ScrollViewer x:Name="page_MyCategory" Style="{DynamicResource PageScroll}">
  <StackPanel x:Name="cards_MyCategory"/>
</ScrollViewer>
```

Add category to `$FmkCategories` in `src/functions/06-cards.ps1` (and `$FmkGridCategories` if it should be a logo grid).

### 5. Rebuild & Verify

```powershell
.\Compile.ps1          # Rebuilds w.ps1
.\Compile.ps1 -Verify  # Validates wiring
```

---

## Risk Ratings

Risk ratings explain potential consequences — technical safety, system changes, privacy, source trustworthiness, licensing/service-policy concerns. **Guidance, not a guarantee of safety.**

| Level | Badge | Criteria |
|-------|-------|----------|
| **Untagged** | *(none)* | No warning warranted — ordinary verified installs (WinGet catalog entries) |
| **LOW** | 🟢 Green | Limited/well-understood risk. Reputable source, predictable changes, no significant security concern. Policy caveats may apply. |
| **MEDIUM** | 🟡 Yellow | Meaningful caveats: significant modifications, admin operations, stability/privacy implications, uncertain compatibility/licensing. |
| **HIGH** | 🔴 Red | Significant risk/serious consequences: unverified executables, disabled security controls, destructive/deceptive behavior. **Licensing alone never makes High.** |

### Key Assessments (rationales in `src/registry.ps1`)

| Module | Risk | Reason |
|--------|------|--------|
| SpotX | Medium | Patches Spotify client; ad-blocking may conflict with Spotify terms |
| Vesktop | Medium | Unofficial Discord client handling account tokens; Discord ToS caveats |
| Spicetify | Medium | Patches Spotify client; scheduled-task install; ToS risk separate |
| FixResolution | Medium | Restarts graphics driver + applies EDID override; display may go blank |
| MSGameBarFix | Medium | Small-repo script changing Game Bar behavior; limited verification |
| WinUtil | Low | Established open utility; inner operations vary and need admin |
| Winhance | Low | Established Windows utility; review options before applying |
| GenP | High | Unverified mirror executables + Adobe-terms impact |
| MAS / IAS | High | Bypass Microsoft/IDM licensing; remote code execution as admin; explicit confirmations required |

> **Hover any badge** in the UI for the reason and dimensions, or click **Risk ratings?** for the full criteria.

---

## Safety & Transparency

- **Read before running:** Some modules download and execute third-party tools or activation scripts. Inspect `w.ps1` (or the `src/` sources) before use.
- **Activation modules** (MAS, IAS, GenP, etc.) may violate license terms or local law — use only where legal and at your own risk. High-risk modules require explicit confirmation; the Software queue shows a combined Medium/High review before installing.
- **No security bypasses:** The queue never bypasses UAC, signature checks, or Windows security controls. It never retries forever, and never swaps in an unrelated installer after a failure.
- **System modifications:** Several actions modify system settings, services, and the registry. Ensure you understand each option you select.
- **Source code is plain PowerShell:** Inspect or edit to suit your environment.
- **Verify before you run:** `irm https://raw.githubusercontent.com/catsmoker/FreeMixKit/main/w.ps1` downloads the file — save and review it first when practical. Check the release channel rather than executing an unverified replacement.

---

## Recommended Practice

- Try high-risk modules on a non-production machine first
- Keep backups before registry or security changes
- Run one change at a time and verify system behavior between steps

---

## Troubleshooting

| Issue | Resolution |
|-------|------------|
| **"Administrator rights are required"** | Right-click PowerShell → **Run as Administrator**, then run the script |
| **Window opens but is blank / crashes immediately** | Ensure you're on STA. Try: `powershell -STA -File .\w.ps1` |
| **"Cannot overwrite variable Host because it is read-only"** | You likely have a variable named `$host` in your code. Rename it (e.g., `$cardPanel`). PowerShell variable names are case-insensitive. |
| **XAML parse error on startup** | Check for `--` inside XAML comments or bare `&` characters. Run `.\Compile.ps1 -Verify` — it catches XML well-formedness. |
| **Module card renders but clicking does nothing ("Unknown action")** | `Register-Module` points to a `$Modules` key that doesn't exist. Run `.\Compile.ps1 -Verify` — it catches this. |
| **Software grid has empty holes after filtering** | You likely set `WrapPanel.ItemWidth`/`ItemHeight`. Move sizing to the tile (`Width`/`Height`) instead. |
| **Theme colours not applying** | Palette tokens (`%%Token%%`) must be substituted **before** `XamlReader.Parse`. Don't write to `Window.Resources` at runtime. |
| **Winget not found / install fails** | Script auto-installs Winget via `Add-AppxPackage`. Ensure internet access. Try `winget --version` manually. |
| **Background task hangs / UI freezes** | Background runspaces use a single STA runspace per invocation. Check `Test-FmkTaskCancellationRequested` polling in long loops. |

### Logs

- **Session log:** `C:\FreeMixKit\logs\FreeMixKit_Session.log` (append-only)
- **Module results:** Exported to CSV on window close
- **Share Logs:** Use the **Share Logs** button on the About tab to save sanitized diagnostics for bug reports

---

## Contributing

1. **Fork** the repository
2. **Create a branch** for your change: `git checkout -b feature/my-change`
3. **Edit files under `src/` only** — never edit `w.ps1` or `src/w.template.ps1` build markers directly
4. **Rebuild:** `.\Compile.ps1`
5. **Validate:** `.\Compile.ps1 -Verify` (must pass)
6. **Test:** Run `.\w.ps1` and exercise your changes
7. **Commit** with a clear message
8. **Open a Pull Request**

### Code Style

- Follow existing patterns in `src/functions/` and `src/modules/`
- Use `Write-Log` for all output (levels: `Info`, `Warn`, `Error`, `Success`)
- Use `Invoke-ExternalCommand` for streaming external process output
- Respect `Test-FmkTaskCancellationRequested` in long-running operations
- Never assign to read-only automatic variables (`$host`, `$HOME`, `$PSScriptRoot`, etc.)
- Keep StrictMode 1.0 compatible (don't access missing keys on `$null`)

### Adding a New Module

See [Adding New Modules](#adding-new-modules) above.

---

## Verify Before Running

The `irm | iex` one-liner downloads and executes code from GitHub. For any production or sensitive environment:

```powershell
# 1. Download without executing
irm https://raw.githubusercontent.com/catsmoker/FreeMixKit/main/w.ps1 -OutFile w.ps1

# 2. Review the file
notepad w.ps1   # or your preferred editor

# 3. Run locally after review
powershell -NoProfile -ExecutionPolicy Bypass -File .\w.ps1
```

---

## License

This project is licensed under the **MIT License**. See [LICENSE](LICENSE) for details.

---

*FreeMixKit — Keep your Windows toolbox in one place.*