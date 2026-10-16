# AGENTS.md — FreeMixKit

Single-file WPF toolbox (Windows PowerShell 5.1 / pwsh). `src/` is the source of truth; root `w.ps1` is a **generated** distributable (committed so `irm <url>/w.ps1 | iex` keeps working). **Never edit `w.ps1` or `src/w.template.ps1`'s build markers directly.**

## Build — the only workflow

```powershell
.\Compile.ps1            # rebuild .\w.ps1 after ANY change under src/
.\Compile.ps1 -Verify    # validate only, writes nothing — run before committing
```

`Compile.ps1` concatenates template + `functions/` + `theme/` + `assets/` + `modules/` + `registry.ps1` + XAML via literal `.Replace()` of `#>>NAME<<#` tokens, and enforces (fails the build):

- every source file parses clean with the real PowerShell parser;
- no assignment to read-only automatic variables (`$host`, `$HOME`, `$PSScriptRoot`, … — case-insensitive). Name panel/locals `$cardPanel`, never `$host`;
- every `Category=` in `registry.ps1` has a matching `x:Name="page_<Category>"` in `src/xaml/ui.xaml`;
- every `%%Token%%` in the XAML has a matching key in `src/theme/theme.ps1`;
- `ui.xaml` is well-formed XML. It is parsed as XML long before WPF sees it, and the failure mode is otherwise nasty: an illegal construct (a `--` inside a comment, a bare `&`) builds a working `w.ps1` that throws from `XamlReader.Parse` at startup, after the console has been hidden. So never write a `--` dash inside an XAML comment;
- every entry in `src/assets/logos.ps1` is either a brand mark (`Color` + `Data`) or a monogram (`Mono`), and every `Logo = "X"` in `registry.ps1` names an entry that exists;
- every `Register-Module "K" … $Modules["M"]` points at a `$Modules["M"] =` definition that exists, and no key is registered twice. A registration pointing at a missing module is the nastiest failure of the lot: the card renders, clicks, and only then reports "Unknown action". Modules defined but never registered are reported as a note, not a failure.

Output is CRLF, UTF-8 without BOM. There are no tests, linters, or CI — `-Verify` is the check.

## Architecture (load order matters)

`src/w.template.ps1` = setup → elevation/STA relaunch → `Set-StrictMode 1.0` → helpers → theme → modules → registry → XAML → boot. Runtime uses **StrictMode 1.0, not Latest** (UI code legitimately touches missing keys/`$null`); keep it that way.

- `functions/01..07` load alphabetically: helpers → registry fns → ui-log → runspace → dialogs → cards → boot.
- Modules (`modules/*.ps1`) only assign `$Modules["Key"] = { … }` scriptblocks using `Write-Log` / `Invoke-ExternalCommand`; `registry.ps1` gives each one label, risk, `Category` (sidebar page), `Group` (section heading), and options (`RequiresNetwork`, `ConfirmMessage`, `Verify`, `RollbackHint`, `Logo`, `ActionLabel`).
- Module result contract (`02-registry.ps1`): the body runs as `& $Modules[key] | Out-Null`, so stray output is discarded; signal failure with `throw` or a non-zero `$LASTEXITCODE`. Set `$Script:ModuleExecutionContext["PreviousState"]` to report what changed.
- Sidebar order lives in `$FmkCategories` (`06-cards.ps1`); a new category needs all three: registry `Category` + `page_<Category>` in XAML + entry in `$FmkCategories`. The Software page is a Grid (filter chips above an inner `scroll_<Category>` viewer), not a bare ScrollViewer like the rest: visibility is toggled on the outer `page_` element, so the inner scroller must never use the `PageScroll` style (its `Visibility=Collapsed` would hide the cards for good) — give it plain scrollbar/padding properties instead. A category listed in `$FmkGridCategories` renders its entries as logo cards in a `WrapPanel` instead of list rows — the grid's card size is fixed by `$FmkTileWidth`/`$FmkTileHeight` (the height carries a documented budget: logo + name + description + button must fit inside it, so raising the description or name line count means raising the tile).
- **Never set `WrapPanel.ItemWidth`/`ItemHeight` for that grid.** They are applied to every child unconditionally, `Collapsed` ones included, so a card the search has filtered out keeps a full-size empty slot: the survivors stay in their original columns and the grid fills up with holes. The size belongs on the tile (`Width`/`Height`), because a `Collapsed` Border measures to nothing and drops out of the layout, which is what lets a filtered grid reflow from the top-left. `Compile.ps1` cannot catch this — it is a layout behaviour, not a wiring mistake.
- Every module installs only what its own card promises. The old single "Dev Choice" bundle was split into one module per component; do not re-merge them, and do not make one card install another's packages.
- Logos are vendored SVG path data in `assets/logos.ps1`, drawn with `Path.Data`. The mini-language is prefixed `F1 ` (Nonzero fill rule) because WPF otherwise defaults to EvenOdd and hollows out marks such as FFmpeg's. A missing or unreadable asset falls back to a monogram, so the grid never depends on the network. Brand colours that fail `Test-FmkLowContrast` against the card surface fall back to the theme foreground.
- Theme: palette keys in `theme.ps1` are substituted into the XAML text **before** `XamlReader.Parse`. Never recolor a live window via `Window.Resources` (throws / frozen brushes) — rebuild and re-parse instead. The palette follows WinUtil's own `config/themes.json` values; the dark chrome is set with `DwmSetWindowAttribute` on `SourceInitialized` (`Set-FmkTitleBar`) because setting it before `Show` is silently dropped.
- Background modules run on **one fresh STA runspace per invocation**, which dot-sources the built script with `-FmkLibraryMode`. The library/UI split is an `if (-not $FmkLibraryMode)` guard, not an early `return` (a `return` would abort the dot-sourcing caller). Bootstrap path is `$PSCommandPath`, else a copy fetched to `C:\FreeMixKit\temp` for the `irm | iex` case.

## Threading / scope traps (all load-bearing)

- Never hand `Dispatcher.Invoke` a scriptblock created in a runspace — it carries that runspace's session state and deadlocks. Park args on `$global:sync` (a synchronized hashtable) and invoke the UI-thread delegates built in `New-FmkSyncState` (`ActLogFlush`, `ActStatus`, `ActBusy`, `ActShowDialog`).
- Click handlers must read `$this.Uid` and `$global:sync`; PowerShell scriptblocks do not close over locals.
- Prefix pipeline-polluting calls with `[void]` (`Add_Click`, `ArrayList.Add`, pump timer) and take the **last** element of `EndInvoke` output — a module's stray output must not become the result object.
- `New-FmkSyncState` auto-exposes every `x:Name` in the XAML as `$sync.<Name>` — no second list to update.

## Runtime notes

- Requires Windows + STA + admin; the script self-elevates in one relaunch (`$env:FMK_RELAUNCHED` is the loop guard, inherited for the `irm | iex` path). `powershell.exe` is the fallback when a host rejects `-STA`.
- State dir is `C:\FreeMixKit` (`temp/`, `logs/`, `downloads/`); session log is append-only (don't add `-Force`), module results export to CSV on window close.
