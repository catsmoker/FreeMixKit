# Development

Source of truth is `src/`; root `w.ps1` is generated. Never edit `w.ps1` directly.

```
src/
├── w.template.ps1     entry point: setup, elevation/STA relaunch, boot
├── xaml/ui.xaml       the WPF window
├── theme/theme.ps1    colour tokens and Set-FmkUITheme
├── assets/logos.ps1   vendored logo path data for the Software grid
├── functions/         UI plumbing, numbered by load order:
│   ├── 01-helpers.ps1 registry-free helpers (winget, network, external commands)
│   ├── 02-registry.ps1 Register-Module plus risk-metadata defaults
│   ├── 03-ui-log.ps1  logging, status, busy state
│   ├── 04-runspace.ps1 background runspace plus single-module guard
│   ├── 05-dialogs.ps1 confirm dialogs (UI-thread delegates)
│   ├── 06-cards.ps1   nav, cards, tiles, filter, multi-select state
│   ├── 07-boot.ps1    sync state, wiring, work-area clamp, launch
│   ├── 08-risk.ps1    risk criteria text, tooltips, help dialog
│   └── 09-queue.ps1   Software installation queue
├── modules/           module bodies grouped by area (one key per app)
└── registry.ps1       registration: label, risk + dimensions, Category/Group
```

## Adding software

1. Add the install body as `$Modules["Key"] = { ... }` in the matching file
   under `src/modules/` (Winget ID verified with
   `winget search --id <Id> --exact` first).
2. Register it in `src/registry.ps1` with `Category = "Software"` and one of the
   consolidated groups: Browsers, Messaging, Media, File Sharing & Downloads,
   Development, Video & Image Tools, Gaming, System & Maintenance, Productivity.
3. Add `Logo = "Key"` plus a vendored entry in `src/assets/logos.ps1`
   (brand mark `Color` + `Data`, raster `Png`, `Shapes`, or `Mono` last resort).
4. Risk: leave `Risk = "None"` for ordinary verified installs. Only set
   Low/Medium/High with `RiskReason` plus dimensions and `RiskReviewed`; High
   also needs `ConfirmMessage`.
5. Run `.\Compile.ps1 -Verify`, then `.\tests\Test-Fmk.ps1`, then rebuild.

## Threading rules (load-bearing)

- Never hand `Dispatcher.Invoke` a runspace scriptblock; park args on
  `$global:sync` and invoke the UI-thread delegates from `New-FmkSyncState`.
- Click handlers read `$this.Uid` and `$global:sync` only.
- Prefix pipeline-polluting calls with `[void]`; take the last `EndInvoke` item.
