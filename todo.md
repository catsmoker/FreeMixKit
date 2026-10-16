# FreeMixKit TODO

> Note (2026-10-16): no `todo.md` existed in the repo when the targeted
> UI/catalog update round started, so this file was created then to track it.

## Targeted UI and Software Catalog Updates (done 2026-10-16)

- [x] Taskbar icon uses the project logo (`freemixkit_icon.ico`, blue
      panes-and-gear mark): vendored as base64 in `src/assets/icon.ps1`
      (`$FmkAppIconB64`, multi-size 16..256 PNG-compressed, alpha preserved)
      with a temp-materialize fallback in `src/w.template.ps1` so the
      `irm | iex` path (no sibling file) also shows the logo; window-icon
      decode in `src/functions/07-boot.ps1` uses `BitmapCacheOption.OnLoad`
      so the temp file is never left locked.
- [x] Risk tags per the requested table, nothing else moved:
  - [x] Utilities / Winhance: Medium -> Low
  - [x] Software / .NET: Medium tag removed (None, no risk detail)
  - [x] Software / Notepad++: Medium tag removed (None, no risk detail)
  - [x] Software / SpotX: Low -> Medium (badge + tooltip via existing model)
  - [x] Software / Legcord: Low -> Medium (badge + tooltip via existing model)
- [x] 7-Zip replaced by NanaZip in Software / Productivity: new
      `App_NanaZip` module (`M2Team.NanaZip`, verified via
      `winget search/show --exact`: NanaZip 7.0.1843.0, M2Team/Kenji Mouri,
      MIT, https://github.com/M2Team/NanaZip), official 256px unplated logo
      vendored in `src/assets/logos.ps1`; old `Dev_SevenZip` module,
      registration, and logo entry removed; existing 7-Zip installs untouched.
- [x] `src/assets/icon.ps1`, `tests/Test-Fmk.ps1` (69 assertions), README risk
      and catalog sections updated; `.\Compile.ps1` + `w.ps1` rebuilt.

## Deferred / known limitations (carried over)

- [x] Replace root `freemixkit_icon.ico` with the multi-size build — DONE this
      round (7 PNG entries 16..256, alpha preserved) once the file lock
      cleared. Note: mid-round the root file was observed replaced by an
      opaque black-background variant of the same artwork (unknown actor;
      none of this round's scripts write that shape). The observed bytes were
      backed up to temp, and the root file now holds the verified transparent
      multi-size build derived from the original committed artwork.
- [ ] Refresh `screenshot.png` / `screenshot-software.png` (stale: predate the
      selection queue, NanaZip, overlay, and About tab).
- [ ] Live taskbar screenshot verification on a machine where the window can
      be shown interactively (headless sessions verified the exact
      `Window.Icon` code path instead).
- [ ] Consider `winget show` re-verification of `M2Team.NanaZip` if its
      version drifts; the module itself is version-agnostic.

## Progress UI, Task Locking, Selection Fixes, About Tab (done 2026-10-16)

- [x] Shared progress bar: determinate step/total (`Installing X (n of m)`)
      from measured percentages via `Set-FmkBusy -Percent` /
      `Get-FmkProgressRequest`; indeterminate otherwise (never fabricated).
      Queue and single tasks share the one mechanism.
- [x] Global busy state + loading overlay (`BusyOverlay`): dims and locks the
      sidebar/pages (mouse swallowed, tab stops dropped) while log output,
      status, shared progress, and a queue-only Cancel stay visible; overlay
      and lock always clear through `Set-FmkBusyNow`, reached from `finally`
      paths. Long work already ran on background runspaces; quick
      interactions (select/filter/navigate) never set Busy.
- [x] Selection bug root cause fixed: the tile border was driven by both the
      `selected` trigger and a later `IsKeyboardFocused` recolor, so a
      deselected-but-focused tile kept a highlight until focus moved. Focus
      now uses its own outside-the-tile adorner (`TileFocusVisual`); the
      border always agrees with the ArrayList state. Order-preserving queue
      behavior intact; selection locked during tasks.
- [x] Links tab renamed to About (sidebar, XAML pages, registry, tests,
      README): product description, implemented-features list, preserved
      repository/website/shortcut cards, Report an Issue button (official
      Issues URL derived from `$FmkRepoUrl`), Share Logs (sanitized `.txt`
      export via save dialog, manual attach, never uploaded).
- [x] Low badge renders in its own theme color via `Test-FmkRiskBadgeVisible`
      (Low/Medium/High shown, None hidden); untagged items stay untagged.
- [x] Tests: 125 assertions green, plus headless runtime checks (real toggle/
      filter/progress/sanitizer/export functions) and built-window structural
      checks (all named elements resolve, overlay hidden by default).

## About Text + Full PowerShell Output (done 2026-10-16)

- [x] About Features mojibake fixed at the root: the `•` bullet (the only
      non-ASCII in `src/`) was decoded as Windows-1252 by `Get-Content -Raw`
      in `Compile.ps1` and re-encoded as UTF-8 (`C3 A2 E2 82 AC C2 A2` in the
      shipped script). `Read-SourceFile` now reads explicit UTF-8, and the
      bullet is an ASCII hyphen matching the rest of the app (BOM-less 5.1
      script parsing mangles the glyph even when built correctly, verified
      empirically). Sources enforced ASCII-only by test.
- [x] Output panel now streams real process output: `Invoke-ExternalCommand`
      rewritten around polled pipe reads (`BeginRead` + 200ms wait, no
      scriptblock event handlers so no runspace deadlock), UTF-8 decoders
      kept across polls, CR/LF segmentation, 500-line/tick burst pacing with
      ordered hold-back, stdout white / stderr amber, exit code still decides
      success. Module bodies run merged-stream (`*>&1`) with per-record
      classification (success/error/warning/verbose/debug/information;
      progress records update the status line throttled, never the log) while
      terminating errors still propagate and the result contract is unchanged.
      Winget refresh/upgrade lines stream instead of being discarded.
- [x] Tests: 163 assertions green covering multi-line order, stderr + exit
      code, all PS streams, byte-identical Unicode, incremental arrival
      observed mid-run by a probe job, 250-line volume in order, burst pacing,
      progress no-spam, real `winget --version` output (`v1.29.380`), plus an
      off-screen render of the real About panel (clean hyphen bullets, no
      mojibake) and of the real log pane showing streamed stdout/stderr lines.
- [x] `w.ps1` rebuilt (5285 lines). Predictable-display caveat: non-ASCII
      glyphs may still show as boxes in some text-capture layers; the byte
      assertions prove pipeline fidelity regardless.

## Real Output + Percentage Progress (done 2026-10-16)

- [x] Traced a real install end to end and measured winget under a pipe:
      `winget download` emits ~6 plain phase lines (`Found …`, download URL,
      hash verification, result path) with ~20 s of silence mid-download and
      never a percentage; stderr is always empty; a no-op `winget install`
      exits `0x8A15…` with `already installed` text. The streaming pipeline
      itself was proven live (chunks arrive mid-run); the sparseness is the
      source's doing, not lost output.
- [x] Genuine percentages, never fabricated: `Get-FmkObservedPercent` parses
      `NN%` only from real output text; queue overall =
      (completed + current fraction) / total via `Get-FmkQueueOverallPercent`;
      lone tasks show their own observed percent. The numeric percentage now
      renders beside the bar (`ProgressPercent`); indeterminate work shows no
      number and the overlay states `waiting for installer progress
      information`. Terminal states persist until the next task: green 100%
      on real success, red/amber frozen at the last measured value on
      failure/cancel.
- [x] App chatter vs real output made structural: framework lifecycle lines
      (`Executing:`, `Requested:`, selection echoes, queue step headers) log
      muted gray via the new `Muted` color key; tool output keeps
      white/amber/red and result verdicts keep their colors.
- [x] Tests: 201 assertions green, including parser/overall-math units,
      task-manager behavior driven on a real parsed (never shown) window,
      muted-color registration, a real `winget download` through
      `Invoke-ExternalCommand`, and arrival-timestamp proof of streaming.
- [x] `w.ps1` rebuilt (5453 lines). Installers that emit no console progress
      (winget, GUI installers, scheduled-task launchers) expose no percentage
      by construction; the UI reports the measured floor plus the waiting
      message instead of inventing numbers.

## System Hardware + Windows Diagnostic Report (done 2026-10-16)

- [x] Replaced the three-line report with `Get-FmkSystemReportData`, one
      shared collection service (`src/functions/11-sysreport.ps1`) covering
      Windows, CPU, RAM, GPU, storage, motherboard/firmware, network,
      battery, audio/displays/peripherals, and runtime/virtualization from
      inbox CIM/WMI, registry, and read-only commands. No installs, no
      drivers, no config changes, no per-field processes.
- [x] Every source is optional: per-field try/catch with `Unavailable` /
      `Not supported` / `Access denied` values plus per-section warnings; a
      failed section degrades instead of stopping the report. Zeros and
      $false pass through as real values, never blanked.
- [x] System tab shows overview, per-section expandable cards, Refresh /
      Copy / Export buttons, a sensitive-data checkbox, and collection
      status. Collection runs on a background runspace (120 s watchdog)
      under the existing busy overlay; first tab visit auto-loads once.
      The System Report card now exports the same full report.
- [x] Privacy: no product keys collected; serials/MACs/IPs hidden by default
      (`[hidden]` on screen, `[withheld]` in default exports, full values
      only after an explicit confirm); computer/user names shown labeled as
      potentially sensitive; nothing uploaded, save path confirmed.
- [x] Tests: 252 assertions green (formatters, type maps, redaction rules,
      zero-vs-missing, multi-disk shaping, warnings, JSON parity).
      Verified live on this machine (non-admin): 11 sections, ~200 lines,
      6 s collection; build/CPU/RAM spot-checked against independent
      queries; serials confirmed withheld. Fields Windows would not expose
      here: DeviceGuard status, NUMA topology, shared GPU memory, DirectX
      feature level, WDDM model, HDR/VRR, SMART (needs admin), TPM
      readability, design battery capacity.
- [x] `w.ps1` rebuilt. Screenshots predate this feature (see carried-over
      item above).

## System Report File Correction (done 2026-10-16)

- [x] Corrected course per clarification: the report belongs in the
      generated `.txt` file, not the UI. Removed the System-tab report
      dashboard (panel, expandable section controls, Refresh/Copy/Export
      buttons, auto-load hook, `ActReportReady` delegate,
      `Invoke-FmkReportRunspace`) and restored the plain button workflow:
      System tab, System Report card, Run generates the file and opens it.
      Kept the reusable engine untouched: collectors, formatters, redaction
      rules, and the `SystemReport` module that writes the file.
- [x] File carries version (`Report format version: 1 (FreeMixKit v7)`) and
      timestamp; serials/MACs/IPs withheld by default; nothing uploaded.
- [x] Verified by executing the real Run action path
      (`Invoke-ModuleAction 'SystemReport'`, the exact call the background
      runspace makes): Success in ~6 s, 203-line/11-section file with real
      values, per-source notes where Windows withheld data, and only the
      spawned viewer closed afterwards.
- [x] Tests: 247+ assertions green, including a guard that none of the
      removed UI helpers reappear in the build. Screenshots still predate
      this feature family.

## System Report Encoding Fix (done 2026-10-16)

- [x] Root cause, reproduced byte-for-byte: the report was written as BOM-less
      UTF-8, which carries no encoding signal. Any reader decoding those bytes
      as UTF-16 renders plain ASCII as CJK (`46 72` -> U+7246, `65 65` ->
      U+6565, exactly the reported `...` shape). No in-repo UTF-16 misuse was
      found; the fix is at the single writer per the task prescription.
- [x] `Write-FmkSysReportFile` helper writes explicit UTF-8 with BOM; the
      System Report module uses it. The BOM makes the file self-describing
      for Notepad and every Windows reader, while PowerShell/.NET readers
      detect and strip it either way.
- [x] Tests: 263 assertions green, including BOM presence, exact UTF-8
      round-trip of heading/`®`/CJK/punctuation, all section headings, and
      the UTF-16-misread sanity check.
- [x] Verified on the real Run action: regenerated file starts `239,187,191`,
      heading reads `FreeMixKit System Report` identically via default,
      auto-detect, and explicit-UTF-8 readers, all 11 sections with live
      values, `®` intact. Only spawned viewer closed afterwards.
- [x] `w.ps1` rebuilt. Note: the product heading has always been the brand
      spelling `FreeMixKit System Report` (matches the window title).

## System Report Stream-Encoding Fix (done 2026-10-16)

- [x] Reproduced real corruption in-session: `sfc.exe` output arrived as
      NUL-interleaved text in the session log because the pipe poller decoded
      every stream as UTF-8. UTF-16LE console output read as single bytes is
      the same mismatch family as the reported CJK mojibake (ASCII bytes read
      as UTF-16 pairs render as `...`-shaped characters; verified the pair
      mapping against the report bytes).
- [x] `Invoke-ExternalCommand` now detects each pipe's encoding from its
      leading bytes (UTF-16/UTF-8 BOM honored and skipped; NUL-interleaved
      prefix selects UTF-16; otherwise UTF-8; short prefixes wait, end of
      stream falls back to UTF-8) with a sticky per-stream decoder, so split
      multibyte characters still survive chunk boundaries. No scriptblock
      event handlers, so the runspace-deadlock rule still holds.
- [x] Tests: 277 assertions green, including detection units (both BOMs,
      NUL pattern, multibyte guard, short-prefix wait, EOF fallback), a
      no-spurious-warning guard on clean runs, and a real `cmd /u` UTF-16
      process end to end through the product path with exact stdout/stderr
      lines and zero NUL leakage.
- [x] Fixed empty-array parameter binding in the stream decoder final flush
      (had emitted a spurious `Stream capture degraded` warning on every
      command and risked dropping a decoder tail; caught by the new guard).
- [x] Real Run action re-verified after the change: BOM intact, heading
      identical across all readers, 11 sections, live values.
- [x] `w.ps1` rebuilt. Known limitation: a UTF-16 stream carrying only
      non-ASCII text with no BOM falls back to UTF-8; Windows console tools
      in practice emit BOMs or ASCII-range text, so this path is documented
      rather than guessed at.
