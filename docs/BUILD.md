# Build and release

## Build

```powershell
.\Compile.ps1            # rebuild .\w.ps1 after any change under src/
.\Compile.ps1 -Verify    # validate only, writes nothing
powershell -NoProfile -ExecutionPolicy Bypass -File .\tests\Test-Fmk.ps1
```

`Compile.ps1` concatenates template + `functions/` + `theme/` + `assets/` +
`modules/` + `registry.ps1` + XAML via literal `.Replace()`, then enforces:

- every source file parses with the real PowerShell parser;
- no assignment to read-only automatic variables;
- every `Category=` has a matching `page_<Category>` in the XAML;
- every `%%Token%%` has a theme key; XAML is well-formed XML;
- every logo entry is shaped correctly and every `Logo =` resolves;
- every `Register-Module` points at a defined module, no duplicates;
- every `Risk` is None/Low/Medium/High, Medium/High set `RiskReason`,
  High sets `ConfirmMessage`.

Output is CRLF, UTF-8 without BOM. `-Verify` is the gate; do not commit when it
fails. `tests/Test-Fmk.ps1` adds metadata, ordering, search, icon, sizing, and
queue-safety checks (50 assertions, no WPF or admin required).

## Distribution

The distributable is root `w.ps1`, committed so the one-liner keeps working:

```powershell
irm https://raw.githubusercontent.com/catsmoker/FreeMixKit/main/w.ps1 | iex
```

Prefer download-and-review when practical: fetch the file, read the risk
entries and module bodies that matter to you, check the release or commit the
file came from, then run the local copy. The script self-elevates once
(`$env:FMK_RELAUNCHED` loop guard) and relaunches with `-STA` for WPF; admin is
required by individual operations, not bypassed for convenience. Never disable
Windows security controls to simplify installation.
