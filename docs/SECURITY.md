# Security

- Prefer official sources: Winget with exact IDs, vendor download pages, and
  official GitHub releases with SHA-256 verification where the publisher
  provides it (e.g. ZapFast).
- Risk ratings are guidance, not guarantees. No audit is claimed; evidence links
  and review dates live in `src/registry.ps1` so contributors can re-check.
- High-risk modules (MAS, IAS, GenP) require an explicit confirmation naming
  the specific licensing, provenance, and privilege concerns.
- The Software queue shows one combined Medium/High review before installing,
  continues past recoverable failures, records each result, supports cancel for
  pending items, and never retries forever or swaps in an unrelated installer.
- Never bypass UAC, signature verification, or Windows security controls for
  bulk-install convenience.
- Do not store credentials or tokens in source; do not suppress errors to hide
  failures; back up or use version control before substantial refactors.
- Review unfamiliar scripts before executing them, including any skill or
  third-party helper that claims to improve security.
