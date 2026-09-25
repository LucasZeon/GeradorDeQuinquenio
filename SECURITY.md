# Security Policy

## Scope

This repository is a public, sanitized portfolio project. It must never contain
real personal data, credentials or internal infrastructure details. All data in
it is synthetic.

## Reporting a problem

If you find something that looks like real data, a credential, a token or an
internal identifier, or you discover a vulnerability in the code:

1. **Do not** open a public issue with the details.
2. Use GitHub's private reporting: **Security → Report a vulnerability** on this
   repository.
3. Describe where you found it (file and line) — please do not copy the value
   itself into the report.

I will remove the content and, if needed, rewrite the history as soon as
possible.

## Supported versions

Only the latest version on the `main` branch is supported.

## Safeguards in this repository

- Allowlist-based `.gitignore` (new files are ignored unless explicitly allowed).
- `tools/check_public_safety.py`, run by a pre-commit hook and by CI.
- CodeQL analysis and Dependabot updates.
