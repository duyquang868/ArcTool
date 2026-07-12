---
name: closed-dossier-policy
description: Section 7 of CLAUDE.md should only track recently closed features; older closed features must move to English-only dossier files under .Dossier.
type: feedback
originSessionId: 308b02c2-3cc3-4e06-b240-90ead7920e0e
---
Use `CLAUDE.md` section 7 only as a short recent-closure record, not as a long-term archive.

**Why:** The user wants `CLAUDE.md` to stay lean and operational, while older or deeper technical history should live in separate dossier files for future bug fixing.

**How to apply:** When a feature is newly closed, create or update one dedicated dossier under `.Dossier`, keep one dossier per feature or clearly bounded subsystem, and leave only a short summary plus pointer in section 7. When it is no longer recent, remove the section-7 summary and rely on the dedicated dossier file. Dossier filenames and dossier contents must both be written in English.
