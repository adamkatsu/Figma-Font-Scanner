# Font Scanner — Version History

This document tracks the evolution of the **Font Scanner** plugin from initial creation to the present, based on commit history on [GitHub](https://github.com/adamkatsu/Figma-Font-Scanner).

| Info | Detail |
|------|--------|
| Repository | [adamkatsu/Figma-Font-Scanner](https://github.com/adamkatsu/Figma-Font-Scanner) |
| Plugin ID | `1603010877465738852` |
| Stack | TypeScript, Figma Plugin API (`documentAccess: dynamic-page`) |
| Contributors | [adamkatsu](https://github.com/adamkatsu), [kemonn98](https://github.com/kemonn98) |

---

## Version Summary

| Version | Date | Commit | Status |
|---------|------|--------|--------|
| **1.0.0** | Dec 1, 2025 – Dec 16, 2025 | `05266a5` → `ee21e1a` | Released |
| **1.1.0** | Feb 16, 2026 | `18a7c7b` | Released |
| **1.2.0** | Feb 24, 2026 | `4a8938a` | Released |
| **1.2.1** | Jul 10, 2026 | `ae0cfe4` | Released |
| **1.2.2** | Jul 10, 2026 | `41ac074` | Released |
| **1.2.3** | Jul 10, 2026 | `b22bf26` | Released |
| **1.2.4** | Jul 10, 2026 | `e782d9f` | Released |
| **1.2.5** | Jul 10, 2026 | — | Released |

---

## v1.2.5 — Released

**Date:** Jul 10, 2026

### Changes

- **Commit message convention** – `git push` releases use descriptive subjects (`fix:` / `feat:` / `chore:` + what changed), not `release: vX.Y.Z`.
- **Hash backfill** – Previous release commit resolved with `git log -S '"version": "X.Y.Z"' -- package.json`.
- Updated `git-push-release.mdc`, `AGENTS.md`, and `VERSION.md` maintenance notes.

**Files changed:** `package.json`, `package-lock.json`, `ui.html`, `README.md`, `VERSION.md`, `AGENTS.md`, `.cursor/rules/git-push-release.mdc`

---

## v1.2.4 — Released

**Commit:** `e782d9f` — *release: v1.2.4* (Jul 10, 2026)

### Changes

- **UI version label** – Footer `.footer-version` in `ui.html` updated to match release (`v1.2.4`); was stale at `v1.2.0`.
- **Release workflow** – `ui.html` is now a mandatory version bump target on every `git push` release.
- Updated `git-push-release.mdc`, `AGENTS.md`, and `VERSION.md` maintenance notes.

**Files changed:** `ui.html`, `package.json`, `package-lock.json`, `README.md`, `VERSION.md`, `AGENTS.md`, `.cursor/rules/git-push-release.mdc`

---

## v1.2.3 — Released

**Commit:** `b22bf26` — *release: v1.2.3* (Jul 10, 2026)

### Changes

- **Release workflow fix** – Removed `git commit --amend` for commit hashes; a release commit cannot self-reference its own hash after amend.
- **Backfill on next release** – Previous version's `—` commit column is filled via `git log` on the `package.json` version bump.
- Updated `git-push-release.mdc` and `AGENTS.md` accordingly.

**Files changed:** `.cursor/rules/git-push-release.mdc`, `AGENTS.md`, `VERSION.md`, `package.json`, `package-lock.json`, `README.md`

---

## v1.2.2 — Released

**Commit:** `41ac074` — *release: v1.2.2* (Jul 10, 2026)

### Changes

- **Release workflow fix** – `git-push-release.mdc` uses `TBD` placeholders for commit hash, then `git commit --amend` before push so `VERSION.md` is never left uncommitted.
- **Clean tree requirement** – Step 6 verifies `git status` is clean after push; no follow-up docs commits.
- Updated `AGENTS.md` release checklist to match amend-before-push flow.
- Corrected v1.2.1 commit references in `VERSION.md` to `ae0cfe4`.

**Files changed:** `.cursor/rules/git-push-release.mdc`, `AGENTS.md`, `VERSION.md`, `package.json`, `package-lock.json`, `README.md`

---

## v1.2.1 — Released

**Commit:** `ae0cfe4` — *release: v1.2.1* (Jul 10, 2026)

### Changes

- **Fix scan selection scope** – Store container references (`SceneNode[]`) at scan time; reuse until next rescan so Select Fonts, replace, and other scoped operations work after the canvas selection changes to text layers.
- **Fix dynamic-page compatibility** – Use direct node references instead of `figma.getNodeById()` (blocked under `documentAccess: dynamic-page`).
- Deleted scanned frames are filtered via `node.removed`.
- Added `VERSION.md`, `AGENTS.md`, and Cursor rules (`.cursor/rules/`) for release workflow and English-only docs.
- Updated `package.json` version and description.

**Files changed:** `code.ts`, `package.json`, `package-lock.json`, `README.md`, `VERSION.md`, `AGENTS.md`, `.cursor/rules/`

---

## v1.2.0 — Released

**Commit:** `4a8938a` — *v1.2.0 new features update* (Feb 24, 2026)

### New features

- **Size list by line height & letter spacing** — Columns for Size · Line height · Letter spacing · Layers to spot inconsistent typography combinations.
- **Edit line height & letter spacing** — Click a value in the size list to change it (line height can be a number or Auto); press Enter to apply to matching layers.
- **Select by weight** — Click a weight row to select all layers with that font and style.
- **Select by size variant** — Click a size row (or the "Layers" cell) to select layers with the same size, line height, and letter spacing.
- **Replace line height & letter spacing** — Bulk replace per typography variant.

### Technical changes

- `FontUsageDetails.sizeVariants` replaces simple `sizes`; variant key: `fontSize_lineHeight_letterSpacing`.
- New functions: `replaceLineHeightForFamily`, `replaceLetterSpacingForFamily`, `selectTextNodesByFontVariant`, `segmentMatchesVariant`.
- README updated with v1.2.0 documentation.

---

## v1.1.0 — Released

**Commit:** `18a7c7b` — *major features updates* (Feb 16, 2026)

### New features

- **Scan Selected Frames** — Scan text only inside selected frames/groups/components/instances/sections.
- **Scope-aware operations** — Replace and Select by font respect the last scan scope (full page vs selection).
- Dynamic scan button: "Scan This Page" / "Scan Selected Frame" / "Scan N Frames" based on selection.
- `lastScanScope` in UI (`page` | `selection`) passed to all plugin operations.

### Technical changes

- `CONTAINER_TYPES`, `collectTextNodesFromSelection()`, `getTextNodesForScope(scope)`.
- All replace/select functions accept a `scope` parameter.
- `manifest.json`: plugin ID updated to `1603010877465738852`.
- README: scan selection feature documented.

### Release prep (Feb 2026)

| Commit | Date | Change |
|--------|------|--------|
| `eaa6035` | Feb 11 | Plugin UI size adjusted (`showUI`) |
| `fa2c9b9`, `3d42b58`, `de4432a` | Feb 10 | `manifest.json` updates (network access, plugin ID) |
| `33b9151`, `2d01bd1`, `379cfcb` | Feb 10 | UI adjustments |
| `ffd4861` | Feb 10 | Brand color update |

---

## v1.0.0 — Released

Initial development period: **Dec 1, 2025 – Dec 16, 2025** (plus UI iterations in Jan 2026 before v1.1).

### v1.0.0-core — First commit

**Commit:** `05266a5` — *first commit* (Dec 1, 2025, adamkatsu)

Base plugin structure:

| File | Description |
|------|-------------|
| `code.ts` | Page scan, replace font family, select by font |
| `ui.html` | Font list, missing font badge, replace drawer |
| `manifest.json` | Figma plugin config, `dynamic-page` |
| `package.json` | TypeScript + ESLint setup |

**Initial features:**
- Scan all text layers on the active page
- Missing font detection (fonts not installed on the system)
- Replace font family with style matching (Bold → Bold, fallback Regular)
- Mixed text support (`figma.mixed`) via `getStyledTextSegments`
- Google Fonts links to download missing fonts
- Custom dropdown font replace in UI

---

### v1.0.0 — UI & UX iterations (Dec 2025)

| Commit | Date | Author | Summary |
|--------|------|--------|---------|
| `9281c17` | Dec 2 | kemonn98 | UI design updates — initial layout & styling |
| `ecc894f` | Dec 2 | kemonn98 | **Font counter** — usage count per font family |
| `402f386` | Dec 2 | kemonn98 | Major UI design — expand/collapse font details |
| `e2e32f4` | Dec 2 | kemonn98 | Major UI design — font list interactions |
| `1f87527` | Dec 2 | kemonn98 | Interaction updates |
| `0c20bab` | Dec 2 | kemonn98 | UI style updates |
| `a53717e` | Dec 2 | kemonn98 | UI style updates (plugin dimensions) |
| `3961859` | Dec 2 | kemonn98 | UI style updates |
| `40a8914` | Dec 2 | kemonn98 | UI style updates |

---

### v1.0.0 — Font details & replace (Dec 2025)

| Commit | Date | Author | Summary |
|--------|------|--------|---------|
| `634479a` | Dec 4 | adamkatsu | **feat: scan + replace font weight and font size** — `FontUsageDetails` (styles + sizes), replace per weight/style, replace per fontSize, UI expand per font |
| `be63a9c` | Dec 4 | adamkatsu | Revert conflicting UI style commit |
| `3662411` | Dec 8 | adamkatsu | **feat: re-enable font replace on missing fonts** |
| `1fd8653` | Dec 9 | adamkatsu | **feat: fix font replace logic on missing fonts** — separate path for `hasMissingFont` |
| `ee21e1a` | Dec 16 | adamkatsu | **fix: font weight & size overflow bug** — UI overflow fix |

---

### v1.0.0 — Logic & UI stabilization (Jan 2026)

| Commit | Date | Author | Summary |
|--------|------|--------|---------|
| `7919a7c` | Jan 19 | kemonn98 | **Major UI and logic updates** — large refactor of `code.ts` + `ui.html`, notifications via postMessage, replacement progress |
| `80f2418` | Jan 19 | kemonn98 | **Fix missing font change** — deep fix for replace on missing/mixed layers |
| `7f5a8e3` | Jan 20 | kemonn98 | UX and UI update — improved interactions & appearance |
| `58f3b80` | Jan 21 | kemonn98 | UI updates — visual polish |

---

## Full Timeline (Chronological)

```
2025-12-01  05266a5  first commit                          [adamkatsu]
2025-12-02  9281c17  ui design updates                     [kemonn98]
2025-12-02  ecc894f  font counter                          [kemonn98]
2025-12-02  402f386  major ui design updates               [kemonn98]
2025-12-02  e2e32f4  major ui design updates               [kemonn98]
2025-12-02  1f87527  interaction updates                   [kemonn98]
2025-12-02  0c20bab  ui style updates                      [kemonn98]
2025-12-02  a53717e  ui style updates                      [kemonn98]
2025-12-02  3961859  ui style updates                      [kemonn98]
2025-12-02  40a8914  ui style updates                      [kemonn98]
2025-12-04  634479a  feat: scan + replace weight & size    [adamkatsu]
2025-12-04  be63a9c  revert ui style updates               [adamkatsu]
2025-12-08  3662411  feat: re-enable missing font replace  [adamkatsu]
2025-12-09  1fd8653  feat: fix missing font replace logic  [adamkatsu]
2025-12-16  ee21e1a  fix: weight & size overflow bug       [adamkatsu]
2026-01-19  7919a7c  major ui and logic updates            [kemonn98]
2026-01-19  80f2418  fix missing font change               [kemonn98]
2026-01-20  7f5a8e3  ux and ui update                      [kemonn98]
2026-01-21  58f3b80  ui updates                            [kemonn98]
2026-02-10  ffd4861  brand color update                    [kemonn98]
2026-02-10  de4432a  manifest json update                  [kemonn98]
2026-02-10  379cfcb  ui html update                        [kemonn98]
2026-02-10  2d01bd1  ui html update                        [kemonn98]
2026-02-10  33b9151  ui html update                        [kemonn98]
2026-02-10  3d42b58  manifest json update                  [kemonn98]
2026-02-10  fa2c9b9  manifest json update                  [kemonn98]
2026-02-11  eaa6035  showUI update                         [kemonn98]
2026-02-16  18a7c7b  major features updates    → v1.1.0   [kemonn98]
2026-02-24  4a8938a  v1.2.0 new features update            [kemonn98]
2026-07-10  ae0cfe4  release v1.2.1 — scan scope fix, docs  [kemonn98]
2026-07-10  41ac074  release v1.2.2 — release workflow fix  [kemonn98]
2026-07-10  b22bf26  release v1.2.3 — fix hash backfill workflow     [kemonn98]
2026-07-10  e782d9f  release v1.2.4 — ui.html version in release workflow    [kemonn98]
2026-07-10  release v1.2.5 — descriptive release commit messages         [kemonn98]
```

---

## Plugin Architecture (Current)

```
code.ts
├── Scanning
│   ├── getFontsFromPage / getFontsFromTextNodes
│   ├── collectTextNodesFromSelection (+ scannedContainers cache)
│   └── FontUsageDetails (styles, sizeVariants)
├── Replacement
│   ├── replaceFontFamily
│   ├── replaceFontStyleForFamily
│   ├── replaceFontSizeForFamily
│   ├── replaceLineHeightForFamily
│   └── replaceLetterSpacingForFamily
├── Selection
│   ├── selectTextNodesByFont
│   ├── selectTextNodesByFontAndStyle
│   └── selectTextNodesByFontVariant
└── Message handler (figma.ui.onmessage)

ui.html
├── performScan() → scan-layers | scan-selection
├── lastScanScope → scope for all operations
└── Font list UI (expand, replace, select, edit inline)
```

---

## How to Update This Document

After each release or significant change:

1. Add an entry to **Version Summary**
2. Write details in the relevant version section
3. Append a line to **Full Timeline** with the commit hash from `git log`
4. Sync version strings in `package.json`, `package-lock.json`, and `ui.html` (footer `.footer-version`)
5. Sync with [README.md](./README.md) → **Version history** and [AGENTS.md](./AGENTS.md) release checklist

On release (`git push` workflow): use `—` for the new release's commit hash in `VERSION.md`; backfill the previous release's hash with `git log -1 -S '"version": "X.Y.Z"' -- package.json`. Commit messages describe **what changed**, not `release: vX.Y.Z`.

```bash
# View recent commits
git log --oneline -10

# Inspect a specific commit
git show <commit-hash> --stat
```
