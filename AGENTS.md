# AGENTS.md — Font Scanner

Guidelines for AI agents (Cursor, Copilot, etc.) working on this repository. This file is the **single entry point** for project conventions. It links to all other docs and Cursor rules.

---

## Document map

| Document | Purpose | Audience |
|----------|---------|----------|
| **[README.md](./README.md)** | User-facing overview, features, how to use, short version history | Users & contributors |
| **[VERSION.md](./VERSION.md)** | Full changelog, commit timeline, technical release notes | Maintainers & agents |
| **AGENTS.md** (this file) | Agent rules, architecture, workflows | AI agents |
| **[.cursor/rules/](./cursor/rules/)** | Enforced Cursor rules (auto-loaded in Cursor) | Cursor agent |

**Keep in sync:** When shipping a release, update **all** of these together:

| File | What to update |
|------|----------------|
| `package.json` | `"version"` |
| `package-lock.json` | root `"version"` + `packages[""].version` |
| `ui.html` | `.footer-version` → `vX.Y.Z` |
| `README.md` | Version history entry |
| `VERSION.md` | Version Summary, release section, Full Timeline |

See [Release workflow](#release-workflow).

---

## Project overview

**Font Scanner** is a Figma plugin that scans text layers for font usage, detects missing fonts, and supports bulk replace/select operations scoped to the full page or selected frames.

| | |
|---|---|
| Repository | [github.com/adamkatsu/Figma-Font-Scanner](https://github.com/adamkatsu/Figma-Font-Scanner) |
| Current version | See `package.json` and [VERSION.md → Version Summary](./VERSION.md#version-summary) |
| Plugin ID | `1603010877465738852` |
| Stack | TypeScript → `code.js`, HTML UI, Figma Plugin API |
| Document access | `dynamic-page` (see [Figma API constraints](#figma-api-constraints)) |

---

## Repository structure

```
Figma-Font-Scanner/
├── code.ts          # Plugin main thread (scan, replace, select, message handler)
├── code.js          # Compiled output (do not edit by hand)
├── ui.html          # Plugin UI (scan button, font list, inline edit, modals)
├── manifest.json    # Figma plugin manifest
├── package.json     # npm version & build scripts
├── tsconfig.json
├── README.md        # User documentation
├── VERSION.md       # Detailed version history & timeline
├── AGENTS.md        # This file
└── .cursor/rules/
    ├── english-only.mdc
    └── git-push-release.mdc
```

### `code.ts` modules

| Section | Responsibility |
|---------|----------------|
| **1. Scanning** | `getFontsFromPage`, `getFontsFromTextNodes`, `collectTextNodesFromSelection`, `scannedContainers` cache |
| **2. Replacement** | `replaceFontFamily`, `replaceFontStyleForFamily`, `replaceFontSizeForFamily`, `replaceLineHeightForFamily`, `replaceLetterSpacingForFamily` |
| **3. Selection** | `selectTextNodesByFont`, `selectTextNodesByFontAndStyle`, `selectTextNodesByFontVariant` |
| **4. Selection tracking** | `figma.on('selectionchange')`, sync UI highlight with canvas selection |

### `ui.html` key concepts

| Concept | Description |
|---------|-------------|
| `lastScanScope` | `'page'` or `'selection'` — passed with every plugin message after scan |
| `performScan()` | Sends `scan-layers` or `scan-selection` based on frame selection count |
| `selectedFrameCount` | Drives dynamic scan button label |

### UI ↔ main thread messages

| UI → `code.ts` | Purpose |
|----------------|---------|
| `scan-layers` | Scan entire current page |
| `scan-selection` | Scan text inside selected containers; stores `scannedContainers` |
| `select-font` | Select layers by font family |
| `select-font-weight` | Select by family + style |
| `select-font-variant` | Select by size + line height + letter spacing |
| `replace-font` | Replace font family |
| `replace-font-weight` | Replace style within family |
| `replace-font-size` | Replace font size (optional variant filter) |
| `replace-font-line-height` | Replace line height for a variant |
| `replace-font-letter-spacing` | Replace letter spacing for a variant |

All replace/select messages accept `scope: 'page' | 'selection'`.

---

## Cursor rules

Rules in `.cursor/rules/` are loaded automatically by Cursor (`alwaysApply: true`).

| Rule file | Trigger | Summary |
|-----------|---------|---------|
| [english-only.mdc](./.cursor/rules/english-only.mdc) | Always | All docs, code comments, UI strings, and commits in **English** |
| [git-push-release.mdc](./.cursor/rules/git-push-release.mdc) | User says `git push` | Full release: bump version → update docs → commit → push |

Agents must read and follow these rules. Do not duplicate them here — refer to the rule files for exact steps.

---

## Agent guidelines

### Language

- **English only** for documentation, code comments, user-facing plugin text, commit messages, and changelogs.
- Rule: [.cursor/rules/english-only.mdc](./.cursor/rules/english-only.mdc)

### Code style

- **Minimize scope** — smallest correct diff; no unrelated changes.
- **Match existing patterns** — naming, message types, UI structure in `ui.html`.
- **No over-engineering** — avoid abstractions for one-off logic.
- **Comments** — only for non-obvious behavior (e.g. `hasMissingFont` paths, `figma.mixed` handling).
- **Build after `code.ts` changes:** `npm run build` (or `npm run watch`).

### Figma API constraints

This plugin uses `documentAccess: dynamic-page` in `manifest.json`:

- **Do not** use synchronous `figma.getNodeById()` — use `figma.getNodeByIdAsync()` or keep direct `SceneNode` references.
- **Scan selection scope:** `scannedContainers` holds frame references from the last `scan-selection`; reuse until next scan. Do not rely on `figma.currentPage.selection` after "Select Fonts" changes selection to text layers.
- **Mixed text:** Always load existing fonts before `setRangeFontName` / `setRangeFontSize` on mixed nodes.
- **Missing fonts:** Separate code path when `node.hasMissingFont` is true.

### Container types (selection scope)

```ts
CONTAINER_TYPES = ['FRAME', 'GROUP', 'COMPONENT', 'INSTANCE', 'SECTION']
```

### What not to commit

- `.DS_Store`
- Secrets (API keys in `ui.html` — review before public release)
- Hand-edited `code.js` (always compile from `code.ts`)

### Git safety

- Never `git push --force` to `main` unless explicitly requested.
- Never update `git config`.
- Only commit when the user asks (or as part of the `git push` release workflow).
- Never `--amend` unless hooks auto-fixed files from your own unpushed commit.

---

## Release workflow

Triggered when the user writes:

| Command | Version bump | Example |
|---------|----------------|---------|
| `git push` | Patch `x.y.Z+1` | `1.2.0` → `1.2.1` |
| `git push minor` | Minor `x.Y+1.0` | `1.2.3` → `1.3.0` |
| `git push major` | Major `X+1.0.0` | `1.2.3` → `2.0.0` |

Full steps: [.cursor/rules/git-push-release.mdc](./.cursor/rules/git-push-release.mdc)

### Checklist (summary)

1. Read [VERSION.md → Version Summary](./VERSION.md#version-summary) and `package.json` for current version.
2. Analyze `git diff` + conversation for changelog.
3. **Bump version in `package.json`** (required).
4. **Bump version in `package-lock.json`** — root and `packages[""]` (required).
5. **Bump version in `ui.html`** — `.footer-version` span: `vX.Y.Z` (required).
6. Add entry to [README.md → Version history](./README.md#version-history).
7. Update [VERSION.md](./VERSION.md): Version Summary (`—` for new release commit), version section, Full Timeline; backfill previous release hash from `git log`.
8. Run `npm run build` if `code.ts` changed.
9. `git commit` with `release: vX.Y.Z` message.
10. `git push -u origin HEAD` — **final step**.
11. Verify `git status` is clean — no uncommitted changes.

**One push, clean tree.** Do not amend for commit hashes. Backfill the previous version's hash when cutting the next release.

### Mandatory files on every release

| File | Required |
|------|----------|
| `package.json` | Yes |
| `package-lock.json` | Yes |
| `ui.html` | Yes — footer `.footer-version` |
| `README.md` | Yes |
| `VERSION.md` | Yes |
| `code.js` | Rebuild locally if `code.ts` changed (gitignored — not committed) |
| `AGENTS.md` | Only if modified |
| `code.ts`, `manifest.json` | If changed in this release |

---

## Versioning policy

| Bump | When to use |
|------|-------------|
| **Patch** | Bug fixes, small improvements, mini features |
| **Minor** | New feature set, non-breaking enhancements |
| **Major** | Breaking changes or large milestones |

- **README.md** — short bullet list per release (user-facing).
- **VERSION.md** — detailed notes, commit hashes, technical changes, timeline.
- **ui.html** — footer `.footer-version` label shown in the plugin UI (**must** match `package.json` on every release).
- **package.json** — canonical semver string for npm/tooling (**must** be bumped on every release).
- **package-lock.json** — must match `package.json` version (root + `packages[""]`).

---

## Development commands

```bash
npm install          # Install dependencies
npm run build        # Compile code.ts → code.js
npm run watch        # Rebuild on save
npm run lint         # ESLint
npm run lint:fix     # ESLint with auto-fix
```

Reload the plugin in Figma after rebuilding to test changes.

---

## Common tasks for agents

### Add a new replace/select operation

1. Implement logic in `code.ts` with `scope: 'page' | 'selection'`.
2. Use `getTextNodesForScope(scope)` — respects `scannedContainers` for selection scope.
3. Add message handler in `figma.ui.onmessage`.
4. Wire UI in `ui.html` with `lastScanScope`.
5. Post `notification` or `scan-result` messages to refresh UI.
6. Document in README features if user-facing.

### Fix selection-scope bugs

- Verify `scannedContainers` is set on `scan-selection` and cleared on `scan-layers`.
- Verify selection-scoped ops use `collectTextNodesFromSelection()`, not raw `figma.currentPage.selection`.
- Test: scan frames → select font A → select font B → replace font — all should target scanned frames.

### Update documentation after a feature

1. **README.md** — feature bullet under Features / How it works if user-visible.
2. **VERSION.md** — add to in-progress section or wait for release.
3. Do not update version numbers until release (`git push` workflow).

---

## Links

- [Figma Plugin API](https://www.figma.com/plugin-docs/)
- [Plugin quickstart](https://www.figma.com/plugin-docs/plugin-quickstart-guide/)
- [dynamic-page document access](https://www.figma.com/plugin-docs/accessing-document/)
- [Repository on GitHub](https://github.com/adamkatsu/Figma-Font-Scanner)
