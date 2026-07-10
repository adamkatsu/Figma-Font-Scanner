# Font Scanner

A Figma plugin that scans your designs for font usage, surfaces missing fonts, and lets you replace or select text by font—for the whole page or just selected frames.

---

## About

**Font Scanner** helps you see which fonts are used in your file, spot fonts that aren’t installed (missing fonts), and quickly replace fonts or select layers by font. You can work on the entire current page or limit everything to selected frames (e.g. a single screen or component).

---

## Features

- **Scan fonts** – Discovers all font families used in text layers, with usage counts.
- **Font details** – For each font: styles (weights) and sizes used, with line height and letter spacing so you can see exactly how it’s applied and spot inconsistent styles.
- **Size list by line height & letter spacing** – Sizes are grouped by line height and letter spacing (e.g. 16px 120% 2%), shown in an aligned list. Click a value to edit, press Enter to apply; the list re-scans and merges identical variants.
- **Edit size, line height, and letter spacing** – Change font size, line height (% or Auto), or letter spacing (%) per row with the same click-to-edit flow.
- **Missing fonts** – Highlights fonts that aren’t available on your system and links to Google Fonts to download them.
- **Replace font** – Replace a font family with another (with best-effort style matching, e.g. Bold → Bold).
- **Replace style/size/line height/letter spacing** – Change weight (e.g. Regular → Medium), fontSize, line height, or letter spacing in bulk for a given font.
- **Select by font** – Select all text layers using a specific font (handy for cleanup or restyling).
- **Select by weight or size variant** – Click a weight row to select all layers with that font and style; click a size row (or its “Layers” cell) to select all layers with that exact size, line height, and letter spacing.
- **Page vs selection** – Scan the whole page or only selected frame(s); replace and select respect the same scope.

---

## How it works

1. **Open the plugin**  
   Run **Font Scanner** from the Figma menu (Plugins → Font Scanner). It runs an initial scan of the current page.

2. **Choose scope**  
   - **Scan This Page** – No selection (or nothing frame-like selected): scan all text on the current page.  
   - **Scan Selected Frame(s)** – Select one or more frames (or groups/components/sections): the main button becomes “Scan Selected Frame” or “Scan N Frames”. Click it to scan only text inside that selection.

3. **Review results**  
   The plugin lists every font family found, with counts. Expand a font to see **weights** (with counts and an Apply dropdown to change style) and a **size list** (Size · Line height · Letter spacing · Layers). Each size row shows one combination (e.g. 24px, 120%, 2%). Missing fonts are marked and can be opened on Google Fonts to install.

4. **Replace or select**  
   - **Weights:** Click a weight row (not the dropdown/Apply) to select all layers with that font and style. Use the dropdown and Apply to replace that weight with another style.  
   - **Sizes:** Click a size row (on the “Layers” cell or row background) to select all layers with that exact size, line height, and letter spacing. Click the size, line height, or letter spacing value to edit it; press Enter to apply.  
   - From a font’s main row you can **Replace** with another font or **Select** all layers using that font.  
   All of these actions use the same scope as the last scan (whole page or selected frames only).

5. **Re-scan**  
   After replacing fonts or changing the selection, run the scan again to refresh the list.

---

## Version history

- **1.2.4** (released)  
  - **UI version label** – Footer in `ui.html` now bumps with every release (`v1.2.4`).
  - **Release workflow** – `ui.html` added to mandatory files on `git push` alongside `package.json` and docs.

- **1.2.3** (released)  
  - **Release workflow fix** – Stop using amend for commit hashes in `VERSION.md`; backfill hashes on the next release instead. Guarantees a clean working tree after `git push`.

- **1.2.2** (released)  
  - **Release workflow fix** – `git push` now amends `VERSION.md` with the commit hash before push, so the working tree is clean with no follow-up commits.
  - Corrected v1.2.1 commit hash in `VERSION.md` (`ae0cfe4`).

- **1.2.1** (released)  
  - **Fix scan selection scope** – After "Scan Selected Frames", Select Fonts, replace, and other operations now keep using the scanned frames until you rescan—not the current canvas selection (which changes after selecting text layers).
  - **dynamic-page compatibility** – Uses direct node references instead of `getNodeById()` for scanned containers.
  - **Project docs** – Added `VERSION.md`, `AGENTS.md`, and Cursor release rules.

- **1.0.0** (released)  
  - Initial release: scan page, font list with counts and details, missing-font detection, replace font/style/size, select layers by font.

- **1.1.0** (released)  
  - **Scan selected frames** – When one or more frames (or groups/components/sections) are selected, you can scan only those. Replace and “Select by font” then apply only within that selection, so you can work on a single frame or a subset of the page.

- **1.2.0** (released)  
  - **Size list grouped by line height & letter spacing** – Sizes are shown in a list with columns (Size, Line height, Letter spacing, Layers) so you can spot inconsistent combinations (e.g. 24px at 120% vs 140%).
  - **Editable line height and letter spacing** – Click a line height or letter spacing value in the size list to change it (line height can be a number or “Auto”); press Enter to apply. Replaces only layers matching that exact variant.
  - **Select by weight** – Click a font weight row (the label/count area) to select all layers using that font and style. Weight rows have a hover state.
  - **Select by size variant** – Click a size row (the “Layers” cell or row background) to select all layers with that exact size, line height, and letter spacing.

---

## Development

This plugin uses TypeScript and npm.

For agent guidelines, architecture, and release workflows, see **[AGENTS.md](./AGENTS.md)**. Full changelog and commit timeline: **[VERSION.md](./VERSION.md)**.

### Prerequisites

- [Node.js](https://nodejs.org/en/download/) (includes npm)
- TypeScript: `npm install -g typescript`

### Setup

```bash
npm install --save-dev @figma/plugin-typings
```

### Build

- **One-off build:** `npm run build`  
- **Watch mode (rebuild on save):** `npm run watch`  
  In VS Code you can use **Terminal → Run Build Task…** and choose **npm: watch**.

TypeScript compiles `code.ts` to `code.js` for the plugin runtime.

---

## Links

- [Figma Plugin quickstart](https://www.figma.com/plugin-docs/plugin-quickstart-guide/)
- [TypeScript](https://www.typescriptlang.org/)
