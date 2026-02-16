# Font Scanner

A Figma plugin that scans your designs for font usage, surfaces missing fonts, and lets you replace or select text by font—for the whole page or just selected frames.

---

## About

**Font Scanner** helps you see which fonts are used in your file, spot fonts that aren’t installed (missing fonts), and quickly replace fonts or select layers by font. You can work on the entire current page or limit everything to selected frames (e.g. a single screen or component).

---

## Features

- **Scan fonts** – Discovers all font families used in text layers, with usage counts.
- **Font details** – For each font: styles (weights) and sizes used, so you can see exactly how it’s applied.
- **Missing fonts** – Highlights fonts that aren’t available on your system and links to Google Fonts to download them.
- **Replace font** – Replace a font family with another (with best-effort style matching, e.g. Bold → Bold).
- **Replace style/size** – Change weight (e.g. Regular → Medium) or fontSize in bulk for a given font.
- **Select by font** – Select all text layers using a specific font (handy for cleanup or restyling).
- **Page vs selection** – Scan the whole page or only selected frame(s); replace and select respect the same scope.

---

## How it works

1. **Open the plugin**  
   Run **Font Scanner** from the Figma menu (Plugins → Font Scanner). It runs an initial scan of the current page.

2. **Choose scope**  
   - **Scan This Page** – No selection (or nothing frame-like selected): scan all text on the current page.  
   - **Scan Selected Frame(s)** – Select one or more frames (or groups/components/sections): the main button becomes “Scan Selected Frame” or “Scan N Frames”. Click it to scan only text inside that selection.

3. **Review results**  
   The plugin lists every font family found, with counts. Expand a font to see styles and sizes. Missing fonts are marked and can be opened on Google Fonts to install.

4. **Replace or select**  
   From a font’s row you can:  
   - **Replace** with another font (or change style/size).  
   - **Select** all layers using that font in Figma.  
   All of these actions use the same scope as the last scan (whole page or selected frames only).

5. **Re-scan**  
   After replacing fonts or changing the selection, run the scan again to refresh the list.

---

## Version history

- **1.0.0** (released)  
  - Initial release: scan page, font list with counts and details, missing-font detection, replace font/style/size, select layers by font.

- **1.1.0** (upcoming)  
  - **Scan selected frames** – When one or more frames (or groups/components/sections) are selected, you can scan only those. Replace and “Select by font” then apply only within that selection, so you can work on a single frame or a subset of the page.

---

## Development

This plugin uses TypeScript and npm.

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
