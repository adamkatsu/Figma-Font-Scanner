// code.ts

figma.showUI(__html__, { width: 420, height: 520 });

// --- 1. SCANNING LOGIC ---

async function getFontsFromPage() {
  const fontFamilies = new Map<string, number>(); // Map to count occurrences
  const missingFontFamilies = new Set<string>();

  // Get all available system fonts for the dropdown
  const availableFontsList = await figma.listAvailableFontsAsync();
  
  // Create a unique list of family names for the UI dropdown
  const systemFonts = Array.from(new Set(availableFontsList.map(f => f.fontName.family))).sort();

  // Create a lookup for fast checking
  const availableFamilies = new Set(systemFonts.map(f => f.toLowerCase()));

  // Find all Text Nodes
  const textNodes = figma.currentPage.findAllWithCriteria({ types: ['TEXT'] });

  textNodes.forEach(node => {
    try {
      if (node.fontName === figma.mixed) {
        // For mixed text, count each segment separately
        const segments = node.getStyledTextSegments(['fontName']);
        segments.forEach(segment => {
          const family = segment.fontName.family;
          fontFamilies.set(family, (fontFamilies.get(family) || 0) + 1);
          if (!availableFamilies.has(family.toLowerCase())) {
            missingFontFamilies.add(family);
          }
        });
      } else {
        // For single font text, count once
        const fontName = node.fontName as FontName;
        const family = fontName.family;
        fontFamilies.set(family, (fontFamilies.get(family) || 0) + 1);
        if (!availableFamilies.has(family.toLowerCase())) {
          missingFontFamilies.add(family);
        }
      }
    } catch (error) {
      // Ignore nodes we can't read
    }
  });

  // Convert Map to array of objects with name and count
  const fontsWithCount = Array.from(fontFamilies.entries())
    .map(([name, count]) => ({ name, count }))
    .sort((a, b) => a.name.localeCompare(b.name));

  figma.ui.postMessage({
    type: 'scan-result',
    fonts: fontsWithCount.map(f => f.name), // Keep for backward compatibility
    fontCounts: fontsWithCount, // New: includes counts
    missingFonts: Array.from(missingFontFamilies).sort(),
    systemFonts: systemFonts // <--- Send list of installed fonts to UI
  });
}

function extractFontFamilies(node: TextNode): Set<string> {
  const nodeFonts = new Set<string>();
  try {
    if (node.fontName === figma.mixed) {
      const segments = node.getStyledTextSegments(['fontName']);
      segments.forEach(segment => nodeFonts.add(segment.fontName.family));
    } else {
      const fontName = node.fontName as FontName;
      nodeFonts.add(fontName.family);
    }
  } catch (error) {
    // Ignore nodes we can't read
  }
  return nodeFonts;
}

// --- 2. REPLACEMENT LOGIC ---

async function replaceFontFamily(oldFamily: string, newFamily: string) {
  const targetOld = oldFamily.toLowerCase();
  
  // Find all nodes using the old font
  const textNodes = figma.currentPage.findAllWithCriteria({ types: ['TEXT'] });
  const nodesToUpdate: TextNode[] = [];

  // Filter nodes first
  for (const node of textNodes) {
    const families = extractFontFamilies(node);
    if (Array.from(families).some(f => f.toLowerCase() === targetOld)) {
      nodesToUpdate.push(node);
    }
  }

  if (nodesToUpdate.length === 0) {
    figma.ui.postMessage({
      type: 'notification',
      message: `No layers found using "${oldFamily}".`,
      count: 0
    });
    return;
  }

  let updateCount = 0;

  // We need to load the replacement font.
  // Challenge: We don't know if "Bold" exists in the new font.
  // We will try to preserve style, otherwise fall back to "Regular".
  
  for (const node of nodesToUpdate) {
    try {
      if (node.fontName === figma.mixed) {
        // Handle Mixed Text
        const segments = node.getStyledTextSegments(['fontName']);
        for (const segment of segments) {
          if (segment.fontName.family.toLowerCase() === targetOld) {
            await applyFontToRange(node, segment.start, segment.end, segment.fontName.style, newFamily);
          }
        }
      } else {
        // Handle Single Text
        const currentStyle = (node.fontName as FontName).style;
        // We have to load the OLD font first to edit the node, 
        // but often it's missing. Figma allows setting a new font 
        // even if the old one is missing, as long as we load the NEW one.
        await applyFontToNode(node, currentStyle, newFamily);
      }
      updateCount++;
    } catch (err) {
      console.error("Failed to replace font on node", err);
    }
  }

  figma.ui.postMessage({
    type: 'notification',
    message: `Replaced "${oldFamily}" with "${newFamily}" in ${updateCount} layer${updateCount === 1 ? '' : 's'}.`,
    count: updateCount,
    fontName: newFamily
  });
  // Re-scan to update UI
  await getFontsFromPage();
}

// Helper: Try to match style (Bold -> Bold), fallback to Regular if needed
async function applyFontToNode(node: TextNode, style: string, newFamily: string) {
  let fontToLoad: FontName = { family: newFamily, style: style };
  
  // Try to load the exact style (e.g. Bold)
  try {
    await figma.loadFontAsync(fontToLoad);
    node.fontName = fontToLoad;
  } catch (e) {
    // If that specific weight doesn't exist, try "Regular"
    try {
      fontToLoad = { family: newFamily, style: "Regular" };
      await figma.loadFontAsync(fontToLoad);
      node.fontName = fontToLoad;
    } catch (e2) {
      console.error(`Could not load Regular style for ${newFamily}`);
    }
  }
}

// Helper: Same as above but for ranges (mixed text)
async function applyFontToRange(node: TextNode, start: number, end: number, style: string, newFamily: string) {
  let fontToLoad: FontName = { family: newFamily, style: style };
  try {
    await figma.loadFontAsync(fontToLoad);
    node.setRangeFontName(start, end, fontToLoad);
  } catch (e) {
    try {
      fontToLoad = { family: newFamily, style: "Regular" };
      await figma.loadFontAsync(fontToLoad);
      node.setRangeFontName(start, end, fontToLoad);
    } catch (e2) {}
  }
}

// --- 3. SELECTION LOGIC ---

async function selectTextNodesByFont(fontFamily: string) {
  const targetFamily = fontFamily.toLowerCase();
  const matches: SceneNode[] = [];
  const textNodes = figma.currentPage.findAllWithCriteria({ types: ['TEXT'] });

  textNodes.forEach(node => {
    const nodeFonts = extractFontFamilies(node);
    if (Array.from(nodeFonts).some(f => f.toLowerCase() === targetFamily)) {
      matches.push(node);
    }
  });

  figma.currentPage.selection = matches;
  if (matches.length > 0) {
    figma.viewport.scrollAndZoomIntoView(matches);
    figma.ui.postMessage({
      type: 'notification',
      message: `All "${fontFamily}" font${matches.length === 1 ? '' : 's'} selected!`,
      count: matches.length,
      fontName: fontFamily
    });
  } else {
    figma.ui.postMessage({
      type: 'notification',
      message: 'No layers found on this page.',
      count: 0
    });
  }
}

// --- MESSAGE HANDLER ---

figma.ui.onmessage = async msg => {
  if (msg.type === 'scan-layers') {
    await getFontsFromPage();
  } else if (msg.type === 'select-font') {
    await selectTextNodesByFont(msg.font);
  } else if (msg.type === 'replace-font') {
    await replaceFontFamily(msg.oldFont, msg.newFont);
  }
};