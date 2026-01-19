// code.ts

figma.showUI(__html__, { width: 400, height: 600 });

// --- 1. SCANNING LOGIC ---

type FontUsageDetails = {
  styles: Map<string, number>;
  sizes: Map<number, number>;
};

async function getFontsFromPage() {
  const fontFamilies = new Map<string, number>(); // Map to count occurrences
  const missingFontFamilies = new Set<string>();
  const fontDetails = new Map<string, FontUsageDetails>();
  const familyStyleMap = new Map<string, Set<string>>();

  // Get all available system fonts for the dropdown
  const availableFontsList = await figma.listAvailableFontsAsync();
  
  // Create a unique list of family names for the UI dropdown
  const systemFonts = Array.from(new Set(availableFontsList.map(f => {
    const family = f.fontName.family;
    const style = f.fontName.style;
    if (!familyStyleMap.has(family)) {
      familyStyleMap.set(family, new Set<string>());
    }
    familyStyleMap.get(family)!.add(style);
    return family;
  }))).sort();

  // Create a lookup for fast checking
  const availableFamilies = new Set(systemFonts.map(f => f.toLowerCase()));

  // Find all Text Nodes
  const textNodes = figma.currentPage.findAllWithCriteria({ types: ['TEXT'] });

  const ensureFontDetail = (family: string): FontUsageDetails => {
    if (!fontDetails.has(family)) {
      fontDetails.set(family, {
        styles: new Map<string, number>(),
        sizes: new Map<number, number>()
      });
    }
    return fontDetails.get(family)!;
  };

  textNodes.forEach(node => {
    try {
      if (node.fontName === figma.mixed || node.fontSize === figma.mixed) {
        // For mixed text, count each segment separately
        const segments = node.getStyledTextSegments(['fontName', 'fontSize']);
        segments.forEach(segment => {
          const family = segment.fontName.family;
          fontFamilies.set(family, (fontFamilies.get(family) || 0) + 1);
          const detail = ensureFontDetail(family);
          const style = segment.fontName.style;
          if (style) {
            detail.styles.set(style, (detail.styles.get(style) || 0) + 1);
          }
          if (typeof segment.fontSize === 'number') {
            detail.sizes.set(segment.fontSize, (detail.sizes.get(segment.fontSize) || 0) + 1);
          }
          if (!availableFamilies.has(family.toLowerCase())) {
            missingFontFamilies.add(family);
          }
        });
      } else {
        // For single font text, count once
        const fontName = node.fontName as FontName;
        const family = fontName.family;
        fontFamilies.set(family, (fontFamilies.get(family) || 0) + 1);
        const detail = ensureFontDetail(family);
        const style = fontName.style;
        if (style) {
          detail.styles.set(style, (detail.styles.get(style) || 0) + 1);
        }
        if (typeof node.fontSize === 'number') {
          const sizeValue = node.fontSize as number;
          detail.sizes.set(sizeValue, (detail.sizes.get(sizeValue) || 0) + 1);
        }
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

  const fontDetailsPayload: Record<
    string,
    {
      styles: { value: string; count: number }[];
      sizes: { value: number; count: number }[];
    }
  > = {};

  fontDetails.forEach((detail, family) => {
    fontDetailsPayload[family] = {
      styles: Array.from(detail.styles.entries())
        .map(([value, count]) => ({ value, count }))
        .sort((a, b) => a.value.localeCompare(b.value)),
      sizes: Array.from(detail.sizes.entries())
        .map(([value, count]) => ({ value, count }))
        .sort((a, b) => a.value - b.value)
    };
  });

  const familyStylesPayload: Record<string, string[]> = {};
  familyStyleMap.forEach((styles, family) => {
    familyStylesPayload[family] = Array.from(styles).sort((a, b) => a.localeCompare(b));
  });

  figma.ui.postMessage({
    type: 'scan-result',
    fonts: fontsWithCount.map(f => f.name), // Keep for backward compatibility
    fontCounts: fontsWithCount, // New: includes counts
    missingFonts: Array.from(missingFontFamilies).sort(),
    systemFonts: systemFonts, // <--- Send list of installed fonts to UI
    fontDetails: fontDetailsPayload,
    familyStyles: familyStylesPayload
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
  
  // Get all available styles for the new font family FIRST
  const availableFonts = await figma.listAvailableFontsAsync();
  const newFontStyles = availableFonts
    .filter(f => f.fontName.family.toLowerCase() === newFamily.toLowerCase())
    .map(f => f.fontName);
  
  if (newFontStyles.length === 0) {
    figma.ui.postMessage({
      type: 'notification',
      message: `Font "${newFamily}" is not available on this system.`,
      count: 0
    });
    return;
  }

  // Create a helper to find best matching style
  const findBestStyle = (oldStyle: string): FontName => {
    // Try exact match first
    let match = newFontStyles.find(f => 
      f.style.toLowerCase() === oldStyle.toLowerCase()
    );
    if (match) return match;
    
    // Try "Regular"
    match = newFontStyles.find(f => 
      f.style.toLowerCase() === 'regular'
    );
    if (match) return match;
    
    // Return first available style
    return newFontStyles[0];
  };
  
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
  const totalNodes = nodesToUpdate.length;
  
  // Send initial progress
  figma.ui.postMessage({
    type: 'replacement-progress',
    current: 0,
    total: totalNodes,
    fontName: newFamily
  });
  
  for (let i = 0; i < nodesToUpdate.length; i++) {
    const node = nodesToUpdate[i];
    try {
      if (node.fontName === figma.mixed) {
        // Handle Mixed Text - CRITICAL: Load ALL existing fonts first!
        const segments = node.getStyledTextSegments(['fontName']);
        
        // STEP 1: Load ALL existing fonts in this text node (even if we won't change them)
        // This is required by Figma before we can modify any segment
        const allFontsInNode = new Set<string>();
        for (const segment of segments) {
          const fontKey = `${segment.fontName.family}::${segment.fontName.style}`;
          allFontsInNode.add(fontKey);
        }
        
        for (const fontKey of allFontsInNode) {
          const [family, style] = fontKey.split('::');
          try {
            await figma.loadFontAsync({ family, style });
          } catch (e) {
            console.log(`Cannot load existing font ${family} ${style}`);
          }
        }
        
        // STEP 2: Now we can safely modify segments
        let hasChanges = false;
        for (const segment of segments) {
          if (segment.fontName.family.toLowerCase() === targetOld) {
            const newFont = findBestStyle(segment.fontName.style);
            
            try {
              await figma.loadFontAsync(newFont);
              node.setRangeFontName(segment.start, segment.end, newFont);
              hasChanges = true;
            } catch (e) {
              console.error(`Failed to apply ${newFont.family} ${newFont.style}`, e);
            }
          }
        }
        
        if (hasChanges) updateCount++;
      } else {
        // Handle Single Text
        const currentFont = node.fontName as FontName;
        if (currentFont.family.toLowerCase() === targetOld) {
          const newFont = findBestStyle(currentFont.style);
          
          try {
            await figma.loadFontAsync(newFont);
            node.fontName = newFont;
            updateCount++;
          } catch (e) {
            console.error(`Failed to apply ${newFont.family} ${newFont.style}`, e);
          }
        }
      }
      
      // Send progress update
      figma.ui.postMessage({
        type: 'replacement-progress',
        current: i + 1,
        total: totalNodes,
        fontName: newFamily
      });
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
async function loadFontWithCache(font: FontName, cache: Set<string>) {
  const key = `${font.family}__${font.style}`;
  if (cache.has(key)) {
    return;
  }
  await figma.loadFontAsync(font);
  cache.add(key);
}

async function replaceFontStyleForFamily(family: string, oldStyle: string, newStyle: string) {
  const targetFamily = family.trim().toLowerCase();
  const targetStyle = oldStyle.trim().toLowerCase();
  const replacementStyle = newStyle.trim();

  if (!targetFamily || !targetStyle || !replacementStyle) {
    return;
  }

  const textNodes = figma.currentPage.findAllWithCriteria({ types: ['TEXT'] });
  const fontLoadCache = new Set<string>();
  let updateCount = 0;

  for (const node of textNodes) {
    let nodeUpdated = false;
    try {
      if (node.fontName === figma.mixed) {
        const segments = node.getStyledTextSegments(['fontName']);
        
        // CRITICAL: Load ALL existing fonts first before modifying any segment
        const allFontsInNode = new Set<string>();
        for (const segment of segments) {
          const fontKey = `${segment.fontName.family}::${segment.fontName.style}`;
          allFontsInNode.add(fontKey);
        }
        
        for (const fontKey of allFontsInNode) {
          const [family, style] = fontKey.split('::');
          try {
            await figma.loadFontAsync({ family, style });
          } catch (e) {
            // Font might be missing, that's ok
          }
        }
        
        // Now we can safely modify segments
        for (const segment of segments) {
          const segFont = segment.fontName;
          if (
            segFont.family.toLowerCase() === targetFamily &&
            segFont.style.toLowerCase() === targetStyle
          ) {
            const newFont: FontName = {
              family: segFont.family,
              style: replacementStyle
            };
            try {
              await loadFontWithCache(newFont, fontLoadCache);
              node.setRangeFontName(segment.start, segment.end, newFont);
              nodeUpdated = true;
            } catch (err) {
              console.error(`Failed to load "${replacementStyle}" for ${segFont.family}`, err);
            }
          }
        }
      } else {
        const currentFont = node.fontName as FontName;
        if (
          currentFont.family.toLowerCase() === targetFamily &&
          currentFont.style.toLowerCase() === targetStyle
        ) {
          const newFont: FontName = {
            family: currentFont.family,
            style: replacementStyle
          };
          try {
            await loadFontWithCache(newFont, fontLoadCache);
            node.fontName = newFont;
            nodeUpdated = true;
          } catch (err) {
            console.error(`Failed to load "${replacementStyle}" for ${currentFont.family}`, err);
          }
        }
      }
    } catch (error) {
      console.error('Failed to process node for style replacement', error);
    }

    if (nodeUpdated) {
      updateCount++;
    }
  }

  figma.ui.postMessage({
    type: 'notification',
    message:
      updateCount === 0
        ? `No "${family}" layers with "${oldStyle}" weight found.`
        : `Updated ${family} ${oldStyle} → ${replacementStyle} in ${updateCount} layer${updateCount === 1 ? '' : 's'}.`,
    count: updateCount,
    fontName: family
  });

  if (updateCount > 0) {
    await getFontsFromPage();
  }
}

async function replaceFontSizeForFamily(family: string, oldSize: number, newSize: number) {
  const targetFamily = family.trim().toLowerCase();
  if (!targetFamily || Number.isNaN(newSize) || newSize <= 0) {
    return;
  }

  const textNodes = figma.currentPage.findAllWithCriteria({ types: ['TEXT'] });
  const fontLoadCache = new Set<string>();
  let updateCount = 0;

  for (const node of textNodes) {
    let nodeUpdated = false;
    try {
      if (node.fontSize === figma.mixed || node.fontName === figma.mixed) {
        const segments = node.getStyledTextSegments(['fontName', 'fontSize']);
        
        // CRITICAL: Load ALL existing fonts first before modifying any segment
        const allFontsInNode = new Set<string>();
        for (const segment of segments) {
          const fontKey = `${segment.fontName.family}::${segment.fontName.style}`;
          allFontsInNode.add(fontKey);
        }
        
        for (const fontKey of allFontsInNode) {
          const [family, style] = fontKey.split('::');
          try {
            await figma.loadFontAsync({ family, style });
          } catch (e) {
            // Font might be missing, that's ok
          }
        }
        
        // Now we can safely modify segments
        for (const segment of segments) {
          const segFont = segment.fontName;
          if (
            segFont.family.toLowerCase() === targetFamily &&
            segment.fontSize === oldSize
          ) {
            try {
              await loadFontWithCache(segFont, fontLoadCache);
              node.setRangeFontSize(segment.start, segment.end, newSize);
              nodeUpdated = true;
            } catch (err) {
              console.error('Failed to update font size for segment', err);
            }
          }
        }
      } else {
        const currentFont = node.fontName as FontName;
        if (
          currentFont.family.toLowerCase() === targetFamily &&
          (node.fontSize as number) === oldSize
        ) {
          try {
            await loadFontWithCache(currentFont, fontLoadCache);
            node.fontSize = newSize;
            nodeUpdated = true;
          } catch (err) {
            console.error('Failed to update font size for node', err);
          }
        }
      }
    } catch (error) {
      console.error('Failed to process node for size replacement', error);
    }

    if (nodeUpdated) {
      updateCount++;
    }
  }

  figma.ui.postMessage({
    type: 'notification',
    message:
      updateCount === 0
        ? `No ${family} layers found with ${oldSize}px text.`
        : `Updated ${family} text ${oldSize}px → ${newSize}px in ${updateCount} layer${updateCount === 1 ? '' : 's'}.`,
    count: updateCount,
    fontName: family
  });

  if (updateCount > 0) {
    await getFontsFromPage();
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

// --- 4. SELECTION CHANGE TRACKING ---

let lastSelectedFont: string | null = null;

// Listen for selection changes in Figma
figma.on('selectionchange', () => {
  const selection = figma.currentPage.selection;
  
  // If nothing is selected, clear the UI selection
  if (selection.length === 0) {
    if (lastSelectedFont) {
      figma.ui.postMessage({
        type: 'deselect-font'
      });
      lastSelectedFont = null;
    }
    return;
  }
  
  // Check if any selected nodes match the last selected font
  if (lastSelectedFont) {
    const targetFamily = lastSelectedFont.toLowerCase();
    let hasMatchingFont = false;
    
    for (const node of selection) {
      if (node.type === 'TEXT') {
        try {
          const nodeFonts = extractFontFamilies(node as TextNode);
          if (Array.from(nodeFonts).some(f => f.toLowerCase() === targetFamily)) {
            hasMatchingFont = true;
            break;
          }
        } catch (e) {
          // Ignore nodes we can't read
        }
      }
    }
    
    // If no matching fonts in selection, deselect in UI
    if (!hasMatchingFont) {
      figma.ui.postMessage({
        type: 'deselect-font'
      });
      lastSelectedFont = null;
    }
  }
});

// --- MESSAGE HANDLER ---

figma.ui.onmessage = async msg => {
  if (msg.type === 'scan-layers') {
    await getFontsFromPage();
  } else if (msg.type === 'select-font') {
    lastSelectedFont = msg.font;
    await selectTextNodesByFont(msg.font);
  } else if (msg.type === 'replace-font') {
    await replaceFontFamily(msg.oldFont, msg.newFont);
  } else if (msg.type === 'replace-font-weight') {
    await replaceFontStyleForFamily(msg.family, msg.oldStyle, msg.newStyle);
  } else if (msg.type === 'replace-font-size') {
    await replaceFontSizeForFamily(msg.family, msg.oldSize, msg.newSize);
  }
};