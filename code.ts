// code.ts

figma.showUI(__html__, { width: 360, height: 600 });

// --- 1. SCANNING LOGIC ---

type FontUsageDetails = {
  styles: Map<string, number>;
  /** Key: "fontSize_lineHeightPercent_letterSpacingPercent" (lineHeight "AUTO" = "auto") */
  sizeVariants: Map<string, { fontSize: number; lineHeightPercent: number | null; letterSpacingPercent: number; count: number }>;
};

const CONTAINER_TYPES = new Set<string>(['FRAME', 'GROUP', 'COMPONENT', 'INSTANCE', 'SECTION']);

function normalizeLineHeight(lineHeight: LineHeight | undefined, fontSize: number): number | null {
  if (!lineHeight || lineHeight.unit === 'AUTO') return null;
  if (lineHeight.unit === 'PERCENT') return Math.round((lineHeight as { value: number }).value * 10) / 10;
  const px = (lineHeight as { value: number }).value;
  return Math.round((px / fontSize) * 1000) / 10;
}

function normalizeLetterSpacing(letterSpacing: LetterSpacing | undefined, fontSize: number): number {
  if (!letterSpacing) return 0;
  if (letterSpacing.unit === 'PERCENT') return Math.round(letterSpacing.value * 10) / 10;
  const px = letterSpacing.value;
  return Math.round((px / fontSize) * 1000) / 10;
}

function sizeVariantKey(fontSize: number, lineHeightPercent: number | null, letterSpacingPercent: number): string {
  const lh = lineHeightPercent === null ? 'auto' : String(lineHeightPercent);
  return `${fontSize}_${lh}_${letterSpacingPercent}`;
}

function getSelectedContainerCount(): number {
  const selection = figma.currentPage.selection;
  return selection.filter(node => CONTAINER_TYPES.has(node.type)).length;
}

function collectTextNodesFromSelection(): TextNode[] {
  const selection = figma.currentPage.selection;
  const containers = selection.filter(node => CONTAINER_TYPES.has(node.type));
  const textNodes: TextNode[] = [];
  const seen = new Set<string>();
  for (const node of containers) {
    const container = node as { findAllWithCriteria(criteria: { types: string[] }): TextNode[] };
    const children = container.findAllWithCriteria({ types: ['TEXT'] });
    for (const text of children) {
      if (!seen.has(text.id)) {
        seen.add(text.id);
        textNodes.push(text);
      }
    }
  }
  return textNodes;
}

function getTextNodesForScope(scope: 'page' | 'selection'): TextNode[] {
  if (scope === 'selection') {
    return collectTextNodesFromSelection();
  }
  return figma.currentPage.findAllWithCriteria({ types: ['TEXT'] });
}

async function getFontsFromTextNodes(textNodes: TextNode[]) {
  const fontFamilies = new Map<string, number>();
  const missingFontFamilies = new Set<string>();
  const fontDetails = new Map<string, FontUsageDetails>();
  const familyStyleMap = new Map<string, Set<string>>();

  const availableFontsList = await figma.listAvailableFontsAsync();

  const systemFonts = Array.from(new Set(availableFontsList.map(f => {
    const family = f.fontName.family;
    const style = f.fontName.style;
    if (!familyStyleMap.has(family)) {
      familyStyleMap.set(family, new Set<string>());
    }
    familyStyleMap.get(family)!.add(style);
    return family;
  }))).sort();

  const availableFamilies = new Set(systemFonts.map(f => f.toLowerCase()));

  const ensureFontDetail = (family: string): FontUsageDetails => {
    if (!fontDetails.has(family)) {
      fontDetails.set(family, {
        styles: new Map<string, number>(),
        sizeVariants: new Map()
      });
    }
    return fontDetails.get(family)!;
  };

  function addSizeVariant(detail: FontUsageDetails, fontSize: number, lineHeight: LineHeight | undefined, letterSpacing: LetterSpacing | undefined) {
    const lh = normalizeLineHeight(lineHeight, fontSize);
    const ls = normalizeLetterSpacing(letterSpacing, fontSize);
    const key = sizeVariantKey(fontSize, lh, ls);
    const existing = detail.sizeVariants.get(key);
    if (existing) {
      existing.count += 1;
    } else {
      detail.sizeVariants.set(key, { fontSize, lineHeightPercent: lh, letterSpacingPercent: ls, count: 1 });
    }
  }

  textNodes.forEach(node => {
    try {
      if (node.fontName === figma.mixed || node.fontSize === figma.mixed) {
        const segments = node.getStyledTextSegments(['fontName', 'fontSize', 'lineHeight', 'letterSpacing']);
        segments.forEach(segment => {
          const family = segment.fontName.family;
          fontFamilies.set(family, (fontFamilies.get(family) || 0) + 1);
          const detail = ensureFontDetail(family);
          const style = segment.fontName.style;
          if (style) {
            detail.styles.set(style, (detail.styles.get(style) || 0) + 1);
          }
          if (typeof segment.fontSize === 'number') {
            const seg = segment as { lineHeight?: LineHeight | typeof figma.mixed; letterSpacing?: LetterSpacing | typeof figma.mixed };
            const lh = seg.lineHeight !== undefined && seg.lineHeight !== figma.mixed ? (seg.lineHeight as LineHeight) : undefined;
            const ls = seg.letterSpacing !== undefined && seg.letterSpacing !== figma.mixed ? (seg.letterSpacing as LetterSpacing) : undefined;
            addSizeVariant(detail, segment.fontSize, lh, ls);
          }
          if (!availableFamilies.has(family.toLowerCase())) {
            missingFontFamilies.add(family);
          }
        });
      } else {
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
          const lh = node.lineHeight !== figma.mixed ? (node.lineHeight as LineHeight) : undefined;
          const ls = node.letterSpacing !== figma.mixed ? (node.letterSpacing as LetterSpacing) : undefined;
          addSizeVariant(detail, sizeValue, lh, ls);
        }
        if (!availableFamilies.has(family.toLowerCase())) {
          missingFontFamilies.add(family);
        }
      }
    } catch (error) {
      // Ignore nodes we can't read
    }
  });

  const fontsWithCount = Array.from(fontFamilies.entries())
    .map(([name, count]) => ({ name, count }))
    .sort((a, b) => a.name.localeCompare(b.name));

  const fontDetailsPayload: Record<
    string,
    {
      styles: { value: string; count: number }[];
      sizeVariants: { fontSize: number; lineHeightPercent: number | null; letterSpacingPercent: number; count: number }[];
    }
  > = {};

  fontDetails.forEach((detail, family) => {
    const variants = Array.from(detail.sizeVariants.values()).sort((a, b) => {
      if (a.fontSize !== b.fontSize) return a.fontSize - b.fontSize;
      const lhA = a.lineHeightPercent ?? -1;
      const lhB = b.lineHeightPercent ?? -1;
      if (lhA !== lhB) return lhA - lhB;
      return a.letterSpacingPercent - b.letterSpacingPercent;
    });
    fontDetailsPayload[family] = {
      styles: Array.from(detail.styles.entries())
        .map(([value, count]) => ({ value, count }))
        .sort((a, b) => a.value.localeCompare(b.value)),
      sizeVariants: variants
    };
  });

  const familyStylesPayload: Record<string, string[]> = {};
  familyStyleMap.forEach((styles, family) => {
    familyStylesPayload[family] = Array.from(styles).sort((a, b) => a.localeCompare(b));
  });

  figma.ui.postMessage({
    type: 'scan-result',
    fonts: fontsWithCount.map(f => f.name),
    fontCounts: fontsWithCount,
    missingFonts: Array.from(missingFontFamilies).sort(),
    systemFonts: systemFonts,
    fontDetails: fontDetailsPayload,
    familyStyles: familyStylesPayload
  });
}

async function getFontsFromPage() {
  const textNodes = figma.currentPage.findAllWithCriteria({ types: ['TEXT'] });
  await getFontsFromTextNodes(textNodes);
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

async function replaceFontFamily(oldFamily: string, newFamily: string, scope: 'page' | 'selection' = 'page') {
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

  // Get nodes to consider (page or selection)
  const textNodes = getTextNodesForScope(scope);
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
      // Check if this node has missing fonts - handle separately!
      const hasMissingFont = node.hasMissingFont;
      
      if (hasMissingFont) {
        // === SEPARATE PATH FOR MISSING FONTS ===
        console.log('Detected missing font in node, using special handling');
        
        if (node.fontName === figma.mixed) {
          // Mixed text with missing fonts
          const segments = node.getStyledTextSegments(['fontName']);
          let hasChanges = false;
          
          // For missing fonts, we try a more direct approach
          for (const segment of segments) {
            if (segment.fontName.family.toLowerCase() === targetOld) {
              const newFont = findBestStyle(segment.fontName.style);
              
              try {
                // Load the new font and try to apply directly
                await figma.loadFontAsync(newFont);
                
                // Try to set the range
                try {
                  node.setRangeFontName(segment.start, segment.end, newFont);
                  hasChanges = true;
                } catch (rangeError) {
                  // If range fails, try replacing the entire node as last resort
                  console.log('Range replacement failed, trying full node replacement');
                  try {
                    node.fontName = newFont;
                    hasChanges = true;
                    break; // Exit loop if we replaced the entire node
                  } catch (nodeError) {
                    console.error(`Failed to replace entire node`, nodeError);
                  }
                }
              } catch (e) {
                console.error(`Failed to load new font ${newFont.family} ${newFont.style}`, e);
              }
            }
          }
          
          if (hasChanges) updateCount++;
        } else {
          // Single text with missing font
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
      } else {
        // === EXISTING LOGIC FOR NON-MISSING FONTS (WORKING) ===
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
  // Re-scan to update UI (same scope as replace)
  if (scope === 'selection') {
    await getFontsFromTextNodes(collectTextNodesFromSelection());
  } else {
    await getFontsFromPage();
  }
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

async function replaceFontStyleForFamily(family: string, oldStyle: string, newStyle: string, scope: 'page' | 'selection' = 'page') {
  const targetFamily = family.trim().toLowerCase();
  const targetStyle = oldStyle.trim().toLowerCase();
  const replacementStyle = newStyle.trim();

  if (!targetFamily || !targetStyle || !replacementStyle) {
    return;
  }

  const textNodes = getTextNodesForScope(scope);
  const fontLoadCache = new Set<string>();
  let updateCount = 0;

  for (const node of textNodes) {
    let nodeUpdated = false;
    try {
      const hasMissingFont = node.hasMissingFont;
      
      if (hasMissingFont) {
        // === SEPARATE PATH FOR MISSING FONTS ===
        console.log('Detected missing font in node (style change), using special handling');
        
        if (node.fontName === figma.mixed) {
          const segments = node.getStyledTextSegments(['fontName']);
          
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
                
                try {
                  node.setRangeFontName(segment.start, segment.end, newFont);
                  nodeUpdated = true;
                } catch (rangeError) {
                  // Try full node replacement as fallback
                  console.log('Range replacement failed, trying full node replacement');
                  try {
                    node.fontName = newFont;
                    nodeUpdated = true;
                    break;
                  } catch (nodeError) {
                    console.error(`Failed to replace entire node`, nodeError);
                  }
                }
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
      } else {
        // === EXISTING LOGIC FOR NON-MISSING FONTS (WORKING) ===
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
    if (scope === 'selection') {
      await getFontsFromTextNodes(collectTextNodesFromSelection());
    } else {
      await getFontsFromPage();
    }
  }
}

type SizeVariantFilter = { lineHeightPercent: number | null; letterSpacingPercent: number };

function segmentMatchesVariant(
  segment: { fontSize: number; lineHeight?: LineHeight | typeof figma.mixed; letterSpacing?: LetterSpacing | typeof figma.mixed },
  oldSize: number,
  variant: SizeVariantFilter
): boolean {
  if (segment.fontSize !== oldSize) return false;
  const lhRaw = segment.lineHeight === figma.mixed ? undefined : segment.lineHeight;
  const lsRaw = segment.letterSpacing === figma.mixed ? undefined : segment.letterSpacing;
  const lh = normalizeLineHeight(lhRaw as LineHeight | undefined, segment.fontSize);
  const ls = normalizeLetterSpacing(lsRaw as LetterSpacing | undefined, segment.fontSize);
  const lhMatch = (variant.lineHeightPercent === null && lh === null) || (variant.lineHeightPercent !== null && lh === variant.lineHeightPercent);
  const lsMatch = ls === variant.letterSpacingPercent;
  return lhMatch && lsMatch;
}

async function replaceFontSizeForFamily(
  family: string,
  oldSize: number,
  newSize: number,
  scope: 'page' | 'selection' = 'page',
  variant?: SizeVariantFilter
) {
  const targetFamily = family.trim().toLowerCase();
  if (!targetFamily || Number.isNaN(newSize) || newSize <= 0) {
    return;
  }

  const textNodes = getTextNodesForScope(scope);
  const fontLoadCache = new Set<string>();
  let updateCount = 0;
  const segmentFields: ('fontName' | 'fontSize' | 'lineHeight' | 'letterSpacing')[] = variant
    ? ['fontName', 'fontSize', 'lineHeight', 'letterSpacing']
    : ['fontName', 'fontSize'];

  for (const node of textNodes) {
    let nodeUpdated = false;
    try {
      const hasMissingFont = node.hasMissingFont;

      if (hasMissingFont) {
        if (node.fontSize === figma.mixed || node.fontName === figma.mixed) {
          const segments = node.getStyledTextSegments(segmentFields);
          for (const segment of segments) {
            const segFont = segment.fontName;
            if (segFont.family.toLowerCase() !== targetFamily || segment.fontSize !== oldSize) continue;
            if (variant && !segmentMatchesVariant(segment, oldSize, variant)) continue;
            try {
              await loadFontWithCache(segFont, fontLoadCache);
              try {
                node.setRangeFontSize(segment.start, segment.end, newSize);
                nodeUpdated = true;
              } catch {
                try {
                  node.fontSize = newSize;
                  nodeUpdated = true;
                  break;
                } catch (_) {}
              }
            } catch (err) {
              console.error('Failed to load font for size change', err);
            }
          }
        } else {
          const currentFont = node.fontName as FontName;
          if (currentFont.family.toLowerCase() !== targetFamily || (node.fontSize as number) !== oldSize) {
            // skip
          } else if (variant) {
            const lhRaw = node.lineHeight === figma.mixed ? undefined : node.lineHeight;
            const lsRaw = node.letterSpacing === figma.mixed ? undefined : node.letterSpacing;
            const lh = normalizeLineHeight(lhRaw as LineHeight | undefined, node.fontSize as number);
            const ls = normalizeLetterSpacing(lsRaw as LetterSpacing | undefined, node.fontSize as number);
            const lhMatch = (variant.lineHeightPercent === null && lh === null) || (variant.lineHeightPercent !== null && lh === variant.lineHeightPercent);
            if (lhMatch && ls === variant.letterSpacingPercent) {
              try {
                await loadFontWithCache(currentFont, fontLoadCache);
                node.fontSize = newSize;
                nodeUpdated = true;
              } catch (err) {
                console.error('Failed to update font size for node', err);
              }
            }
          } else {
            try {
              await loadFontWithCache(currentFont, fontLoadCache);
              node.fontSize = newSize;
              nodeUpdated = true;
            } catch (err) {
              console.error('Failed to update font size for node', err);
            }
          }
        }
      } else {
        if (node.fontSize === figma.mixed || node.fontName === figma.mixed) {
          const segments = node.getStyledTextSegments(segmentFields);
          const allFontsInNode = new Set<string>();
          for (const segment of segments) {
            const fontKey = `${segment.fontName.family}::${segment.fontName.style}`;
            allFontsInNode.add(fontKey);
          }
          for (const fontKey of allFontsInNode) {
            const [fam, style] = fontKey.split('::');
            try {
              await figma.loadFontAsync({ family: fam, style });
            } catch (e) {}
          }
          for (const segment of segments) {
            const segFont = segment.fontName;
            if (segFont.family.toLowerCase() !== targetFamily || segment.fontSize !== oldSize) continue;
            if (variant && !segmentMatchesVariant(segment, oldSize, variant)) continue;
            try {
              await loadFontWithCache(segFont, fontLoadCache);
              node.setRangeFontSize(segment.start, segment.end, newSize);
              nodeUpdated = true;
            } catch (err) {
              console.error('Failed to update font size for segment', err);
            }
          }
        } else {
          const currentFont = node.fontName as FontName;
          if (currentFont.family.toLowerCase() !== targetFamily || (node.fontSize as number) !== oldSize) {
            // skip
          } else if (variant) {
            const lhRaw = node.lineHeight === figma.mixed ? undefined : node.lineHeight;
            const lsRaw = node.letterSpacing === figma.mixed ? undefined : node.letterSpacing;
            const lh = normalizeLineHeight(lhRaw as LineHeight | undefined, node.fontSize as number);
            const ls = normalizeLetterSpacing(lsRaw as LetterSpacing | undefined, node.fontSize as number);
            const lhMatch = (variant.lineHeightPercent === null && lh === null) || (variant.lineHeightPercent !== null && lh === variant.lineHeightPercent);
            if (lhMatch && ls === variant.letterSpacingPercent) {
              try {
                await loadFontWithCache(currentFont, fontLoadCache);
                node.fontSize = newSize;
                nodeUpdated = true;
              } catch (err) {
                console.error('Failed to update font size for node', err);
              }
            }
          } else {
            try {
              await loadFontWithCache(currentFont, fontLoadCache);
              node.fontSize = newSize;
              nodeUpdated = true;
            } catch (err) {
              console.error('Failed to update font size for node', err);
            }
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
    if (scope === 'selection') {
      await getFontsFromTextNodes(collectTextNodesFromSelection());
    } else {
      await getFontsFromPage();
    }
  }
}

function toFigmaLineHeight(percent: number | null): LineHeight {
  if (percent === null) return { unit: 'AUTO' };
  return { unit: 'PERCENT', value: percent };
}

async function replaceLineHeightForFamily(
  family: string,
  fontSize: number,
  oldLineHeightPercent: number | null,
  newLineHeightPercent: number | null,
  letterSpacingPercent: number,
  scope: 'page' | 'selection' = 'page'
) {
  const targetFamily = family.trim().toLowerCase();
  const variant: SizeVariantFilter = { lineHeightPercent: oldLineHeightPercent, letterSpacingPercent };
  const textNodes = getTextNodesForScope(scope);
  const fontLoadCache = new Set<string>();
  const segmentFields: ('fontName' | 'fontSize' | 'lineHeight' | 'letterSpacing')[] = ['fontName', 'fontSize', 'lineHeight', 'letterSpacing'];
  let updateCount = 0;
  const newLineHeight = toFigmaLineHeight(newLineHeightPercent);

  for (const node of textNodes) {
    let nodeUpdated = false;
    try {
      if (node.fontSize === figma.mixed || node.fontName === figma.mixed) {
        const segments = node.getStyledTextSegments(segmentFields);
        const allFontsInNode = new Set<string>();
        for (const segment of segments) {
          const fontKey = `${segment.fontName.family}::${segment.fontName.style}`;
          allFontsInNode.add(fontKey);
        }
        for (const fontKey of allFontsInNode) {
          const [fam, style] = fontKey.split('::');
          try {
            await figma.loadFontAsync({ family: fam, style });
          } catch (e) {}
        }
        for (const segment of segments) {
          const segFont = segment.fontName;
          if (segFont.family.toLowerCase() !== targetFamily || segment.fontSize !== fontSize) continue;
          if (!segmentMatchesVariant(segment, fontSize, variant)) continue;
          try {
            await loadFontWithCache(segFont, fontLoadCache);
            node.setRangeLineHeight(segment.start, segment.end, newLineHeight);
            nodeUpdated = true;
          } catch (err) {
            console.error('Failed to set range line height', err);
          }
        }
      } else {
        const currentFont = node.fontName as FontName;
        if (currentFont.family.toLowerCase() !== targetFamily || (node.fontSize as number) !== fontSize) continue;
        const lhRaw = node.lineHeight === figma.mixed ? undefined : node.lineHeight;
        const lsRaw = node.letterSpacing === figma.mixed ? undefined : node.letterSpacing;
        const lh = normalizeLineHeight(lhRaw as LineHeight | undefined, node.fontSize as number);
        const ls = normalizeLetterSpacing(lsRaw as LetterSpacing | undefined, node.fontSize as number);
        const lhMatch = (oldLineHeightPercent === null && lh === null) || (oldLineHeightPercent !== null && lh === oldLineHeightPercent);
        if (lhMatch && ls === letterSpacingPercent) {
          try {
            await loadFontWithCache(currentFont, fontLoadCache);
            node.lineHeight = newLineHeight;
            nodeUpdated = true;
          } catch (err) {
            console.error('Failed to set node line height', err);
          }
        }
      }
    } catch (error) {
      console.error('Failed to process node for line height replacement', error);
    }
    if (nodeUpdated) updateCount++;
  }

  const lhOldStr = oldLineHeightPercent == null ? 'Auto' : `${oldLineHeightPercent}%`;
  const lhNewStr = newLineHeightPercent == null ? 'Auto' : `${newLineHeightPercent}%`;
  figma.ui.postMessage({
    type: 'notification',
    message: updateCount === 0
      ? `No ${family} layers found with ${fontSize}px, ${lhOldStr} line height.`
      : `Updated line height ${lhOldStr} → ${lhNewStr} in ${updateCount} layer${updateCount === 1 ? '' : 's'}.`,
    count: updateCount,
    fontName: family
  });
  if (updateCount > 0) {
    if (scope === 'selection') {
      await getFontsFromTextNodes(collectTextNodesFromSelection());
    } else {
      await getFontsFromPage();
    }
  }
}

async function replaceLetterSpacingForFamily(
  family: string,
  fontSize: number,
  lineHeightPercent: number | null,
  oldLetterSpacingPercent: number,
  newLetterSpacingPercent: number,
  scope: 'page' | 'selection' = 'page'
) {
  const targetFamily = family.trim().toLowerCase();
  const variant: SizeVariantFilter = { lineHeightPercent, letterSpacingPercent: oldLetterSpacingPercent };
  const textNodes = getTextNodesForScope(scope);
  const fontLoadCache = new Set<string>();
  const segmentFields: ('fontName' | 'fontSize' | 'lineHeight' | 'letterSpacing')[] = ['fontName', 'fontSize', 'lineHeight', 'letterSpacing'];
  let updateCount = 0;
  const newLetterSpacing: LetterSpacing = { unit: 'PERCENT', value: newLetterSpacingPercent };

  for (const node of textNodes) {
    let nodeUpdated = false;
    try {
      if (node.fontSize === figma.mixed || node.fontName === figma.mixed) {
        const segments = node.getStyledTextSegments(segmentFields);
        const allFontsInNode = new Set<string>();
        for (const segment of segments) {
          const fontKey = `${segment.fontName.family}::${segment.fontName.style}`;
          allFontsInNode.add(fontKey);
        }
        for (const fontKey of allFontsInNode) {
          const [fam, style] = fontKey.split('::');
          try {
            await figma.loadFontAsync({ family: fam, style });
          } catch (e) {}
        }
        for (const segment of segments) {
          const segFont = segment.fontName;
          if (segFont.family.toLowerCase() !== targetFamily || segment.fontSize !== fontSize) continue;
          if (!segmentMatchesVariant(segment, fontSize, variant)) continue;
          try {
            await loadFontWithCache(segFont, fontLoadCache);
            node.setRangeLetterSpacing(segment.start, segment.end, newLetterSpacing);
            nodeUpdated = true;
          } catch (err) {
            console.error('Failed to set range letter spacing', err);
          }
        }
      } else {
        const currentFont = node.fontName as FontName;
        if (currentFont.family.toLowerCase() !== targetFamily || (node.fontSize as number) !== fontSize) continue;
        const lhRaw = node.lineHeight === figma.mixed ? undefined : node.lineHeight;
        const lsRaw = node.letterSpacing === figma.mixed ? undefined : node.letterSpacing;
        const lh = normalizeLineHeight(lhRaw as LineHeight | undefined, node.fontSize as number);
        const ls = normalizeLetterSpacing(lsRaw as LetterSpacing | undefined, node.fontSize as number);
        const lhMatch = (lineHeightPercent === null && lh === null) || (lineHeightPercent !== null && lh === lineHeightPercent);
        if (lhMatch && ls === oldLetterSpacingPercent) {
          try {
            await loadFontWithCache(currentFont, fontLoadCache);
            node.letterSpacing = newLetterSpacing;
            nodeUpdated = true;
          } catch (err) {
            console.error('Failed to set node letter spacing', err);
          }
        }
      }
    } catch (error) {
      console.error('Failed to process node for letter spacing replacement', error);
    }
    if (nodeUpdated) updateCount++;
  }

  figma.ui.postMessage({
    type: 'notification',
    message: updateCount === 0
      ? `No ${family} layers found with ${fontSize}px, ${oldLetterSpacingPercent}% letter spacing.`
      : `Updated letter spacing ${oldLetterSpacingPercent}% → ${newLetterSpacingPercent}% in ${updateCount} layer${updateCount === 1 ? '' : 's'}.`,
    count: updateCount,
    fontName: family
  });
  if (updateCount > 0) {
    if (scope === 'selection') {
      await getFontsFromTextNodes(collectTextNodesFromSelection());
    } else {
      await getFontsFromPage();
    }
  }
}

// --- 3. SELECTION LOGIC ---

async function selectTextNodesByFont(fontFamily: string, scope: 'page' | 'selection' = 'page') {
  const targetFamily = fontFamily.toLowerCase();
  const matches: SceneNode[] = [];
  const textNodes = getTextNodesForScope(scope);

  textNodes.forEach(node => {
    const nodeFonts = extractFontFamilies(node);
    if (Array.from(nodeFonts).some(f => f.toLowerCase() === targetFamily)) {
      matches.push(node);
    }
  });

  figma.currentPage.selection = matches;
  if (matches.length > 0) {
    figma.viewport.scrollAndZoomIntoView(matches);
    const scopeLabel = scope === 'selection' ? ' in selection' : ' on page';
    figma.ui.postMessage({
      type: 'notification',
      message: `${matches.length} "${fontFamily}" layer${matches.length === 1 ? '' : 's'} selected${scopeLabel}!`,
      count: matches.length,
      fontName: fontFamily
    });
  } else {
    figma.ui.postMessage({
      type: 'notification',
      message: scope === 'selection' ? 'No layers with this font in selection.' : 'No layers found on this page.',
      count: 0
    });
  }
}

function selectTextNodesByFontVariant(
  fontFamily: string,
  fontSize: number,
  lineHeightPercent: number | null,
  letterSpacingPercent: number,
  scope: 'page' | 'selection' = 'page'
) {
  const targetFamily = fontFamily.trim().toLowerCase();
  const variant: SizeVariantFilter = { lineHeightPercent, letterSpacingPercent };
  const textNodes = getTextNodesForScope(scope);
  const segmentFields: ('fontName' | 'fontSize' | 'lineHeight' | 'letterSpacing')[] = ['fontName', 'fontSize', 'lineHeight', 'letterSpacing'];
  const matches: SceneNode[] = [];

  textNodes.forEach(node => {
    let hasMatch = false;
    try {
      if (node.fontSize === figma.mixed || node.fontName === figma.mixed) {
        const segments = node.getStyledTextSegments(segmentFields);
        for (const segment of segments) {
          if (segment.fontName.family.toLowerCase() !== targetFamily || segment.fontSize !== fontSize) continue;
          if (segmentMatchesVariant(segment, fontSize, variant)) {
            hasMatch = true;
            break;
          }
        }
      } else {
        const currentFont = node.fontName as FontName;
        if (currentFont.family.toLowerCase() === targetFamily && (node.fontSize as number) === fontSize) {
          const lhRaw = node.lineHeight === figma.mixed ? undefined : node.lineHeight;
          const lsRaw = node.letterSpacing === figma.mixed ? undefined : node.letterSpacing;
          const lh = normalizeLineHeight(lhRaw as LineHeight | undefined, node.fontSize as number);
          const ls = normalizeLetterSpacing(lsRaw as LetterSpacing | undefined, node.fontSize as number);
          const lhMatch = (lineHeightPercent === null && lh === null) || (lineHeightPercent !== null && lh === lineHeightPercent);
          if (lhMatch && ls === letterSpacingPercent) hasMatch = true;
        }
      }
      if (hasMatch) matches.push(node);
    } catch (_) {}
  });

  figma.currentPage.selection = matches;
  if (matches.length > 0) {
    figma.viewport.scrollAndZoomIntoView(matches);
    const scopeLabel = scope === 'selection' ? ' in selection' : ' on page';
    const lhStr = lineHeightPercent == null ? 'Auto' : `${lineHeightPercent}%`;
    figma.ui.postMessage({
      type: 'notification',
      message: `${matches.length} layer${matches.length === 1 ? '' : 's'} selected (${fontFamily} ${fontSize}px, ${lhStr}, ${letterSpacingPercent}%)${scopeLabel}!`,
      count: matches.length,
      fontName: fontFamily
    });
  } else {
    figma.ui.postMessage({
      type: 'notification',
      message: scope === 'selection' ? 'No layers with this variant in selection.' : 'No layers found on this page.',
      count: 0
    });
  }
}

function selectTextNodesByFontAndStyle(
  fontFamily: string,
  fontStyle: string,
  scope: 'page' | 'selection' = 'page'
) {
  const targetFamily = fontFamily.trim().toLowerCase();
  const targetStyle = fontStyle.trim().toLowerCase();
  const textNodes = getTextNodesForScope(scope);
  const matches: SceneNode[] = [];

  textNodes.forEach(node => {
    let hasMatch = false;
    try {
      if (node.fontName === figma.mixed) {
        const segments = node.getStyledTextSegments(['fontName']);
        for (const segment of segments) {
          if (
            segment.fontName.family.toLowerCase() === targetFamily &&
            (segment.fontName.style || '').toLowerCase() === targetStyle
          ) {
            hasMatch = true;
            break;
          }
        }
      } else {
        const fontName = node.fontName as FontName;
        if (
          fontName.family.toLowerCase() === targetFamily &&
          (fontName.style || '').toLowerCase() === targetStyle
        ) {
          hasMatch = true;
        }
      }
      if (hasMatch) matches.push(node);
    } catch (_) {}
  });

  figma.currentPage.selection = matches;
  if (matches.length > 0) {
    figma.viewport.scrollAndZoomIntoView(matches);
    const scopeLabel = scope === 'selection' ? ' in selection' : ' on page';
    figma.ui.postMessage({
      type: 'notification',
      message: `${matches.length} layer${matches.length === 1 ? '' : 's'} selected (${fontFamily} ${fontStyle})${scopeLabel}!`,
      count: matches.length,
      fontName: fontFamily
    });
  } else {
    figma.ui.postMessage({
      type: 'notification',
      message: scope === 'selection' ? 'No layers with this weight in selection.' : 'No layers found on this page.',
      count: 0
    });
  }
}

// --- 4. SELECTION CHANGE TRACKING ---

let lastSelectedFont: string | null = null;

function notifySelectionChanged() {
  const count = getSelectedContainerCount();
  figma.ui.postMessage({ type: 'selection-changed', count });
}

// Notify UI of initial selection when plugin opens
setTimeout(() => notifySelectionChanged(), 100);

// Listen for selection changes in Figma
figma.on('selectionchange', () => {
  const selection = figma.currentPage.selection;

  notifySelectionChanged();

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
  } else if (msg.type === 'scan-selection') {
    const textNodes = collectTextNodesFromSelection();
    await getFontsFromTextNodes(textNodes);
  } else if (msg.type === 'select-font') {
    lastSelectedFont = msg.font;
    const scope = (msg.scope === 'selection' ? 'selection' : 'page') as 'page' | 'selection';
    await selectTextNodesByFont(msg.font, scope);
  } else if (msg.type === 'select-font-variant') {
    const scope = (msg.scope === 'selection' ? 'selection' : 'page') as 'page' | 'selection';
    selectTextNodesByFontVariant(
      msg.family,
      msg.fontSize,
      msg.lineHeightPercent,
      msg.letterSpacingPercent,
      scope
    );
  } else if (msg.type === 'select-font-weight') {
    const scope = (msg.scope === 'selection' ? 'selection' : 'page') as 'page' | 'selection';
    selectTextNodesByFontAndStyle(msg.family, msg.style, scope);
  } else if (msg.type === 'replace-font') {
    const scope = (msg.scope === 'selection' ? 'selection' : 'page') as 'page' | 'selection';
    await replaceFontFamily(msg.oldFont, msg.newFont, scope);
  } else if (msg.type === 'replace-font-weight') {
    const scope = (msg.scope === 'selection' ? 'selection' : 'page') as 'page' | 'selection';
    await replaceFontStyleForFamily(msg.family, msg.oldStyle, msg.newStyle, scope);
  } else if (msg.type === 'replace-font-size') {
    const scope = (msg.scope === 'selection' ? 'selection' : 'page') as 'page' | 'selection';
    const variant = msg.variant as SizeVariantFilter | undefined;
    await replaceFontSizeForFamily(msg.family, msg.oldSize, msg.newSize, scope, variant);
  } else if (msg.type === 'replace-font-line-height') {
    const scope = (msg.scope === 'selection' ? 'selection' : 'page') as 'page' | 'selection';
    await replaceLineHeightForFamily(
      msg.family,
      msg.fontSize,
      msg.oldLineHeightPercent,
      msg.newLineHeightPercent,
      msg.letterSpacingPercent,
      scope
    );
  } else if (msg.type === 'replace-font-letter-spacing') {
    const scope = (msg.scope === 'selection' ? 'selection' : 'page') as 'page' | 'selection';
    await replaceLetterSpacingForFamily(
      msg.family,
      msg.fontSize,
      msg.lineHeightPercent,
      msg.oldLetterSpacingPercent,
      msg.newLetterSpacingPercent,
      scope
    );
  }
};