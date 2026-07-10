import {
  setFontDetails,
  selectedFontItem,
  setSelectedFontItem,
  setSelectedFontName,
  lastScanScope,
  postToPlugin,
  FontDetailsMap
} from '../state';
import { getGoogleFontSet, isGoogleFont, isIconFont } from '../googleFonts';
import { setInlineDetailsVisible } from './details';
import { showReplaceModal } from '../replace/modals';
import { showNotification, hideScanOverlay } from '../notifications';

export function restoreFontSelection(fontName: string) {
  const allItems = document.querySelectorAll('.item');
  allItems.forEach(node => {
    const item = node as HTMLElement;
    const nameElement = item.querySelector('.font-name');
    if (nameElement && nameElement.textContent && nameElement.textContent.includes(fontName)) {
      item.classList.add('selected');
      setSelectedFontItem(item);
      setSelectedFontName(fontName);

      const dropdownIcon = item.querySelector('.dropdown-icon');
      if (dropdownIcon) dropdownIcon.classList.add('open');

      const isMissing = item.querySelector('.badge-missing') !== null;
      setInlineDetailsVisible(item, true, fontName, isMissing);
    }
  });
}

function renderSection(
  container: HTMLElement,
  title: string,
  fonts: string[],
  googleSet: ReadonlySet<string>,
  isMissing: boolean,
  countMap: Map<string, number>
) {
  const header = document.createElement('div');
  header.className = 'section-header';
  
  let headerHTML = `<h3>${title} <span style="color: #999; font-weight: 400;">(${fonts.length})</span></h3>`;
  
  // ADD THE DISCLAIMER FOR MISSING FONTS HERE
  if (isMissing) {
    header.style.flexDirection = 'column';
    header.style.alignItems = 'flex-start';
  }

  header.innerHTML = headerHTML;
  container.appendChild(header);

  const ul = document.createElement('ul');
  ul.className = 'list';
  fonts.forEach((font: string) => {
    const count = countMap.get(font);
    // If count is undefined, it means the font wasn't in the countMap - this shouldn't happen but fallback to 1
    const finalCount = count !== undefined && count !== null ? count : 1;
    ul.appendChild(createRow(font, googleSet.has(font.toLowerCase()), isMissing, finalCount));
  });
  container.appendChild(ul);
}

export function createRow(
  fontName: string,
  isGoogle: boolean,
  isMissing: boolean,
  count: number
) {
  const wrapper = document.createElement('li');
  wrapper.className = 'c';

  const mainRow = document.createElement('div');
  mainRow.className = 'item';

  // Create main row container for font name and actions
  const itemMainRow = document.createElement('div');
  itemMainRow.className = 'item-main-row';

  const nameDiv = document.createElement('div');
  nameDiv.className = 'font-name';
  const displayCount = count || 0;
  const countBadge = `<span style="color: #999; font-weight: 400; font-size: 12px; margin-left: 4px;">(${displayCount})</span>`;
  
  // Create dropdown triangle icon
  const dropdownIcon = document.createElement('div');
  dropdownIcon.className = 'dropdown-icon';
  dropdownIcon.innerHTML = `<i class="fa-solid fa-chevron-right"></i>`;
  
  nameDiv.appendChild(dropdownIcon);
  
  const nameText = document.createElement('span');
  nameText.innerHTML = `${fontName}${countBadge} ${isMissing ? '<span class="badge-missing">Missing</span>' : ''}`;
  nameDiv.appendChild(nameText);

  const actionsDiv = document.createElement('div');
  actionsDiv.className = 'actions';

  const linkGroup = document.createElement('div');
  linkGroup.className = 'link-group';
  const encoded = fontName.replace(/\s+/g, '+');

  // Select fonts button (before download) – selects all layers using this font (in selection or page)
  const selectBtn = document.createElement('button');
  selectBtn.className = 'ext-link';
  selectBtn.type = 'button';
  selectBtn.setAttribute('data-tooltip', 'Select fonts');
  selectBtn.innerHTML = `<i class="fa-solid fa-arrow-pointer"></i>`;
  selectBtn.onclick = (event) => {
    event.stopPropagation();
    postToPlugin({ type: 'select-font', font: fontName, scope: lastScanScope });
  };
  linkGroup.appendChild(selectBtn);

  const appendLink = (href: string, tooltip: string, svgMarkup: string) => {
    const anchor = document.createElement('a');
    anchor.href = href;
    anchor.className = 'ext-link';
    anchor.setAttribute('data-tooltip', tooltip);
    anchor.target = '_blank';
    anchor.innerHTML = svgMarkup;
    linkGroup.appendChild(anchor);
  };

  if (isGoogle) {
    appendLink(
      `https://fonts.google.com/download?family=${encoded}`,
      'Download font',
      `<i class="fa-solid fa-download"></i>`
    );
    appendLink(
      `https://fonts.google.com/specimen/${encoded}`,
      'View in browser',
      `<i class="fa-solid fa-globe"></i>`
    );
  } else {
    appendLink(
      `https://www.google.com/search?q=${encoded}+font`,
      'Search in browser',
      `<i class="fa-solid fa-globe"></i>`
    );
  }

  const replaceBtn = document.createElement('button');
  replaceBtn.className = 'ext-link';
  replaceBtn.type = 'button';
  replaceBtn.setAttribute('data-tooltip', 'Replace font');
  replaceBtn.innerHTML = `<i class="fa-solid fa-arrows-rotate"></i>`;
  replaceBtn.onclick = (event) => {
    event.stopPropagation();
    showReplaceModal(fontName);
  };
  linkGroup.appendChild(replaceBtn);

  actionsDiv.appendChild(linkGroup);

  // Assemble the main row
  itemMainRow.appendChild(nameDiv);
  itemMainRow.appendChild(actionsDiv);

  // Create details section
  const detailsDiv = document.createElement('div');
  detailsDiv.className = 'font-details';

  // Append to main row
  mainRow.appendChild(itemMainRow);
  mainRow.appendChild(detailsDiv);

  // Make action buttons stop propagation to prevent triggering the dropdown
  actionsDiv.onclick = (e) => {
    e.stopPropagation();
  };
  
  // Handle item-main-row click to toggle details
  itemMainRow.onclick = (e) => {
    const target = e.target as HTMLElement | null;
    // Don't trigger if clicking on actions or links
    if (target?.closest('.actions') || target?.closest('.ext-link') || target?.closest('a')) {
      return;
    }
    
    const isCurrentlySelected = mainRow.classList.contains('selected');
    
    if (isCurrentlySelected) {
      // Close details if already open
      setInlineDetailsVisible(mainRow, false);
      mainRow.classList.remove('selected');
      dropdownIcon.classList.remove('open');
      setSelectedFontItem(null);
      setSelectedFontName(null);
      postToPlugin({ type: 'deselect-font' });
    } else {
      // Close previous item if any
      if (selectedFontItem) {
        setInlineDetailsVisible(selectedFontItem, false);
        selectedFontItem.classList.remove('selected');
        const prevIcon = selectedFontItem.querySelector('.dropdown-icon');
        if (prevIcon) prevIcon.classList.remove('open');
      }
      
      // Open details for this item (no Figma selection)
      mainRow.classList.add('selected');
      dropdownIcon.classList.add('open');
      setSelectedFontItem(mainRow);
      setSelectedFontName(fontName);
      setInlineDetailsVisible(mainRow, true, fontName, isMissing);
    }
  };
  
  wrapper.appendChild(mainRow);

  return wrapper;
}

export function processFonts(
  figmaFonts: string[],
  missingFonts: string[],
  fontCounts: { name: string; count: number }[] = [],
  fontDetails: FontDetailsMap = {}
) {
  const googleFontNames = getGoogleFontSet();
  const countMap = new Map<string, number>();

  if (fontCounts && fontCounts.length > 0) {
    fontCounts.forEach(f => {
      countMap.set(f.name, f.count);
    });
  } else {
    figmaFonts.forEach(f => countMap.set(f, 1));
  }

  const loader = document.getElementById('loader');
  if (loader) loader.style.display = 'none';
  const container = document.getElementById('resultsContainer')!;
  container.innerHTML = '';
  setFontDetails(fontDetails || {});

  if (figmaFonts.length === 0 && missingFonts.length === 0) {
    container.innerHTML =
      '<p style="margin-top:24px; color:#999; font-size:13px; text-align:center;">No text layers found.</p>';
    return;
  }

  const missingSet = new Set(missingFonts);
  const installedFonts = figmaFonts.filter(f => !missingSet.has(f));

  if (missingFonts.length > 0) {
    renderSection(container, 'Missing Fonts', missingFonts, googleFontNames, true, countMap);
  }

  const iconMatches: string[] = [];
  const googleMatches: string[] = [];
  const localMatches: string[] = [];

  installedFonts.forEach(font => {
    if (isIconFont(font)) {
      iconMatches.push(font);
    } else if (isGoogleFont(font)) {
      googleMatches.push(font);
    } else {
      localMatches.push(font);
    }
  });

  if (googleMatches.length > 0) renderSection(container, 'Google Fonts', googleMatches, googleFontNames, false, countMap);
  if (localMatches.length > 0) renderSection(container, 'Local / System Fonts', localMatches, googleFontNames, false, countMap);
  if (iconMatches.length > 0) renderSection(container, 'Icon Fonts', iconMatches, googleFontNames, false, countMap);
}

export function handleScanResult(msg: {
  fonts?: string[];
  missingFonts?: string[];
  fontCounts?: { name: string; count: number }[];
  fontDetails?: FontDetailsMap;
  familyStyles?: Record<string, string[]>;
  textNodeCount?: number;
  skippedCount?: number;
  largePage?: boolean;
}) {
  hideScanOverlay();
  setFontDetails(msg.fontDetails || {});
  processFonts(msg.fonts || [], msg.missingFonts || [], msg.fontCounts || [], msg.fontDetails || {});

  if (msg.skippedCount && msg.skippedCount > 0) {
    showNotification(`Skipped ${msg.skippedCount} unreadable text layer(s).`, msg.skippedCount);
  }
  if (msg.largePage) {
    showNotification(`Large page (${msg.textNodeCount} text layers).`, 0);
  }
}
