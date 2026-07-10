import {
  fontDetailsGlobal,
  familyStylesGlobal,
  lastScanScope,
  postToPlugin
} from '../state';
import { splitWeightGroups } from '../weights';
import { showNotification } from '../notifications';
import { showReplaceWeightModal } from '../replace/modals';

export function setInlineDetailsVisible(
  row: HTMLElement,
  shouldShow: boolean,
  fontName?: string,
  isMissing?: boolean
) {
  const details = row.querySelector('.font-details');
  if (!details) return;
  if (!shouldShow) {
    details.classList.remove('show');
    details.innerHTML = '';
    return;
  }
  if (!fontName) return;
  const fontData = fontDetailsGlobal[fontName];
  if (!fontData) {
    details.innerHTML = '<div style="color: #999; font-size: 12px; padding: 8px 0;">No usage data yet.</div>';
    details.classList.add('show');
    return;
  }
  if (isMissing) {
    details.innerHTML = '<div style="color: #999; font-size: 12px; padding: 0 0 8px 0;">Missing font - no details available.</div>';
    details.classList.add('show');
    return;
  }

  // Clear previous content
  details.innerHTML = '';
  
  // Get available styles for this font
  const availableStyles = familyStylesGlobal[fontName] || [];
  
  // Weights section
  const weights = fontData.styles || [];
  if (weights.length > 0) {
    
    weights.forEach(item => {
      const itemDiv = document.createElement('div');
      itemDiv.className = 'font-detail-item';
      
      const labelDiv = document.createElement('div');
      labelDiv.className = 'font-detail-label';
      
      const valueSpan = document.createElement('div');
      valueSpan.className = 'font-detail-value';
      valueSpan.textContent = item.value;
      
      const countSpan = document.createElement('div');
      countSpan.className = 'font-detail-count';
        countSpan.textContent = `${item.count} layer${item.count === 1 ? '' : 's'}`;

      labelDiv.appendChild(valueSpan);
      labelDiv.appendChild(countSpan);
      
      const controlDiv = document.createElement('div');
      controlDiv.className = 'font-detail-control';
      
      // Create dropdown for weight selection
      if (availableStyles.length > 0) {
        const select = document.createElement('select');
        select.className = 'font-detail-select';
        availableStyles.forEach(style => {
          const option = document.createElement('option');
          option.value = style;
          option.textContent = style;
          select.appendChild(option);
        });
        
        // Set the current weight as the selected value
        select.value = item.value;
        
        const applyBtn = document.createElement('button');
        applyBtn.className = 'font-detail-btn';
        applyBtn.textContent = 'Apply';
        applyBtn.onclick = () => {
          const newStyle = select.value;
          if (newStyle && newStyle !== item.value) {
            // Disable button and show loading state
            applyBtn.disabled = true;
            applyBtn.textContent = 'Applying...';
            applyBtn.style.opacity = '0.6';
            
            // Show toast notification
            showNotification(`Updating ${fontName} ${item.value} → ${newStyle}...`, 0);
            
            postToPlugin({
                type: 'replace-font-weight',
                family: fontName,
                oldStyle: item.value,
                newStyle: newStyle,
                scope: lastScanScope
              });
          }
        };
        
        controlDiv.appendChild(select);
        controlDiv.appendChild(applyBtn);
      }
      
      itemDiv.appendChild(labelDiv);
      itemDiv.appendChild(controlDiv);
      
      itemDiv.onclick = (e) => {
        if (e.target.closest('.font-detail-control')) return;
        postToPlugin({
            type: 'select-font-weight',
            family: fontName,
            style: item.value,
            scope: lastScanScope
          });
      };
      
      if (controlDiv.children.length > 0) {
        controlDiv.onclick = (e) => e.stopPropagation();
      }
      
      details.appendChild(itemDiv);
    });
  }
  
  // Sizes section (grouped by line height & letter spacing)
  let sizeVariants = fontData.sizeVariants || [];
  if (sizeVariants.length === 0 && (fontData.sizes || []).length > 0) {
    sizeVariants = fontData.sizes.map(s => ({
      fontSize: s.value,
      lineHeightPercent: null,
      letterSpacingPercent: 0,
      count: s.count
    }));
  }
  if (sizeVariants.length > 0) {
    const listContainer = document.createElement('div');
    listContainer.className = 'size-list';
    
    const headerRow = document.createElement('div');
    headerRow.className = 'size-list-header';
    headerRow.innerHTML = '<span>Size</span><span>Line height</span><span>Letter spacing</span><span class="size-list-col-count">Layers</span>';
    listContainer.appendChild(headerRow);
    
    sizeVariants.forEach(item => {
      const fontSize = item.fontSize;
      const sizeDisplay = Number.isInteger(fontSize) ? fontSize : parseFloat(fontSize.toFixed(2));
      const lhDisplay = item.lineHeightPercent == null ? 'Auto' : `${item.lineHeightPercent}%`;
      const lsDisplay = `${item.letterSpacingPercent}%`;
      
      const row = document.createElement('div');
      row.className = 'size-list-row';
      
      const valueSpan = document.createElement('span');
      valueSpan.className = 'size-list-value';
      valueSpan.textContent = String(sizeDisplay);
      
      const lhSpan = document.createElement('span');
      lhSpan.className = 'size-list-lh';
      lhSpan.textContent = lhDisplay;
      
      const lsSpan = document.createElement('span');
      lsSpan.className = 'size-list-ls';
      lsSpan.textContent = lsDisplay;
      
      const countSpan = document.createElement('span');
      countSpan.className = 'size-list-count';
      countSpan.textContent = `${item.count} layer${item.count === 1 ? '' : 's'}`;
      
      row.appendChild(valueSpan);
      row.appendChild(lhSpan);
      row.appendChild(lsSpan);
      row.appendChild(countSpan);
      
      let editingField = null;
      let originalSize = sizeDisplay;
      let originalLh = lhDisplay;
      let originalLs = lsDisplay;
      
      function cancelEdit() {
        if (!editingField) return;
        if (editingField === 'size') {
          valueSpan.contentEditable = 'false';
          valueSpan.textContent = String(originalSize);
        } else if (editingField === 'lh') {
          lhSpan.contentEditable = 'false';
          lhSpan.textContent = originalLh;
        } else if (editingField === 'ls') {
          lsSpan.contentEditable = 'false';
          lsSpan.textContent = originalLs;
        }
        row.classList.remove('editing');
        editingField = null;
      }
      
      function startEdit(field, span, text) {
        if (editingField) cancelEdit();
        editingField = field;
        row.classList.add('editing');
        span.contentEditable = 'true';
        span.textContent = text;
        span.focus();
        const range = document.createRange();
        range.selectNodeContents(span);
        const sel = window.getSelection();
        sel.removeAllRanges();
        sel.addRange(range);
      }
      
      valueSpan.onclick = (e) => {
        e.stopPropagation();
        if (editingField) return;
        startEdit('size', valueSpan, String(originalSize));
      };
      lhSpan.onclick = (e) => {
        e.stopPropagation();
        if (editingField) return;
        startEdit('lh', lhSpan, item.lineHeightPercent == null ? 'Auto' : String(item.lineHeightPercent));
      };
      lsSpan.onclick = (e) => {
        e.stopPropagation();
        if (editingField) return;
        startEdit('ls', lsSpan, String(item.letterSpacingPercent));
      };
      
      row.onclick = (e) => {
        if (e.target.closest('.size-list-value') || e.target.closest('.size-list-lh') || e.target.closest('.size-list-ls')) return;
        postToPlugin({
            type: 'select-font-variant',
            family: fontName,
            fontSize,
            lineHeightPercent: item.lineHeightPercent,
            letterSpacingPercent: item.letterSpacingPercent,
            scope: lastScanScope
          });
      };
      
      function applySize() {
        const newValue = parseFloat(valueSpan.textContent.replace(/,/g, '.'));
        if (isNaN(newValue) || newValue <= 0) {
          valueSpan.textContent = String(originalSize);
          cancelEdit();
          return;
        }
        if (Math.abs(newValue - fontSize) < 0.01) {
          cancelEdit();
          return;
        }
        editingField = null;
        row.classList.remove('editing');
        row.classList.add('updating');
        valueSpan.contentEditable = 'false';
        const formatted = Number.isInteger(newValue) ? newValue : parseFloat(newValue.toFixed(2));
        valueSpan.textContent = String(formatted);
        originalSize = formatted;
        showNotification(`Updating ${fontName} ${sizeDisplay}px → ${newValue}px...`, 0);
        postToPlugin({
            type: 'replace-font-size',
            family: fontName,
            oldSize: fontSize,
            newSize: newValue,
            scope: lastScanScope,
            variant: { lineHeightPercent: item.lineHeightPercent, letterSpacingPercent: item.letterSpacingPercent }
          });
        setTimeout(() => row.classList.remove('updating'), 500);
      }
      
      function applyLineHeight() {
        const raw = (lhSpan.textContent || '').trim();
        const isAuto = /^auto$/i.test(raw);
        const newValue = isAuto ? null : parseFloat(raw.replace(/,/g, '.'));
        if (!isAuto && (isNaN(newValue) || newValue < 0)) {
          lhSpan.textContent = originalLh;
          cancelEdit();
          return;
        }
        const same = (item.lineHeightPercent == null && newValue === null) ||
          (item.lineHeightPercent !== null && newValue !== null && Math.abs(item.lineHeightPercent - newValue) < 0.01);
        if (same) {
          cancelEdit();
          return;
        }
        editingField = null;
        row.classList.remove('editing');
        row.classList.add('updating');
        lhSpan.contentEditable = 'false';
        originalLh = newValue == null ? 'Auto' : `${newValue}%`;
        lhSpan.textContent = originalLh;
        const oldStr = item.lineHeightPercent == null ? 'Auto' : `${item.lineHeightPercent}%`;
        showNotification(`Updating line height ${oldStr} → ${originalLh}...`, 0);
        postToPlugin({
            type: 'replace-font-line-height',
            family: fontName,
            fontSize,
            oldLineHeightPercent: item.lineHeightPercent,
            newLineHeightPercent: newValue,
            letterSpacingPercent: item.letterSpacingPercent,
            scope: lastScanScope
          });
        setTimeout(() => row.classList.remove('updating'), 500);
      }
      
      function applyLetterSpacing() {
        const newValue = parseFloat(lsSpan.textContent.replace(/,/g, '.'));
        if (isNaN(newValue)) {
          lsSpan.textContent = originalLs;
          cancelEdit();
          return;
        }
        if (Math.abs(newValue - item.letterSpacingPercent) < 0.01) {
          cancelEdit();
          return;
        }
        editingField = null;
        row.classList.remove('editing');
        row.classList.add('updating');
        lsSpan.contentEditable = 'false';
        const formatted = Math.round(newValue * 10) / 10;
        originalLs = `${formatted}%`;
        lsSpan.textContent = originalLs;
        showNotification(`Updating letter spacing ${item.letterSpacingPercent}% → ${formatted}%...`, 0);
        postToPlugin({
            type: 'replace-font-letter-spacing',
            family: fontName,
            fontSize,
            lineHeightPercent: item.lineHeightPercent,
            oldLetterSpacingPercent: item.letterSpacingPercent,
            newLetterSpacingPercent: formatted,
            scope: lastScanScope
          });
        setTimeout(() => row.classList.remove('updating'), 500);
      }
      
      valueSpan.addEventListener('keydown', (e) => {
        if (e.key === 'Enter') { e.preventDefault(); applySize(); }
        else if (e.key === 'Escape') { e.preventDefault(); cancelEdit(); }
      });
      valueSpan.addEventListener('input', () => {
        const text = valueSpan.textContent || '';
        const cleaned = text.replace(/[^\d.,]/g, '').replace(',', '.');
        if (text !== cleaned) valueSpan.textContent = cleaned;
      });
      
      lhSpan.addEventListener('keydown', (e) => {
        if (e.key === 'Enter') { e.preventDefault(); applyLineHeight(); }
        else if (e.key === 'Escape') { e.preventDefault(); cancelEdit(); }
      });
      lsSpan.addEventListener('keydown', (e) => {
        if (e.key === 'Enter') { e.preventDefault(); applyLetterSpacing(); }
        else if (e.key === 'Escape') { e.preventDefault(); cancelEdit(); }
      });
      lsSpan.addEventListener('input', () => {
        const text = lsSpan.textContent || '';
        const cleaned = text.replace(/[^\d.,-]/g, '').replace(',', '.');
        if (text !== cleaned) lsSpan.textContent = cleaned;
      });
      
      function onBlur() {
        setTimeout(() => { if (editingField) cancelEdit(); }, 150);
      }
      valueSpan.addEventListener('blur', onBlur);
      lhSpan.addEventListener('blur', onBlur);
      lsSpan.addEventListener('blur', onBlur);
      
      listContainer.appendChild(row);
    });
    
    details.appendChild(listContainer);
  }
  
  if (weights.length === 0 && sizeVariants.length === 0) {
    details.innerHTML = '<div style="color: #999; font-size: 12px; padding: 8px 0;">No weights or sizes detected.</div>';
  }
  
  details.classList.add('show');
}
