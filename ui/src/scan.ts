import {
  selectedFontItem,
  selectedFrameCount,
  setLastScanScope,
  setSelectedFontItem,
  postToPlugin
} from './state';
import { showScanOverlay } from './notifications';

export function updateScanButtonLabel() {
  const label = document.getElementById('scanBtnLabel');
  if (!label) return;
  if (selectedFrameCount === 0) {
    label.textContent = 'Scan This Page';
  } else if (selectedFrameCount === 1) {
    label.textContent = 'Scan Selected Frame';
  } else {
    label.textContent = `Scan ${selectedFrameCount} Frames`;
  }
}

export function performScan() {
  const results = document.getElementById('resultsContainer');
  const loader = document.getElementById('loader');
  if (results) results.innerHTML = '';
  // Prefer the progress overlay immediately — do not wait for the first
  // scan-progress message (that arrives only after listAvailableFontsAsync).
  if (loader) loader.style.display = 'none';
  showScanOverlay();

  if (selectedFontItem) {
    selectedFontItem.classList.remove('selected');
    setSelectedFontItem(null);
  }

  const type = selectedFrameCount > 0 ? 'scan-selection' : 'scan-layers';
  setLastScanScope(selectedFrameCount > 0 ? 'selection' : 'page');
  postToPlugin({ type });
}
