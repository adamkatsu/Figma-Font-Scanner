import {
  selectedFontItem,
  selectedFontName,
  setSelectedFontItem,
  setSelectedFontName,
  setSelectedFrameCount,
  setFamilyStyles,
  setSystemFonts,
  postToPlugin
} from './state';
import { updateScanButtonLabel, performScan } from './scan';
import {
  handleScanProgress,
  handleReplacementProgress,
  showNotification
} from './notifications';
import { handleScanResult, restoreFontSelection } from './render/fontList';
import { setInlineDetailsVisible } from './render/details';
import { closeReplaceModal } from './replace/modals';

function clearSelectionUi() {
  if (!selectedFontItem) return;
  setInlineDetailsVisible(selectedFontItem, false);
  selectedFontItem.classList.remove('selected');
  const dropdownIcon = selectedFontItem.querySelector('.dropdown-icon');
  if (dropdownIcon) dropdownIcon.classList.remove('open');
  setSelectedFontItem(null);
  setSelectedFontName(null);
}

document.getElementById('scanBtn')!.onclick = () => performScan();

document.addEventListener('click', e => {
  const target = e.target as HTMLElement;
  if (!target.closest('.item') && !target.closest('.modal-overlay') && selectedFontItem) {
    clearSelectionUi();
    postToPlugin({ type: 'deselect-font' });
  }
});

document.getElementById('modalClose')!.onclick = () => closeReplaceModal();
document.getElementById('replaceModal')!.onclick = e => {
  if ((e.target as HTMLElement).id === 'replaceModal') {
    closeReplaceModal();
  }
};

onmessage = async event => {
  const msg = event.data.pluginMessage;
  if (!msg) return;

  if (msg.type === 'selection-changed') {
    setSelectedFrameCount(msg.count ?? 0);
    updateScanButtonLabel();
    return;
  }

  if (msg.type === 'scan-progress') {
    handleScanProgress(msg.current ?? 0, msg.total ?? 0);
    return;
  }

  if (msg.type === 'scan-result') {
    if (msg.familyStyles) setFamilyStyles(msg.familyStyles);
    const previouslySelectedFont = selectedFontName;
    handleScanResult(msg);
    if (previouslySelectedFont) {
      restoreFontSelection(previouslySelectedFont);
    }
    return;
  }

  if (msg.type === 'system-fonts') {
    setSystemFonts(msg.systemFonts || []);
    return;
  }

  if (msg.type === 'notification') {
    showNotification(msg.message, msg.count, msg.fontName);
    return;
  }

  if (msg.type === 'replacement-progress') {
    handleReplacementProgress(msg.current, msg.total, msg.fontName);
    return;
  }

  if (msg.type === 'deselect-font') {
    clearSelectionUi();
  }
};

// Initial scan on plugin open
performScan();
