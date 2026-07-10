export function showNotification(message: string, count?: number, _fontName?: string) {
  const toast = document.getElementById('notificationToast');
  if (!toast) return;

  let displayMessage = message;
  if (count && count > 0) {
    displayMessage = message.replace(/(\d+)/, '<span class="notification-count">$1</span>');
  }

  toast.innerHTML = displayMessage;
  toast.classList.add('show');

  setTimeout(() => {
    toast.classList.remove('show');
  }, 3000);
}

function setOverlayTitle(title: string) {
  const el = document.querySelector('.loading-title');
  if (el) el.textContent = title;
}

/** Show scan overlay immediately (before main-thread progress messages arrive). */
export function showScanOverlay(progressText = 'Preparing…') {
  const overlay = document.getElementById('loadingOverlay');
  const progress = document.getElementById('loadingProgress');
  const loader = document.getElementById('loader');
  if (!overlay || !progress) return;
  setOverlayTitle('Scanning...');
  progress.textContent = progressText;
  overlay.classList.add('show');
  if (loader) loader.style.display = 'none';
}

export function hideScanOverlay() {
  const overlay = document.getElementById('loadingOverlay');
  if (!overlay) return;
  overlay.classList.remove('show');
  setOverlayTitle('Replacing fonts...');
}

export function handleReplacementProgress(current: number, total: number, _fontName?: string) {
  const overlay = document.getElementById('loadingOverlay');
  const progress = document.getElementById('loadingProgress');
  if (!overlay || !progress) return;

  setOverlayTitle('Replacing fonts...');

  if (current === 0) {
    overlay.classList.add('show');
    progress.textContent = `0/${total}`;
  } else if (current >= total) {
    progress.textContent = `${total}/${total}`;
    setTimeout(() => {
      overlay.classList.remove('show');
    }, 500);
  } else {
    progress.textContent = `${current}/${total}`;
  }
}

export function handleScanProgress(current: number, total: number) {
  const overlay = document.getElementById('loadingOverlay');
  const progress = document.getElementById('loadingProgress');
  const loader = document.getElementById('loader');
  if (!overlay || !progress) return;

  setOverlayTitle('Scanning...');
  if (loader) loader.style.display = 'none';
  overlay.classList.add('show');

  // total < 0 → still collecting layers; count not known yet
  if (total < 0) {
    progress.textContent = 'Preparing…';
    return;
  }

  // Empty page — keep overlay until scan-result clears it
  if (total === 0) {
    progress.textContent = '0/0';
    return;
  }

  // Always show real counts (e.g. 12/341). Do not hide here — scan-result does.
  progress.textContent = `${current}/${total}`;
}

export function showLoader(message: string) {
  const loader = document.getElementById('loader');
  if (!loader) return;
  loader.textContent = message || 'Loading...';
  loader.style.display = 'block';
}
