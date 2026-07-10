import {
  familyStylesGlobal,
  lastScanScope,
  beginSystemFontsRequest,
  postToPlugin
} from '../state';
import { showLoader, showNotification } from '../notifications';
import { splitWeightGroups } from '../weights';

type ReplaceModalOptions = {
  title?: string;
  currentLabel?: string;
  placeholder?: string;
  inputType?: string;
  loadingText?: string;
  suggestions?: string[] | (() => string[] | Promise<string[]>) | null;
  onSubmit?: (value: string) => boolean | void;
};

function attachSuggestions(
  inputWrapper: HTMLElement,
  input: HTMLInputElement,
  suggestionSource: string[]
) {
  if (!suggestionSource.length) return;
  const suggestionsBox = document.createElement('div');
  suggestionsBox.className = 'suggestions-list';

  const renderSuggestions = (query: string) => {
    suggestionsBox.innerHTML = '';
    const val = (query || '').toLowerCase();
    const matches = val
      ? suggestionSource.filter(f => f.toLowerCase().includes(val))
      : suggestionSource;

    if (matches.length > 0) {
      suggestionsBox.classList.add('show');
      matches.forEach(match => {
        const item = document.createElement('div');
        item.className = 'suggestion-item';
        item.textContent = match;
        item.onclick = () => {
          input.value = match;
          suggestionsBox.classList.remove('show');
        };
        suggestionsBox.appendChild(item);
      });
    } else {
      suggestionsBox.classList.remove('show');
    }
  };

  input.addEventListener('focus', () => renderSuggestions(input.value));
  input.addEventListener('input', () => renderSuggestions(input.value));

  const closeSuggestions = (e: MouseEvent) => {
    if (!inputWrapper.contains(e.target as Node)) {
      suggestionsBox.classList.remove('show');
      document.removeEventListener('click', closeSuggestions);
    }
  };
  document.addEventListener('click', closeSuggestions);
  inputWrapper.appendChild(suggestionsBox);
}

export async function openValueReplaceModal(options: ReplaceModalOptions) {
  const {
    title,
    currentLabel,
    placeholder,
    inputType = 'text',
    loadingText = 'Updating...',
    suggestions = null,
    onSubmit
  } = options;

  const modal = document.getElementById('replaceModal')!;
  const modalTitle = document.getElementById('modalTitle')!;
  const modalContent = document.getElementById('modalContent')!;

  modalTitle.textContent = title || 'Replace';
  modalContent.innerHTML = '';

  const inputRow = document.createElement('div');
  inputRow.className = 'replace-row';

  const beforeGroup = document.createElement('div');
  beforeGroup.className = 'input-group';

  const beforeInput = document.createElement('input');
  beforeInput.className = 'replace-input readonly';
  beforeInput.value = currentLabel || '';
  beforeInput.readOnly = true;

  const dividerLabel = document.createElement('span');
  dividerLabel.className = 'replace-divider-label';
  dividerLabel.textContent = 'Replace with';

  beforeGroup.appendChild(beforeInput);
  beforeGroup.appendChild(dividerLabel);

  const afterGroup = document.createElement('div');
  afterGroup.className = 'input-group';

  const inputWrapper = document.createElement('div');
  inputWrapper.className = 'input-wrapper';

  const input = document.createElement('input');
  input.className = 'replace-input';
  input.placeholder = placeholder || 'Enter value...';
  input.autocomplete = 'off';
  input.type = inputType;

  inputWrapper.appendChild(input);

  let suggestionSource: string[] = [];
  if (suggestions) {
    const raw = typeof suggestions === 'function' ? suggestions() : suggestions;
    suggestionSource = Array.isArray(raw) ? raw : await raw;
  }
  attachSuggestions(inputWrapper, input, suggestionSource);

  afterGroup.appendChild(inputWrapper);

  const applyBtn = document.createElement('button');
  applyBtn.className = 'apply-btn';
  applyBtn.textContent = 'Apply';
  applyBtn.onclick = () => {
    const value = input.value.trim();
    if (!value) {
      input.focus();
      return;
    }
    const shouldClose = onSubmit ? onSubmit(value) : true;
    if (shouldClose === false) {
      return;
    }
    closeReplaceModal();
    if (loadingText) {
      showLoader(loadingText);
    }
  };

  inputRow.appendChild(beforeGroup);
  inputRow.appendChild(afterGroup);
  inputRow.appendChild(applyBtn);
  modalContent.appendChild(inputRow);

  modal.classList.add('show');
  setTimeout(() => input.focus(), 50);
}

export async function showReplaceModal(fontName: string) {
  await openValueReplaceModal({
    title: 'Replace Font',
    currentLabel: fontName,
    placeholder: 'Type font name...',
    loadingText: 'Replacing...',
    suggestions: () => beginSystemFontsRequest(),
    onSubmit: (newFont) => {
      if (newFont.toLowerCase() === fontName.toLowerCase()) {
        alert('New font is the same as current font.');
        return false;
      }
      postToPlugin({
        type: 'replace-font',
        oldFont: fontName,
        newFont,
        scope: lastScanScope
      });
      return true;
    }
  });
}


export function showReplaceWeightModal(fontName, weightValue) {
  const availableStyles = familyStylesGlobal[fontName];
  if (!availableStyles || availableStyles.length === 0) {
    openValueReplaceModal({
      title: `Replace "${fontName}" weight`,
      currentLabel: weightValue,
      placeholder: 'e.g. Regular, Medium, Bold',
      loadingText: 'Updating weight...',
      onSubmit: (value) => {
        postToPlugin({
            type: 'replace-font-weight',
            family: fontName,
            oldStyle: weightValue,
            newStyle: value,
            scope: lastScanScope
          });
        return true;
      }
    });
    return;
  }

  const modal = document.getElementById('replaceModal')!;
  const modalTitle = document.getElementById('modalTitle')!;
  const modalContent = document.getElementById('modalContent')!;

  modalTitle.textContent = `Choose a weight for ${fontName}`;
  modalContent.innerHTML = '';

  const info = document.createElement('p');
  info.style.fontSize = '13px';
  info.style.color = '#555';
  info.style.margin = '0 0 12px 0';
  info.textContent = 'Select an available style to replace the current weight.';
  modalContent.appendChild(info);

  const groups = splitWeightGroups(availableStyles, value => value);
  const currentLower = String(weightValue || '').toLowerCase();

  const renderGroup = (title, entries) => {
    if (entries.length === 0) return;
    const groupWrapper = document.createElement('div');
    groupWrapper.style.marginBottom = '16px';

    const subtitle = document.createElement('div');
    subtitle.className = 'detail-subtitle';
    subtitle.textContent = title;
    groupWrapper.appendChild(subtitle);

    const buttonRow = document.createElement('div');
    buttonRow.className = 'weight-picker';

    entries.forEach(entry => {
      const style = entry.value;
      const btn = document.createElement('button');
      btn.className = 'weight-picker-btn';
      if (style.toLowerCase() === currentLower) {
        btn.classList.add('current');
        btn.disabled = true;
      }
      btn.textContent = style;
      btn.onclick = () => {
        if (btn.disabled) return;
        closeReplaceModal();
        showLoader('Updating weight...');
        postToPlugin({
            type: 'replace-font-weight',
            family: fontName,
            oldStyle: weightValue,
            newStyle: style,
            scope: lastScanScope
          });
      };
      buttonRow.appendChild(btn);
    });

    groupWrapper.appendChild(buttonRow);
    modalContent.appendChild(groupWrapper);
  };

  renderGroup('Normal', groups.normal);
  renderGroup('Italic', groups.italic);

  modal.classList.add('show');
}

export function showReplaceSizeModal(fontName, sizeValue) {
  const modal = document.getElementById('replaceModal')!;
  const modalTitle = document.getElementById('modalTitle')!;
  const modalContent = document.getElementById('modalContent')!;
  
  // Format size for display
  const formattedSize = typeof sizeValue === 'number' ? 
    (Number.isInteger(sizeValue) ? sizeValue : parseFloat(sizeValue.toFixed(2))) : 
    sizeValue;
  
  modalTitle.textContent = 'Change Font Size';
  modalContent.innerHTML = '';
  
  // Create input row
  const inputRow = document.createElement('div');
  inputRow.className = 'replace-row';
  
  // Current size display
  const beforeGroup = document.createElement('div');
  beforeGroup.className = 'input-group';
  
  const beforeInput = document.createElement('input');
  beforeInput.className = 'replace-input readonly';
  beforeInput.value = `${formattedSize}px`;
  beforeInput.readOnly = true;
  
  const dividerLabel = document.createElement('span');
  dividerLabel.className = 'replace-divider-label';
  dividerLabel.textContent = 'Change to';
  
  beforeGroup.appendChild(beforeInput);
  beforeGroup.appendChild(dividerLabel);
  
  // New size input
  const afterGroup = document.createElement('div');
  afterGroup.className = 'input-group';
  
  const input = document.createElement('input');
  input.className = 'replace-input';
  input.type = 'number';
  input.placeholder = 'Enter new size';
  input.step = '0.01';
  input.min = '1';
  input.autocomplete = 'off';
  
  afterGroup.appendChild(input);
  
  // Apply button
  const applyBtn = document.createElement('button');
  applyBtn.className = 'apply-btn';
  applyBtn.textContent = 'Apply';
  applyBtn.onclick = () => {
    const newSize = parseFloat(input.value);
    if (isNaN(newSize) || newSize <= 0) {
      alert('Please enter a size greater than 0.');
      input.focus();
      return;
    }
    if (newSize === sizeValue) {
      alert('New size is the same as current size.');
      input.focus();
      return;
    }
    
    closeReplaceModal();
    showNotification(`Updating ${fontName} ${formattedSize}px → ${newSize}px...`, 0);
    
    postToPlugin({
        type: 'replace-font-size',
        family: fontName,
        oldSize: sizeValue,
        newSize: newSize,
        scope: lastScanScope
      });
  };

  // Allow Enter key to submit
  input.addEventListener('keypress', (e) => {
    if (e.key === 'Enter') {
      applyBtn.click();
    }
  });
  
  inputRow.appendChild(beforeGroup);
  inputRow.appendChild(afterGroup);
  inputRow.appendChild(applyBtn);
  modalContent.appendChild(inputRow);
  
  // Show modal and focus input
  modal.classList.add('show');
  setTimeout(() => input.focus(), 100);
}

export function closeReplaceModal() {
  const modal = document.getElementById('replaceModal')!;
  modal.classList.remove('show');
}