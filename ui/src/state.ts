/** Shared UI state for Font Scanner */

export type FontDetailsMap = Record<
  string,
  {
    styles: { value: string; count: number }[];
    sizeVariants: {
      fontSize: number;
      lineHeightPercent: number | null;
      letterSpacingPercent: number;
      count: number;
    }[];
  }
>;

export let systemFontsGlobal: string[] = [];
export let systemFontsPromise: Promise<string[]> | null = null;
let systemFontsResolve: ((fonts: string[]) => void) | null = null;

export let selectedFontItem: HTMLElement | null = null;
export let selectedFontName: string | null = null;
export let fontDetailsGlobal: FontDetailsMap = {};
export let familyStylesGlobal: Record<string, string[]> = {};
export let selectedFrameCount = 0;
/** 'page' or 'selection' — used when applying replace so only selection is updated */
export let lastScanScope: 'page' | 'selection' = 'page';

export function postToPlugin(message: Record<string, unknown>) {
  parent.postMessage({ pluginMessage: message }, '*');
}

export function setSystemFonts(fonts: string[]) {
  systemFontsGlobal = fonts;
  if (systemFontsResolve) {
    systemFontsResolve(fonts);
    systemFontsResolve = null;
  }
  systemFontsPromise = null;
}

export function beginSystemFontsRequest(): Promise<string[]> {
  if (systemFontsGlobal.length > 0) {
    return Promise.resolve(systemFontsGlobal);
  }
  if (systemFontsPromise) return systemFontsPromise;
  const p = new Promise<string[]>(resolve => {
    systemFontsResolve = resolve;
  });
  systemFontsPromise = p;
  postToPlugin({ type: 'get-system-fonts' });
  // Timeout fallback so the modal is never stuck
  setTimeout(() => {
    if (systemFontsResolve) {
      systemFontsResolve(systemFontsGlobal);
      systemFontsResolve = null;
      systemFontsPromise = null;
    }
  }, 8000);
  return p;
}

export function setSelectedFontItem(item: HTMLElement | null) {
  selectedFontItem = item;
}

export function setSelectedFontName(name: string | null) {
  selectedFontName = name;
}

export function setFontDetails(details: FontDetailsMap) {
  fontDetailsGlobal = details;
}

export function setFamilyStyles(styles: Record<string, string[]>) {
  familyStylesGlobal = styles;
}

export function setSelectedFrameCount(count: number) {
  selectedFrameCount = count;
}

export function setLastScanScope(scope: 'page' | 'selection') {
  lastScanScope = scope;
}
