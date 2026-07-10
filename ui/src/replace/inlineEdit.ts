/** Shared helpers for inline size / line-height / letter-spacing edits in details rows. */

export function formatSizeDisplay(size: number): string {
  return Number.isInteger(size) ? String(size) : parseFloat(size.toFixed(2)).toString();
}

export function parsePositiveNumber(raw: string): number | null {
  const n = parseFloat(raw);
  if (isNaN(n) || n <= 0) return null;
  return n;
}

export function parseLineHeightInput(raw: string): number | null | 'invalid' {
  const trimmed = raw.trim();
  if (/^auto$/i.test(trimmed)) return null;
  const n = parseFloat(trimmed);
  if (isNaN(n) || n < 0) return 'invalid';
  return n;
}
