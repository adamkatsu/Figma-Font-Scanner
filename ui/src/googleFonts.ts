import { GOOGLE_FONT_FAMILY_SET } from './data/googleFontFamilies';

export const ICON_KEYWORDS = [
  'awesome',
  'icon',
  'symbol',
  'glyph',
  'icomoon',
  'entypo',
  'dingbat',
  'ionicons',
  'material icons',
  'material symbols'
];

export function isGoogleFont(family: string): boolean {
  return GOOGLE_FONT_FAMILY_SET.has(family.toLowerCase());
}

export function isIconFont(family: string): boolean {
  const lower = family.toLowerCase();
  return ICON_KEYWORDS.some(keyword => lower.includes(keyword));
}

export function getGoogleFontSet(): ReadonlySet<string> {
  return GOOGLE_FONT_FAMILY_SET;
}
