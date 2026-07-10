const WEIGHT_KEYWORDS = [
  { weight: 100, tokens: ['thin', 'hairline'] },
  { weight: 200, tokens: ['extra light', 'extralight', 'ultra light', 'ultralight'] },
  { weight: 300, tokens: ['light'] },
  { weight: 350, tokens: ['book'] },
  { weight: 400, tokens: ['regular', 'normal', 'roman'] },
  { weight: 500, tokens: ['medium'] },
  { weight: 600, tokens: ['semibold', 'semi bold', 'demibold', 'demi bold'] },
  { weight: 700, tokens: ['bold'] },
  { weight: 800, tokens: ['extra bold', 'extrabold', 'ultra bold', 'ultrabold'] },
  { weight: 900, tokens: ['black', 'heavy'] }
];

export function getWeightMeta(style: string | undefined) {
  const value = (style || '').toString();
  const lower = value.toLowerCase();
  const isItalic = lower.includes('italic') || lower.includes('oblique');
  const numericMatch = lower.match(/(\d{3})/);
  let weight = numericMatch ? parseInt(numericMatch[1], 10) : null;
  if (!weight) {
    for (const entry of WEIGHT_KEYWORDS) {
      if (entry.tokens.some(token => lower.includes(token))) {
        weight = entry.weight;
        break;
      }
    }
  }
  return {
    weight: weight ?? 400,
    isItalic
  };
}

export function splitWeightGroups<T>(
  items: T[],
  accessor: (item: T) => string = item => item as unknown as string
) {
  const enriched = items.map(item => {
    const value = accessor(item);
    return {
      item,
      value,
      meta: getWeightMeta(value)
    };
  });

  enriched.sort((a, b) => {
    if (a.meta.weight !== b.meta.weight) {
      return a.meta.weight - b.meta.weight;
    }
    if (a.meta.isItalic !== b.meta.isItalic) {
      return Number(a.meta.isItalic) - Number(b.meta.isItalic);
    }
    return a.value.localeCompare(b.value);
  });

  return {
    normal: enriched.filter(entry => !entry.meta.isItalic),
    italic: enriched.filter(entry => entry.meta.isItalic)
  };
}
