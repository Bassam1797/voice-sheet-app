// Utility helpers

export const letters = (n: number): string => {
  let s = '';
  n++;
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
};

export const clamp = (v: number, lo: number, hi: number): number =>
  Math.max(lo, Math.min(hi, v));

export const copyToClipboard = async (text: string): Promise<void> => {
  try {
    await navigator.clipboard.writeText(text);
  } catch (e) {
    console.warn('Clipboard write failed', e);
  }
};

export const readFromClipboard = async (): Promise<string> => {
  try {
    return await navigator.clipboard.readText();
  } catch (e) {
    console.warn('Clipboard read failed', e);
    return '';
  }
};

export const parseTSV = (tsv: string): string[][] =>
  tsv.replace(/\r/g, '').split('\n').map(r => r.split('\t'));

export const toTSV = (rows: string[][]): string =>
  rows.map(r => r.map(c => c ?? '').join('\t')).join('\n');
