export const round2 = (n) => Math.round((Number(n) + Number.EPSILON) * 100) / 100;

export const num = (v) => {
  if (v === null || v === undefined || v === '') return null;
  if (typeof v === 'number') return v;
  const n = Number(String(v).replace(',', '.').trim());
  return Number.isFinite(n) ? n : null;
};

export const text = (v) => (v === null || v === undefined ? '' : String(v).trim());

export const attendanceValue = (v) => {
  if (typeof v === 'number') return v;
  const m = text(v).match(/-?\d+(?:[.,]\d+)?/);
  return m ? Number(m[0].replace(',', '.')) : 0;
};

export const fmt = (v) => {
  if (v === null || v === undefined || v === '') return '';
  if (typeof v === 'number') return Number.isInteger(v) ? String(v) : String(round2(v));
  return String(v);
};

export const safeName = (s) =>
  String(s || 'hoc-sinh')
    .normalize('NFD')
    .replace(/[̀-ͯ]/g, '')
    .replace(/đ/g, 'd')
    .replace(/Đ/g, 'D')
    .replace(/[^a-zA-Z0-9-_ ]/g, '')
    .trim()
    .replace(/\s+/g, '_');
