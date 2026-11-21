export const MN = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
export const FYS = ['FY 2026–27', 'FY 2025–26'];

export const QM: Record<string, string[]> = {
  Q1: ['Apr', 'May', 'Jun'],
  Q2: ['Jul', 'Aug', 'Sep'],
  Q3: ['Oct', 'Nov', 'Dec'],
  Q4: ['Jan', 'Feb', 'Mar']
};

export const inr = (n: number): string => '₹' + Math.round(n || 0).toLocaleString('en-IN');

export const fm = (p: number): string => {
  const num = Number(p) || 0;
  const a = Math.abs(num);
  const rupees = Math.floor(a / 100);
  const paise = String(a % 100).padStart(2, '0');
  return (num < 0 ? '-' : '') + '₹' + rupees.toLocaleString('en-IN') + '.' + paise;
};

export const fmtVal = (p: number): string => {
  const num = (Number(p) || 0) / 100;
  if (num === 0) return '0';
  if (Number.isInteger(num)) {
    return num.toLocaleString('en-IN');
  }
  return num.toLocaleString('en-IN', { minimumFractionDigits: 1, maximumFractionDigits: 2 });
};

export const pp = (v: any): number => {
  const t = String(v == null ? '' : v).replace(/[₹,\s]/g, '');
  if (t === '') return 0;
  const m = t.match(/^(\d*)\.?(\d*)$/);
  if (!m || (m[1] === '' && m[2] === '')) return NaN;
  let r = Number(m[1] || 0) * 100 + Number((m[2] + '00').slice(0, 2));
  if (m[2].length > 2 && +m[2][2] >= 5) r++;
  return r;
};

export const qOf = (mi: number): string => (mi >= 3 && mi <= 5 ? 'Q1' : mi >= 6 && mi <= 8 ? 'Q2' : mi >= 9 && mi <= 11 ? 'Q3' : 'Q4');

export const fyOf = (y: number, mi: number): string => {
  const a = mi >= 3 ? y : y - 1;
  return 'FY ' + a + '–' + String(a + 1).slice(2);
};

export const fyList = (fy: string) => {
  const y = +fy.slice(3, 7);
  return [3, 4, 5, 6, 7, 8, 9, 10, 11, 0, 1, 2].map(mi => ({
    l: MN[mi] + ' ' + (mi >= 3 ? y : y + 1),
    q: qOf(mi)
  }));
};

export const dmy = (d: string): string => (d ? d.split('-').reverse().join('-') : '');
export const dmySlash = (d: string): string => (d ? d.split('-').reverse().join('/') : '');

export const iso = (v: any): string => {
  if (v instanceof Date && !isNaN(v.getTime())) {
    const t = new Date(v.getTime() + 12 * 36e5);
    return t.getFullYear() + '-' + String(t.getMonth() + 1).padStart(2, '0') + '-' + String(t.getDate()).padStart(2, '0');
  }
  const t = String(v || '').trim();
  let m = t.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (m) return m[1] + '-' + m[2].padStart(2, '0') + '-' + m[3].padStart(2, '0');
  m = t.match(/^(\d{1,2})[-\/.](\d{1,2})[-\/.](\d{2,4})$/);
  if (m) return (m[3].length === 2 ? '20' + m[3] : m[3]) + '-' + m[2].padStart(2, '0') + '-' + m[1].padStart(2, '0');
  return '';
};
