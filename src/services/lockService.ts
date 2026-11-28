/**
 * Month-wise and Period Lock Service
 * Provides single source of truth for locking/unlocking months across all sheets & dropdowns.
 */

export const GST_LOCKS_STORAGE_KEY = 'gst_reco_locks_v1';
export const GST_LOCK_CHANGED_EVENT = 'gst-lock-changed';

export interface MonthLockData {
  isLocked: boolean;
  lockedAt?: string;
  lockedBy?: string;
}

export function parseWorkingMonth(monthStr?: string, fyStr?: string): string {
  let yr = '2026';
  if (fyStr) {
    const m = fyStr.replace(/[^0-9]/g, '').slice(0, 4);
    if (m) yr = m;
  }
  let m = (monthStr || '').trim();
  if (!m || m.toUpperCase() === 'ALL') {
    m = 'Oct';
  }

  const MONTH_MAP: Record<string, string> = {
    jan: 'Jan', january: 'Jan', feb: 'Feb', february: 'Feb',
    mar: 'Mar', march: 'Mar', apr: 'Apr', april: 'Apr',
    may: 'May', jun: 'Jun', june: 'Jun', jul: 'Jul', july: 'Jul',
    aug: 'Aug', august: 'Aug', sep: 'Sep', sept: 'Sep', september: 'Sep',
    oct: 'Oct', october: 'Oct', nov: 'Nov', november: 'Nov', dec: 'Dec'
  };

  const stdMonth = MONTH_MAP[m.toLowerCase()] || (m ? m.charAt(0).toUpperCase() + m.slice(1, 3) : 'Oct');
  const startYr = parseInt(yr, 10) || 2026;
  const endYr = startYr + 1;
  const finalYr = ['Jan', 'Feb', 'Mar'].includes(stdMonth) ? endYr : startYr;
  return `${stdMonth}-${finalYr}`;
}

export function getMonthLockInfo(workingMonth: string): MonthLockData {
  try {
    const raw = localStorage.getItem(GST_LOCKS_STORAGE_KEY);
    if (raw) {
      const parsed = JSON.parse(raw);
      if (parsed && parsed[workingMonth]?.isLocked) {
        return parsed[workingMonth];
      }
    }
  } catch {}
  return { isLocked: false };
}

export function isWorkingMonthLocked(workingMonth: string): boolean {
  return getMonthLockInfo(workingMonth).isLocked;
}

export function isMonthLocked(monthCode: string, fy?: string): boolean {
  const workingMonth = parseWorkingMonth(monthCode, fy);
  return isWorkingMonthLocked(workingMonth);
}

export function setWorkingMonthLock(workingMonth: string, locked: boolean, user: string = 'Admin'): void {
  try {
    const raw = localStorage.getItem(GST_LOCKS_STORAGE_KEY);
    const parsed = raw ? JSON.parse(raw) : {};
    if (locked) {
      parsed[workingMonth] = {
        isLocked: true,
        lockedAt: new Date().toISOString(),
        lockedBy: user
      };
    } else {
      parsed[workingMonth] = {
        isLocked: false,
        lockedAt: '',
        lockedBy: ''
      };
    }
    localStorage.setItem(GST_LOCKS_STORAGE_KEY, JSON.stringify(parsed));
    window.dispatchEvent(new CustomEvent(GST_LOCK_CHANGED_EVENT, { detail: { workingMonth, locked } }));
  } catch (e) {
    console.error('Failed to set month lock:', e);
  }
}

export function getAllLockedMonths(): string[] {
  try {
    const raw = localStorage.getItem(GST_LOCKS_STORAGE_KEY);
    if (raw) {
      const parsed = JSON.parse(raw);
      return Object.keys(parsed).filter((k) => parsed[k]?.isLocked);
    }
  } catch {}
  return [];
}
