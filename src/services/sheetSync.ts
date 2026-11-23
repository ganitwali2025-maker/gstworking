export const APPS_SCRIPT_URL = 'https://script.google.com/macros/s/AKfycby-lHw-wHVBXrS0SsuJ0iBeB96Ve_Zd_CGx3KvUJZQ7lvO32QSrkREr1V7xC_E9eK3X/exec';
export const APPS_SCRIPT_SECRET = 'MY_GST_SECRET_KEY_123';

export interface SheetSyncPayload {
  action: 'upsert' | 'delete' | 'clear' | 'lock' | 'unlock' | 'read';
  sheet: string;
  headers?: string[];
  rows?: { id: string; values: (string | number)[] }[];
  ids?: string[];
}

export async function syncModuleToSheets(
  payload: SheetSyncPayload
): Promise<{ success: boolean; count?: number; error?: string | null }> {
  // 1. Try local backend proxy
  try {
    const res = await fetch('/api/sheets/sync', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload)
    });

    if (res.ok) {
      const data = await res.json();
      if (data.success) {
        return data;
      }
    }
  } catch (_) {}

  // 2. Direct fallback to Google Apps Script Web App (vital on Vercel, uses text/plain to avoid browser CORS preflight)
  try {
    const directRes = await fetch(APPS_SCRIPT_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify({
        secret: APPS_SCRIPT_SECRET,
        ...payload
      })
    });

    const directJson = await directRes.json();
    return {
      success: Boolean(directJson.ok),
      count: directJson.count || 0,
      error: directJson.error || null
    };
  } catch (err: unknown) {
    const msg = err instanceof Error ? err.message : 'Network error';
    return { success: false, error: msg };
  }
}

export async function readModuleFromSheets(
  sheet: string
): Promise<{ success: boolean; rows: any[]; count: number; error?: string | null }> {
  // 1. Try local backend proxy
  try {
    const res = await fetch('/api/sheets/sync', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ action: 'read', sheet })
    });

    if (res.ok) {
      const data = await res.json();
      if (data.success && Array.isArray(data.rows)) {
        return { success: true, rows: data.rows, count: data.count || data.rows.length };
      }
    }
  } catch (_) {}

  // 2. Direct fallback to Google Apps Script Web App (vital on Vercel, uses text/plain to avoid browser CORS preflight)
  try {
    const directRes = await fetch(APPS_SCRIPT_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify({
        secret: APPS_SCRIPT_SECRET,
        action: 'read',
        sheet
      })
    });

    const directJson = await directRes.json();
    if (directJson.ok && Array.isArray(directJson.rows)) {
      return { success: true, rows: directJson.rows, count: directJson.rows.length };
    }
  } catch (err: unknown) {
    const msg = err instanceof Error ? err.message : 'Network error';
    return { success: false, rows: [], count: 0, error: msg };
  }

  return { success: false, rows: [], count: 0 };
}
