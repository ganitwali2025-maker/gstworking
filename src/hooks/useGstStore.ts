import { useState, useEffect, useCallback } from 'react';
import { ViewType, InvoiceRecord, StepItem, BusinessSettings } from '../types/gst';
import { SEED_MONTHLY_DATA, SEED_STEPS, SEED_SETTINGS, SEED_INVOICES, SEED_RCM } from '../data/seedData';
import { FYS } from '../utils/gstFormat';
import { getViewFromPath, navigateTo, getQueryParams } from '../utils/router';
import { syncModuleToSheets, readModuleFromSheets } from '../services/sheetSync';
import { GSTR1_SHEET_HEADERS, mapRowToSheetRow, mapSheetRowToAppRecord } from '../services/registerService';

const LOCAL_STORAGE_KEY_INVOICES = 'gstr1_register_v4';
const LOCAL_STORAGE_KEY_LOCK = 'gstr1_lock_v1';
const LOCAL_STORAGE_KEY_SETTINGS = 'gst_settings_v1';
const LOCAL_STORAGE_KEY_STEPS = 'gst_steps_v1';
const LOCAL_STORAGE_KEY_LOGGED_OUT = 'gst_is_logged_out';

export function useGstStore() {
  const [view, setViewState] = useState<ViewType>(() => {
    if (typeof window !== 'undefined') {
      const matched = getViewFromPath(window.location.pathname);
      if (matched) return matched;
    }
    return 'itc';
  });

  const [fy, setFyState] = useState<string>(() => {
    if (typeof window !== 'undefined') {
      const q = getQueryParams();
      if (q.fy) return q.fy;
      const saved = localStorage.getItem('gst_active_fy');
      if (saved) return saved;
    }
    return 'FY 2026–27';
  });

  const [quarter, setQuarter] = useState<string>('');

  const [month, setMonthState] = useState<string>(() => {
    if (typeof window !== 'undefined') {
      const q = getQueryParams();
      if (q.month) return q.month;
      const saved = localStorage.getItem('gst_active_month');
      if (saved) return saved;
    }
    return 'Oct';
  });

  const setView = useCallback((nextView: ViewType) => {
    setViewState(nextView);
    navigateTo(nextView);
  }, []);

  const setFy = useCallback((newFy: string) => {
    setFyState(newFy);
    try {
      localStorage.setItem('gst_active_fy', newFy);
    } catch {}
  }, []);

  const setMonth = useCallback((newMonth: string) => {
    setMonthState(newMonth);
    try {
      localStorage.setItem('gst_active_month', newMonth);
    } catch {}
  }, []);

  // Listen to browser Back/Forward (popstate) and ensure valid canonical route on initial load
  useEffect(() => {
    if (typeof window === 'undefined') return;

    const currentPath = window.location.pathname;
    const matched = getViewFromPath(currentPath);
    if (!matched || currentPath === '/') {
      navigateTo(matched || 'itc', undefined, undefined, true);
    }

    const handlePopState = () => {
      const currentMatched = getViewFromPath(window.location.pathname);
      if (currentMatched) {
        setViewState(currentMatched);
      } else {
        setViewState('itc');
        navigateTo('itc', undefined, undefined, true);
      }
      const q = getQueryParams();
      if (q.fy) setFyState(q.fy);
      else {
        const saved = localStorage.getItem('gst_active_fy');
        if (saved) setFyState(saved);
      }
      if (q.month) setMonthState(q.month);
      else {
        const saved = localStorage.getItem('gst_active_month');
        if (saved) setMonthState(saved);
      }
    };

    window.addEventListener('popstate', handlePopState);
    return () => window.removeEventListener('popstate', handlePopState);
  }, []);

  // Dashboard Data
  const [monthlyData] = useState(SEED_MONTHLY_DATA);
  const [rcmList] = useState(SEED_RCM);

  // Invoices & Lock
  const [invoices, setInvoices] = useState<InvoiceRecord[]>(() => {
    try {
      const saved = localStorage.getItem(LOCAL_STORAGE_KEY_INVOICES);
      if (saved) {
        const parsed = JSON.parse(saved);
        if (Array.isArray(parsed) && parsed.length > 0) return parsed;
      }
      return SEED_INVOICES;
    } catch {
      return SEED_INVOICES;
    }
  });

  const [isLocked, setIsLocked] = useState<boolean>(() => {
    try {
      return localStorage.getItem(LOCAL_STORAGE_KEY_LOCK) === '1';
    } catch {
      return false;
    }
  });

  const [selectedInvoiceId, setSelectedInvoiceId] = useState<string | null>(null);
  const [pastInvoices, setPastInvoices] = useState<InvoiceRecord[][]>([]);
  const [futureInvoices, setFutureInvoices] = useState<InvoiceRecord[][]>([]);

  const pushInvoiceHistory = (prev: InvoiceRecord[]) => {
    setPastInvoices(p => [...p.slice(-25), prev]);
    setFutureInvoices([]);
  };

  const undoInvoices = () => {
    if (pastInvoices.length === 0) {
      showToast('No more actions to undo');
      return;
    }
    const previous = pastInvoices[pastInvoices.length - 1];
    const newPast = pastInvoices.slice(0, pastInvoices.length - 1);
    setFutureInvoices(f => [invoices, ...f]);
    setPastInvoices(newPast);
    setInvoices(previous);
    persistInvoices(previous);
    showToast('Undo successful');
  };

  const redoInvoices = () => {
    if (futureInvoices.length === 0) {
      showToast('No more actions to redo');
      return;
    }
    const next = futureInvoices[0];
    const newFuture = futureInvoices.slice(1);
    setPastInvoices(p => [...p, invoices]);
    setFutureInvoices(newFuture);
    setInvoices(next);
    persistInvoices(next);
    showToast('Redo successful');
  };

  // Steps
  const [steps, setSteps] = useState<StepItem[]>(() => {
    try {
      const saved = localStorage.getItem(LOCAL_STORAGE_KEY_STEPS);
      return saved ? JSON.parse(saved) : SEED_STEPS;
    } catch {
      return SEED_STEPS;
    }
  });

  // Settings
  const [settings, setSettings] = useState<BusinessSettings>(() => {
    try {
      const saved = localStorage.getItem(LOCAL_STORAGE_KEY_SETTINGS);
      return saved ? JSON.parse(saved) : SEED_SETTINGS;
    } catch {
      return SEED_SETTINGS;
    }
  });

  // UI
  const [toastMessage, setToastMessage] = useState<string | null>(null);
  const [isLoggedOut, setIsLoggedOutState] = useState<boolean>(() => {
    try {
      return localStorage.getItem(LOCAL_STORAGE_KEY_LOGGED_OUT) === '1';
    } catch {
      return false;
    }
  });

  const setIsLoggedOut = useCallback((val: boolean) => {
    setIsLoggedOutState(val);
    try {
      if (val) {
        localStorage.setItem(LOCAL_STORAGE_KEY_LOGGED_OUT, '1');
      } else {
        localStorage.removeItem(LOCAL_STORAGE_KEY_LOGGED_OUT);
      }
    } catch {}
  }, []);

  const showToast = useCallback((msg: string) => {
    setToastMessage(msg);
    setTimeout(() => {
      setToastMessage((prev) => (prev === msg ? null : prev));
    }, 3000);
  }, []);

  // Save changes locally
  const persistInvoices = (newInvoices: InvoiceRecord[], locked = isLocked) => {
    try {
      localStorage.setItem(LOCAL_STORAGE_KEY_INVOICES, JSON.stringify(newInvoices));
      localStorage.setItem(LOCAL_STORAGE_KEY_LOCK, locked ? '1' : '0');
    } catch (e) {
      console.warn('LocalStorage save notice:', e);
    }
  };

  const persistSteps = (newSteps: StepItem[]) => {
    try {
      localStorage.setItem(LOCAL_STORAGE_KEY_STEPS, JSON.stringify(newSteps));
    } catch (e) {}
  };

  const persistSettings = (newSettings: BusinessSettings) => {
    try {
      localStorage.setItem(LOCAL_STORAGE_KEY_SETTINGS, JSON.stringify(newSettings));
    } catch (e) {}
  };

  // Sync with API / Google Sheets when available
  const fetchApiInvoices = async () => {
    const currentLock = (() => {
      try {
        return localStorage.getItem(LOCAL_STORAGE_KEY_LOCK) === '1';
      } catch {
        return isLocked;
      }
    })();

    try {
      const res = await fetch('/api/invoices');
      if (res.ok) {
        const json = await res.json();
        if (json.success && Array.isArray(json.data)) {
          setInvoices(json.data);
          const effectiveLock = typeof json.locked === 'boolean' ? json.locked : currentLock;
          persistInvoices(json.data, effectiveLock);
          setIsLocked(effectiveLock);
          return;
        }
      }
    } catch (e) {}

    // Direct fallback to Google Sheets Web App (crucial on Vercel)
    try {
      const sheetRes = await readModuleFromSheets('GSTR-1');
      if (sheetRes.success && Array.isArray(sheetRes.rows)) {
        const mapped = sheetRes.rows.map((r: any) => mapSheetRowToAppRecord(r.values || r, 'GSTR-1')).filter(Boolean);
        setInvoices(mapped);
        persistInvoices(mapped, currentLock);
        setIsLocked(currentLock);
        return;
      }
    } catch (_) {}
  };

  const fetchApiSteps = async () => {
    try {
      const res = await fetch('/api/steps');
      if (res.ok) {
        const json = await res.json();
        if (json.success && Array.isArray(json.steps)) {
          setSteps(json.steps);
          persistSteps(json.steps);
        }
      }
    } catch (e) {}
  };

  const fetchApiSettings = async () => {
    try {
      const res = await fetch('/api/settings');
      if (res.ok) {
        const json = await res.json();
        if (json.success && json.settings) {
          setSettings(json.settings);
          persistSettings(json.settings);
        }
      }
    } catch (e) {}
  };

  useEffect(() => {
    fetchApiInvoices();
    fetchApiSteps();
    fetchApiSettings();

    const onWindowFocus = () => {
      fetchApiInvoices();
    };
    window.addEventListener('focus', onWindowFocus);

    const handleSynced = (e: any) => {
      if (e.detail?.tab === 'GSTR-1') {
        try {
          const saved = localStorage.getItem(LOCAL_STORAGE_KEY_INVOICES);
          if (saved) setInvoices(JSON.parse(saved));
        } catch {}
      }
    };
    window.addEventListener('register-synced', handleSynced);
    return () => {
      window.removeEventListener('focus', onWindowFocus);
      window.removeEventListener('register-synced', handleSynced);
    };
  }, [fy]);

  // Invoice Actions
  const addInvoice = async (record: InvoiceRecord) => {
    pushInvoiceHistory(invoices);
    const updated = [record, ...invoices];
    setInvoices(updated);
    persistInvoices(updated);
    try {
      const res = await fetch('/api/invoices', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(record)
      });
      if (res.ok) {
        const json = await res.json();
        if (json.sheetWarning) {
          showToast(json.sheetWarning);
        }
        if (json.data && typeof json.data.sheetSynced === 'boolean') {
          setInvoices((prev) =>
            prev.map((r) => (r.id === record.id ? { ...r, sheetSynced: json.data.sheetSynced } : r))
          );
        }
      }
    } catch (e) {}
  };

  const updateInvoice = async (id: string, updatedRecord: Partial<InvoiceRecord>) => {
    pushInvoiceHistory(invoices);
    const updated = invoices.map(r => r.id === id ? { ...r, ...updatedRecord } : r);
    setInvoices(updated);
    persistInvoices(updated);
    try {
      const res = await fetch(`/api/invoices/${id}`, {
        method: 'PUT',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(updatedRecord)
      });
      if (res.ok) {
        const json = await res.json();
        if (json.sheetWarning) {
          showToast(json.sheetWarning);
        }
        if (json.data && typeof json.data.sheetSynced === 'boolean') {
          setInvoices((prev) =>
            prev.map((r) => (r.id === id ? { ...r, sheetSynced: json.data.sheetSynced } : r))
          );
        }
      }
    } catch (e) {}
  };

  const deleteInvoice = async (id: string) => {
    // 1. Pehle Google Sheet se delete karo (Rule 3)
    const sheetRes = await syncModuleToSheets({
      action: 'delete',
      sheet: 'GSTR-1',
      headers: GSTR1_SHEET_HEADERS,
      ids: [id]
    });

    if (!sheetRes.success) {
      showToast(sheetRes.error ? `Google Sheet delete failed: ${sheetRes.error}` : 'Google Sheet delete failed. Record not deleted.');
      return;
    }

    // 2. Sheet me success milne par hi app state & localStorage se hatao
    pushInvoiceHistory(invoices);
    const updated = invoices.filter(r => r.id !== id);
    setInvoices(updated);
    if (selectedInvoiceId === id) setSelectedInvoiceId(null);
    persistInvoices(updated);
    showToast('Invoice deleted');

    // 3. Backend DB se hatao
    try {
      const res = await fetch(`/api/invoices/${id}`, { method: 'DELETE' });
      if (!res.ok) {
        console.warn('Backend delete returned status', res.status);
      }
    } catch (e) {
      console.warn('Backend delete network error:', e);
    }
  };

  const clearInvoices = async (targetFy?: string, targetMonth?: string) => {
    const normalizeMonth = (m?: string): string => {
      if (!m) return '';
      const clean = m.trim().toLowerCase();
      const months = ['jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec'];
      for (const mon of months) {
        if (clean.startsWith(mon) || clean.includes(mon)) return mon;
      }
      return clean.slice(0, 3);
    };
    const normalizeFy = (fyStr?: string): string => {
      if (!fyStr) return '';
      const digits = fyStr.replace(/[^0-9]/g, '');
      if (digits) return digits;
      return fyStr.trim().toLowerCase().replace(/[\s–-]/g, '');
    };

    const hasFyFilter = targetFy && targetFy !== 'ALL';
    const hasMonthFilter = targetMonth && targetMonth !== 'ALL';

    const toDeleteIds: string[] = [];
    const remaining = invoices.filter(r => {
      const fyMatch = !hasFyFilter || normalizeFy(r.fy) === normalizeFy(targetFy);
      const monthMatch = !hasMonthFilter || normalizeMonth(r.m) === normalizeMonth(targetMonth);
      if (fyMatch && monthMatch) {
        toDeleteIds.push(r.id);
        return false;
      }
      return true;
    });

    if (toDeleteIds.length === 0) {
      showToast('No records to clear.');
      return;
    }

    // 1. Pehle Google Sheet se ek hi batch request me saare IDs delete karo (Rule 3)
    const sheetRes = await syncModuleToSheets({
      action: 'delete',
      sheet: 'GSTR-1',
      headers: GSTR1_SHEET_HEADERS,
      ids: toDeleteIds
    });

    if (!sheetRes.success) {
      showToast(sheetRes.error ? `Google Sheet clear failed: ${sheetRes.error}` : 'Google Sheet clear failed. Records not deleted.');
      return;
    }

    // 2. Sheet me success milne par hi app state & localStorage se hatao
    pushInvoiceHistory(invoices);
    setInvoices(remaining);
    setSelectedInvoiceId(null);
    persistInvoices(remaining);
    showToast(`Cleared ${toDeleteIds.length} records`);

    // 3. Backend DB se hatao
    try {
      const params = new URLSearchParams();
      if (hasFyFilter) params.append('fy', targetFy);
      if (hasMonthFilter) params.append('m', targetMonth);
      await fetch(`/api/invoices/actions/clear?${params.toString()}`, { method: 'DELETE' });
    } catch (e) {
      console.warn('Backend clear network error:', e);
    }
  };

  const batchImportInvoices = async (
    newRecords: InvoiceRecord[],
    importMonth?: string,
    importFy?: string
  ) => {
    pushInvoiceHistory(invoices);
    const updated = [...invoices, ...newRecords];
    setInvoices(updated);
    persistInvoices(updated);

    const effectiveFy = (importFy && importFy !== 'ALL' && importFy.trim()) ? importFy : undefined;
    try {
      const res = await fetch('/api/invoices/actions/bulk-import', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          items: newRecords,
          fixedMonth: importMonth,
          fixedFy: effectiveFy
        })
      });
      if (res.ok) {
        const json = await res.json();
        if (json.sheetWarning) {
          showToast(json.sheetWarning);
        } else if (json.sheetSynced) {
          const syncedIds = new Set(newRecords.map((r) => r.id));
          const finalInvoices = updated.map((r) => (syncedIds.has(r.id) ? { ...r, sheetSynced: true } : r));
          setInvoices(finalInvoices);
          persistInvoices(finalInvoices);
          showToast(`${newRecords.length} invoices imported & Google Sheet me sync ho gaye!`);
        }
      } else {
        // Direct sheet sync fallback
        const rows = newRecords.map((r) => mapRowToSheetRow(r, 'GSTR-1'));
        const sheetRes = await syncModuleToSheets({
          action: 'upsert',
          sheet: 'GSTR-1',
          headers: GSTR1_SHEET_HEADERS,
          rows
        });
        if (sheetRes.success) {
          const syncedIds = new Set(newRecords.map((r) => r.id));
          const finalInvoices = updated.map((r) => (syncedIds.has(r.id) ? { ...r, sheetSynced: true } : r));
          setInvoices(finalInvoices);
          persistInvoices(finalInvoices);
          showToast(`${newRecords.length} invoices imported & Google Sheet me sync ho gaye!`);
        }
      }
    } catch (e) {
      // Direct sheet sync fallback
      try {
        const rows = newRecords.map((r) => mapRowToSheetRow(r, 'GSTR-1'));
        const sheetRes = await syncModuleToSheets({
          action: 'upsert',
          sheet: 'GSTR-1',
          headers: GSTR1_SHEET_HEADERS,
          rows
        });
        if (sheetRes.success) {
          const syncedIds = new Set(newRecords.map((r) => r.id));
          const finalInvoices = updated.map((r) => (syncedIds.has(r.id) ? { ...r, sheetSynced: true } : r));
          setInvoices(finalInvoices);
          persistInvoices(finalInvoices);
          showToast(`${newRecords.length} invoices imported & Google Sheet me sync ho gaye!`);
        }
      } catch {}
    }
  };

  const toggleLock = async (locked?: boolean) => {
    const newLock = typeof locked === 'boolean' ? locked : !isLocked;
    setIsLocked(newLock);
    try {
      localStorage.setItem(LOCAL_STORAGE_KEY_LOCK, newLock ? '1' : '0');
    } catch (e) {}
    persistInvoices(invoices, newLock);
    showToast(newLock ? 'Sheet locked' : 'Sheet unlocked');

    // 1. Backend sync
    try {
      await fetch('/api/invoices/actions/toggle-lock', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ locked: newLock })
      });
    } catch (e) {}

    // 2. Direct Google Sheets sync fallback
    try {
      const ids = invoices.map((r) => r.id);
      await syncModuleToSheets({
        action: newLock ? 'lock' : 'unlock',
        sheet: 'GSTR-1',
        headers: GSTR1_SHEET_HEADERS,
        ids
      });
    } catch (e) {}
  };

  const syncGoogleSheets = async (): Promise<{ success: boolean; count?: number; message?: string }> => {
    try {
      const res = await fetch('/api/invoices/actions/sync-sheets', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' }
      });
      const json = await res.json();
      if (json.message) {
        showToast(json.message);
      }
      if (json.success) {
        await fetchApiInvoices();
      }
      return json;
    } catch (e: unknown) {
      const msg = e instanceof Error ? e.message : 'Sync request failed';
      showToast('Google Sheet sync fail hua: ' + msg);
      return { success: false, message: msg };
    }
  };

  // Step Actions
  const toggleStepDone = async (id: number) => {
    const updated = steps.map(s => s.id === id ? { ...s, done: !s.done } : s);
    setSteps(updated);
    persistSteps(updated);
    try {
      const step = updated.find(s => s.id === id);
      await fetch(`/api/steps/${id}/status`, {
        method: 'PUT',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ done: step?.done })
      });
    } catch (e) {}
  };

  // Settings Actions
  const saveSettings = async (newSettings: BusinessSettings) => {
    setSettings(newSettings);
    persistSettings(newSettings);
    try {
      await fetch('/api/settings', {
        method: 'PUT',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(newSettings)
      });
    } catch (e) {}
    showToast('Settings saved successfully');
  };

  return {
    view,
    setView,
    fy,
    setFy,
    quarter,
    setQuarter,
    month,
    setMonth,
    monthlyData,
    rcmList,
    invoices,
    setInvoices,
    isLocked,
    setIsLocked,
    selectedInvoiceId,
    setSelectedInvoiceId,
    steps,
    setSteps,
    settings,
    setSettings,
    toastMessage,
    showToast,
    isLoggedOut,
    setIsLoggedOut,
    addInvoice,
    updateInvoice,
    deleteInvoice,
    clearInvoices,
    batchImportInvoices,
    toggleLock,
    syncGoogleSheets,
    toggleStepDone,
    saveSettings,
    undoInvoices,
    redoInvoices,
    canUndo: pastInvoices.length > 0,
    canRedo: futureInvoices.length > 0,
    refreshInvoices: fetchApiInvoices
  };
}
