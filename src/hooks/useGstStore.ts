import { useState, useEffect, useCallback } from 'react';
import { ViewType, InvoiceRecord, StepItem, BusinessSettings } from '../types/gst';
import { SEED_MONTHLY_DATA, SEED_STEPS, SEED_SETTINGS, SEED_INVOICES, SEED_RCM } from '../data/seedData';
import { FYS, QM } from '../utils/gstFormat';

const LOCAL_STORAGE_KEY_INVOICES = 'gstr1_register_v4';
const LOCAL_STORAGE_KEY_LOCK = 'gstr1_lock_v1';
const LOCAL_STORAGE_KEY_SETTINGS = 'gst_settings_v1';
const LOCAL_STORAGE_KEY_STEPS = 'gst_steps_v1';

export function useGstStore() {
  const [view, setView] = useState<ViewType>('itc');
  const [fy, setFy] = useState<string>(FYS[0]);
  const [quarter, setQuarter] = useState<string>('');
  const [month, setMonth] = useState<string>('Oct 2026');

  // Dashboard Data
  const [monthlyData, setMonthlyData] = useState(SEED_MONTHLY_DATA);
  const [rcmList] = useState(SEED_RCM);

  // Invoices & Lock
  const [invoices, setInvoices] = useState<InvoiceRecord[]>(() => {
    try {
      const saved = localStorage.getItem(LOCAL_STORAGE_KEY_INVOICES);
      return saved ? JSON.parse(saved) : SEED_INVOICES;
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
  const [isLoggedOut, setIsLoggedOut] = useState<boolean>(false);

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

  // Sync with API when available
  const fetchApiInvoices = async () => {
    try {
      const res = await fetch(`/api/invoices?fy=${encodeURIComponent(fy)}`);
      if (res.ok) {
        const json = await res.json();
        if (json.success && Array.isArray(json.data)) {
          setInvoices(json.data);
          setIsLocked(!!json.locked);
          persistInvoices(json.data, !!json.locked);
        }
      }
    } catch (e) {
      // Backend offline or running standalone
    }
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
  }, [fy]);

  // Invoice Actions
  const addInvoice = async (record: InvoiceRecord) => {
    pushInvoiceHistory(invoices);
    const updated = [record, ...invoices];
    setInvoices(updated);
    persistInvoices(updated);
    try {
      await fetch('/api/invoices', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(record)
      });
    } catch (e) {}
  };

  const updateInvoice = async (id: string, updatedRecord: Partial<InvoiceRecord>) => {
    pushInvoiceHistory(invoices);
    const updated = invoices.map(r => r.id === id ? { ...r, ...updatedRecord } : r);
    setInvoices(updated);
    persistInvoices(updated);
    try {
      await fetch(`/api/invoices/${id}`, {
        method: 'PUT',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(updatedRecord)
      });
    } catch (e) {}
  };

  const deleteInvoice = async (id: string) => {
    pushInvoiceHistory(invoices);
    const updated = invoices.filter(r => r.id !== id);
    setInvoices(updated);
    if (selectedInvoiceId === id) setSelectedInvoiceId(null);
    persistInvoices(updated);
    try {
      await fetch(`/api/invoices/${id}`, { method: 'DELETE' });
    } catch (e) {}
  };

  const clearInvoices = async () => {
    pushInvoiceHistory(invoices);
    const updated = invoices.filter(r => r.fy !== fy);
    setInvoices(updated);
    setSelectedInvoiceId(null);
    persistInvoices(updated);
    try {
      await fetch(`/api/invoices/actions/clear?fy=${encodeURIComponent(fy)}`, { method: 'DELETE' });
    } catch (e) {}
  };

  const toggleLock = async (locked?: boolean) => {
    const newLock = typeof locked === 'boolean' ? locked : !isLocked;
    setIsLocked(newLock);
    persistInvoices(invoices, newLock);
    try {
      await fetch('/api/invoices/actions/toggle-lock', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ locked: newLock })
      });
    } catch (e) {}
    showToast(newLock ? 'Sheet locked' : 'Sheet unlocked');
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
    toggleLock,
    toggleStepDone,
    saveSettings,
    undoInvoices,
    redoInvoices,
    canUndo: pastInvoices.length > 0,
    canRedo: futureInvoices.length > 0,
    refreshInvoices: fetchApiInvoices
  };
}
