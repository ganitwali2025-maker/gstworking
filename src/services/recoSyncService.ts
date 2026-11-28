import { useState, useEffect, useCallback, useMemo } from 'react';
import defaultRegisters from '../data/defaultRegisters.json';

export interface ReconciledItem {
  id: string;
  status: 'Matched' | 'Not Match' | 'Not in GSTR-2B' | 'Not in Purchase Register' | 'Debit Note';
  month2b: string;
  date: string;
  invNo: string;
  gstin: string;
  supplierName: string;
  taxable: number;
  igst: number;
  cgst: number;
  sgst: number;
  total: number;
  auditStatus?: string;
  auditNote?: string;
}

export interface ReconciledInvoiceSummary {
  totalCount: number;
  matchedCount: number;
  notMatchCount: number;
  notIn2bCount: number;
  notInBookCount: number;
  debitNoteCount: number;
}

export interface LiveReconciliationResult {
  workingMonth: string;
  summary: ReconciledInvoiceSummary;
  allInvoices: ReconciledItem[];
  matched: ReconciledItem[];
  notInBook: ReconciledItem[];
  notIn2b: ReconciledItem[];
  notMatch: ReconciledItem[];
  debitNote: ReconciledItem[];
  updatedAt: number;
}

export const RECO_SYNC_EVENT = 'reco-data-updated';
export const RECO_STORAGE_KEY = 'gst_reco_active_data_v1';

export const ALL_FY_MONTHS = [
  { value: 'Apr-2026', label: 'April-2026', short: 'Apr' },
  { value: 'May-2026', label: 'May-2026', short: 'May' },
  { value: 'Jun-2026', label: 'June-2026', short: 'Jun' },
  { value: 'Jul-2026', label: 'July-2026', short: 'Jul' },
  { value: 'Aug-2026', label: 'August-2026', short: 'Aug' },
  { value: 'Sep-2026', label: 'September-2026', short: 'Sep' },
  { value: 'Oct-2026', label: 'October-2026', short: 'Oct' },
  { value: 'Nov-2026', label: 'November-2026', short: 'Nov' },
  { value: 'Dec-2026', label: 'December-2026', short: 'Dec' },
  { value: 'Jan-2027', label: 'January-2027', short: 'Jan' },
  { value: 'Feb-2027', label: 'February-2027', short: 'Feb' },
  { value: 'Mar-2027', label: 'March-2027', short: 'Mar' }
];

export function getFyMonthList(fyStr?: string): { value: string; label: string; short: string }[] {
  let yr = '2026';
  if (fyStr) {
    const m = fyStr.replace(/[^0-9]/g, '').slice(0, 4);
    if (m) yr = m;
  }
  const startYr = parseInt(yr, 10) || 2026;
  const endYr = startYr + 1;
  const months = [
    { short: 'Apr', label: 'April' },
    { short: 'May', label: 'May' },
    { short: 'Jun', label: 'June' },
    { short: 'Jul', label: 'July' },
    { short: 'Aug', label: 'August' },
    { short: 'Sep', label: 'September' },
    { short: 'Oct', label: 'October' },
    { short: 'Nov', label: 'November' },
    { short: 'Dec', label: 'December' },
    { short: 'Jan', label: 'January' },
    { short: 'Feb', label: 'February' },
    { short: 'Mar', label: 'March' }
  ];
  return months.map((m) => {
    const y = ['Jan', 'Feb', 'Mar'].includes(m.short) ? endYr : startYr;
    return {
      value: `${m.short}-${y}`,
      label: `${m.label}-${y}`,
      short: m.short
    };
  });
}

export function getMonthPriority(targetM: string, currentM: string, fyStr?: string): number {
  if (targetM === currentM) return 0;
  const monthList = getFyMonthList(fyStr);
  const currIdx = monthList.findIndex((m) => m.value === currentM);
  const targetIdx = monthList.findIndex((m) => m.value === targetM);
  if (currIdx === -1 || targetIdx === -1) return 999;
  if (targetIdx > currIdx) {
    return 10 + (targetIdx - currIdx);
  }
  return 100 + (currIdx - targetIdx);
}

export const cleanStr = (s?: string): string => (s || '').trim().toUpperCase().replace(/\s+/g, ' ');
export const cleanInv = (s?: string): string => (s || '').trim().toUpperCase().replace(/[\s\-_/]/g, '');
export const round2 = (val: number): number => Math.round((val + Number.EPSILON) * 100) / 100;

export const parseRecordMonth = (item: any, defaultMonth: string = 'Oct'): string => {
  const mRaw = (item.m || '').trim();
  const dRaw = (item.d || item.date || '').trim();
  let parsedMonth = mRaw;
  let parsedYear = '2026';
  if (item.fy) {
    const fyMatch = String(item.fy).replace(/[^0-9]/g, '').slice(0, 4);
    if (fyMatch) parsedYear = fyMatch;
  } else if (dRaw) {
    const yrMatch = dRaw.match(/\b(20\d\d)\b/);
    if (yrMatch) parsedYear = yrMatch[1];
  }

  if (!parsedMonth && dRaw) {
    const parts = dRaw.split(/[-/]/);
    const MONTH_NAMES = ['', 'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
    if (parts.length === 3) {
      if (parts[0].length === 4) {
        const mNum = parseInt(parts[1], 10);
        if (mNum >= 1 && mNum <= 12) parsedMonth = MONTH_NAMES[mNum];
      } else if (parts[2].length === 4) {
        const mNum = parseInt(parts[1], 10);
        if (mNum >= 1 && mNum <= 12) parsedMonth = MONTH_NAMES[mNum];
      }
    }
  }

  const MONTH_MAP: Record<string, string> = {
    jan: 'Jan', january: 'Jan',
    feb: 'Feb', february: 'Feb',
    mar: 'Mar', march: 'Mar',
    apr: 'Apr', april: 'Apr',
    may: 'May',
    jun: 'Jun', june: 'Jun',
    jul: 'Jul', july: 'Jul',
    aug: 'Aug', august: 'Aug',
    sep: 'Sep', sept: 'Sep', september: 'Sep',
    oct: 'Oct', october: 'Oct',
    nov: 'Nov', november: 'Nov',
    dec: 'Dec', december: 'Dec'
  };
  const mClean = parsedMonth.toLowerCase();
  const stdMonth = MONTH_MAP[mClean] || (mClean ? mClean.charAt(0).toUpperCase() + mClean.slice(1, 3) : defaultMonth);
  const startYr = parseInt(parsedYear, 10) || 2026;
  if (['Jan', 'Feb', 'Mar'].includes(stdMonth)) {
    parsedYear = String(startYr + 1);
  } else {
    parsedYear = String(startYr);
  }
  return `${stdMonth}-${parsedYear}`;
};

/**
 * Single source of truth directional matching engine (exact match with ReconciliationView)
 */
export function computeLiveReconciliation(workingMonth: string = 'Oct-2026'): LiveReconciliationResult {
  // 1. Read Books Data
  let rawBooksList: any[] = [];
  try {
    const raw = localStorage.getItem('itc_register_v1');
    if (raw) {
      const parsed = JSON.parse(raw);
      if (Array.isArray(parsed) && parsed.length > 0) {
        rawBooksList = parsed;
      }
    }
  } catch {}
  if (rawBooksList.length === 0 && Array.isArray((defaultRegisters as any)?.itc)) {
    rawBooksList = (defaultRegisters as any).itc;
  }

  // 2. Read GSTR-2B Data
  let raw2bList: any[] = [];
  try {
    const raw = localStorage.getItem('gstr2b_register_v1');
    if (raw) {
      const parsed = JSON.parse(raw);
      if (Array.isArray(parsed) && parsed.length > 0) {
        raw2bList = parsed;
      }
    }
  } catch {}
  if (raw2bList.length === 0 && Array.isArray((defaultRegisters as any)?.g2b)) {
    raw2bList = (defaultRegisters as any).g2b;
  }

  // 3. Read Audit Notes
  let auditNotes: Record<string, any> = {};
  try {
    const raw = localStorage.getItem('gst_reco_audit_v1');
    if (raw) {
      auditNotes = JSON.parse(raw);
    }
  } catch {}

  // 4. Map Books Records
  const booksRecords = rawBooksList.map((item: any, idx: number) => {
    let rawTaxable = typeof item.t === 'number' ? item.t : (typeof item.taxable === 'number' ? item.taxable : 0);
    let rawIgst = typeof item.i === 'number' ? item.i : (typeof item.igst === 'number' ? item.igst : 0);
    let rawCgst = typeof item.c === 'number' ? item.c : (typeof item.cgst === 'number' ? item.cgst : 0);
    let rawSgst = typeof item.s === 'number' ? item.s : (typeof item.sgst === 'number' ? item.sgst : 0);
    let rawCess = typeof item.cess === 'number' ? item.cess : 0;

    const sumShift = rawIgst + rawCgst + rawSgst + rawCess;
    if (sumShift !== 0 && Math.abs(rawTaxable - sumShift) < 10) {
      rawTaxable = rawIgst;
      rawIgst = rawCgst;
      rawCgst = rawSgst;
      rawSgst = rawCess;
      rawCess = 0;
    }

    const isPaise = Math.abs(rawTaxable) >= 100000;
    const monthFormatted = parseRecordMonth(item, 'Oct');

    return {
      id: item.id || `book-${idx}`,
      date: item.d || '',
      invNo: item.invNo || item.n || '',
      gstin: item.g || item.gstin || '',
      supplierName: item.p || item.tradeName || item.party || '',
      taxable: isPaise ? rawTaxable / 100 : rawTaxable,
      igst: isPaise ? rawIgst / 100 : rawIgst,
      cgst: isPaise ? rawCgst / 100 : rawCgst,
      sgst: isPaise ? rawSgst / 100 : rawSgst,
      month: monthFormatted
    };
  });

  // 5. Map 2B Records
  const gstr2bRecords = raw2bList.map((item: any, idx: number) => {
    const rawTaxable = typeof item.taxable === 'number' ? item.taxable : (typeof item.t === 'number' ? item.t : 0);
    const rawIgst = typeof item.igst === 'number' ? item.igst : (typeof item.i === 'number' ? item.i : 0);
    const rawCgst = typeof item.cgst === 'number' ? item.cgst : (typeof item.c === 'number' ? item.c : 0);
    const rawSgst = typeof item.sgst === 'number' ? item.sgst : (typeof item.s === 'number' ? item.s : 0);
    const isPaise = Math.abs(rawTaxable) >= 100000;
    const monthFormatted = parseRecordMonth(item, 'Oct');

    return {
      id: item.id || `g2b-${idx}`,
      date: item.d || '',
      invNo: item.invNo || item.n || '',
      gstin: item.gstin || item.g || '',
      supplierName: item.tradeName || item.p || item.supplierName || '',
      taxable: isPaise ? rawTaxable / 100 : rawTaxable,
      igst: isPaise ? rawIgst / 100 : rawIgst,
      cgst: isPaise ? rawCgst / 100 : rawCgst,
      sgst: isPaise ? rawSgst / 100 : rawSgst,
      month: monthFormatted
    };
  });

  const results: ReconciledItem[] = [];
  const paired2bIds = new Set<string>();

  // Filter Books for Working Month
  const workingBooks = booksRecords.filter((b) => b.month === workingMonth);
  const debitNoteBooks: typeof booksRecords = [];
  const regularBooks: typeof booksRecords = [];

  workingBooks.forEach((b) => {
    const isMinus = b.taxable < 0 || b.igst < 0 || b.cgst < 0 || b.sgst < 0;
    if (isMinus) debitNoteBooks.push(b);
    else regularBooks.push(b);
  });

  // Debit Notes -> Always Debit Note status
  debitNoteBooks.forEach((b) => {
    const total = round2(b.taxable + b.igst + b.cgst + b.sgst);
    results.push({
      id: b.id,
      status: 'Debit Note',
      month2b: '-',
      date: b.date,
      invNo: b.invNo,
      gstin: b.gstin,
      supplierName: b.supplierName,
      taxable: b.taxable,
      igst: b.igst,
      cgst: b.cgst,
      sgst: b.sgst,
      total
    });
  });

  // Sort candidate GSTR-2B records by month search priority
  const sorted2bRecords = [...gstr2bRecords].sort(
    (a, b) => getMonthPriority(a.month, workingMonth) - getMonthPriority(b.month, workingMonth)
  );

  // Step 1: Exact 6-Field Match (or Verified Audit Match)
  const unpairedBooksStage1: typeof regularBooks = [];
  regularBooks.forEach((b) => {
    const bGstinClean = cleanStr(b.gstin);
    const bInvClean = cleanInv(b.invNo);
    const audit = auditNotes[b.id] || auditNotes[`${bGstinClean}_${bInvClean}`];

    let matched2b: (typeof sorted2bRecords)[0] | undefined;
    for (const g of sorted2bRecords) {
      if (paired2bIds.has(g.id)) continue;
      const gGstinClean = cleanStr(g.gstin);
      const gInvClean = cleanInv(g.invNo);

      const gstinMatch = bGstinClean === gGstinClean;
      const invMatch = bInvClean === gInvClean;
      const taxMatch = Math.abs(round2(b.taxable - g.taxable)) <= 1.00;
      const igstMatch = Math.abs(round2(b.igst - g.igst)) <= 1.00;
      const cgstMatch = Math.abs(round2(b.cgst - g.cgst)) <= 1.00;
      const sgstMatch = Math.abs(round2(b.sgst - g.sgst)) <= 1.00;

      const is6Match = gstinMatch && invMatch && taxMatch && igstMatch && cgstMatch && sgstMatch;
      const isAuditedMatch = audit?.status === 'Verified' && gstinMatch && invMatch;

      if (is6Match || isAuditedMatch) {
        matched2b = g;
        break;
      }
    }

    if (matched2b) {
      paired2bIds.add(matched2b.id);
      const total = round2(b.taxable + b.igst + b.cgst + b.sgst);
      results.push({
        id: b.id,
        status: 'Matched',
        month2b: matched2b.month || workingMonth,
        date: b.date,
        invNo: b.invNo,
        gstin: b.gstin,
        supplierName: b.supplierName,
        taxable: b.taxable,
        igst: b.igst,
        cgst: b.cgst,
        sgst: b.sgst,
        total,
        auditStatus: audit?.status,
        auditNote: audit?.note
      });
    } else {
      unpairedBooksStage1.push(b);
    }
  });

  // Step 2: Key Match (GSTIN + InvNo) -> 'Not Match'
  const unpairedBooksStage2: typeof regularBooks = [];
  unpairedBooksStage1.forEach((b) => {
    const bGstinClean = cleanStr(b.gstin);
    const bInvClean = cleanInv(b.invNo);
    const audit = auditNotes[b.id] || auditNotes[`${bGstinClean}_${bInvClean}`];

    let paired2b: (typeof sorted2bRecords)[0] | undefined;
    for (const g of sorted2bRecords) {
      if (paired2bIds.has(g.id)) continue;
      if (bGstinClean === cleanStr(g.gstin) && bInvClean === cleanInv(g.invNo)) {
        paired2b = g;
        break;
      }
    }

    if (paired2b) {
      paired2bIds.add(paired2b.id);
      const total = round2(b.taxable + b.igst + b.cgst + b.sgst);
      results.push({
        id: b.id,
        status: 'Not Match',
        month2b: paired2b.month || workingMonth,
        date: b.date,
        invNo: b.invNo,
        gstin: b.gstin,
        supplierName: b.supplierName,
        taxable: b.taxable,
        igst: b.igst,
        cgst: b.cgst,
        sgst: b.sgst,
        total,
        auditStatus: audit?.status,
        auditNote: audit?.note
      });
    } else {
      unpairedBooksStage2.push(b);
    }
  });

  // Step 3: Fuzzy Match (GSTIN + Taxable within ₹1) -> 'Not Match'
  const unpairedBooksStage3: typeof regularBooks = [];
  unpairedBooksStage2.forEach((b) => {
    const bGstinClean = cleanStr(b.gstin);
    const audit = auditNotes[b.id] || auditNotes[`${bGstinClean}_${cleanInv(b.invNo)}`];

    let paired2b: (typeof sorted2bRecords)[0] | undefined;
    for (const g of sorted2bRecords) {
      if (paired2bIds.has(g.id)) continue;
      if (bGstinClean === cleanStr(g.gstin) && Math.abs(round2(b.taxable - g.taxable)) <= 1.00) {
        paired2b = g;
        break;
      }
    }

    if (paired2b) {
      paired2bIds.add(paired2b.id);
      const total = round2(b.taxable + b.igst + b.cgst + b.sgst);
      results.push({
        id: b.id,
        status: 'Not Match',
        month2b: paired2b.month || workingMonth,
        date: b.date,
        invNo: b.invNo,
        gstin: b.gstin,
        supplierName: b.supplierName,
        taxable: b.taxable,
        igst: b.igst,
        cgst: b.cgst,
        sgst: b.sgst,
        total,
        auditStatus: audit?.status,
        auditNote: audit?.note
      });
    } else {
      unpairedBooksStage3.push(b);
    }
  });

  // Step 4: Unpaired Books -> 'Not in GSTR-2B'
  unpairedBooksStage3.forEach((b) => {
    const total = round2(b.taxable + b.igst + b.cgst + b.sgst);
    const audit = auditNotes[b.id] || auditNotes[`${cleanStr(b.gstin)}_${cleanInv(b.invNo)}`];
    results.push({
      id: b.id,
      status: 'Not in GSTR-2B',
      month2b: '-',
      date: b.date,
      invNo: b.invNo,
      gstin: b.gstin,
      supplierName: b.supplierName,
      taxable: b.taxable,
      igst: b.igst,
      cgst: b.cgst,
      sgst: b.sgst,
      total,
      auditStatus: audit?.status,
      auditNote: audit?.note
    });
  });

  // Step 5: Unpaired GSTR-2B for WORKING MONTH -> 'Not in Purchase Register'
  gstr2bRecords.forEach((g) => {
    if (g.month === workingMonth && !paired2bIds.has(g.id)) {
      const total = round2(g.taxable + g.igst + g.cgst + g.sgst);
      const audit = auditNotes[g.id] || auditNotes[`${cleanStr(g.gstin)}_${cleanInv(g.invNo)}`];
      results.push({
        id: g.id,
        status: 'Not in Purchase Register',
        month2b: g.month || workingMonth,
        date: g.date,
        invNo: g.invNo,
        gstin: g.gstin,
        supplierName: g.supplierName,
        taxable: g.taxable,
        igst: g.igst,
        cgst: g.cgst,
        sgst: g.sgst,
        total,
        auditStatus: audit?.status,
        auditNote: audit?.note
      });
    }
  });

  // Partitioned categories (strict single categorization, zero double counts)
  const matched = results.filter((r) => r.status === 'Matched');
  const notMatch = results.filter((r) => r.status === 'Not Match');
  const notIn2b = results.filter((r) => r.status === 'Not in GSTR-2B');
  const notInBook = results.filter((r) => r.status === 'Not in Purchase Register');
  const debitNote = results.filter((r) => r.status === 'Debit Note');

  const summary: ReconciledInvoiceSummary = {
    totalCount: regularBooks.length + debitNoteBooks.length,
    matchedCount: matched.length,
    notMatchCount: notMatch.length,
    notIn2bCount: notIn2b.length,
    notInBookCount: notInBook.length,
    debitNoteCount: debitNote.length
  };

  const payload: LiveReconciliationResult = {
    workingMonth,
    summary,
    allInvoices: results,
    matched,
    notInBook,
    notIn2b,
    notMatch,
    debitNote,
    updatedAt: Date.now()
  };

  // Cache in localStorage
  try {
    localStorage.setItem(RECO_STORAGE_KEY, JSON.stringify(payload));
  } catch {}

  return payload;
}

/**
 * Get active reconciliation data (from localStorage cache if fresh, otherwise compute)
 */
export function getLiveReconciliationData(workingMonth: string = 'Oct-2026'): LiveReconciliationResult {
  try {
    const raw = localStorage.getItem(RECO_STORAGE_KEY);
    if (raw) {
      const parsed = JSON.parse(raw);
      if (
        parsed &&
        parsed.workingMonth === workingMonth &&
        Array.isArray(parsed.allInvoices) &&
        parsed.summary
      ) {
        return parsed as LiveReconciliationResult;
      }
    }
  } catch {}
  return computeLiveReconciliation(workingMonth);
}

/**
 * React hook for live one-way reading of Reco sheet data across components
 */
export function useLiveReconciliation(workingMonth: string = 'Oct-2026') {
  const [data, setData] = useState<LiveReconciliationResult>(() => getLiveReconciliationData(workingMonth));

  const refreshData = useCallback(() => {
    const live = computeLiveReconciliation(workingMonth);
    setData(live);
  }, [workingMonth]);

  useEffect(() => {
    // Initial fetch / calculation
    refreshData();

    const handleRecoUpdate = (e: any) => {
      if (e?.detail && e.detail.workingMonth === workingMonth && e.detail.summary) {
        const d = e.detail;
        const matched = (d.invoices || []).filter((r: any) => r.status === 'Matched');
        const notInBook = (d.invoices || []).filter((r: any) => r.status === 'Not in Purchase Register');
        const notIn2b = (d.invoices || []).filter((r: any) => r.status === 'Not in GSTR-2B');
        const notMatch = (d.invoices || []).filter((r: any) => r.status === 'Not Match');
        const debitNote = (d.invoices || []).filter((r: any) => r.status === 'Debit Note');

        setData({
          workingMonth,
          summary: d.summary,
          allInvoices: d.invoices || [],
          matched,
          notInBook,
          notIn2b,
          notMatch,
          debitNote,
          updatedAt: Date.now()
        });
      } else {
        refreshData();
      }
    };

    const handleStorage = (e: StorageEvent) => {
      if (
        !e.key ||
        e.key === 'itc_register_v1' ||
        e.key === 'gstr2b_register_v1' ||
        e.key === 'gst_reco_audit_v1' ||
        e.key === RECO_STORAGE_KEY
      ) {
        refreshData();
      }
    };

    const handleRegisterSynced = () => {
      refreshData();
    };

    window.addEventListener(RECO_SYNC_EVENT, handleRecoUpdate as EventListener);
    window.addEventListener('register-synced', handleRegisterSynced);
    window.addEventListener('storage', handleStorage);

    return () => {
      window.removeEventListener(RECO_SYNC_EVENT, handleRecoUpdate as EventListener);
      window.removeEventListener('register-synced', handleRegisterSynced);
      window.removeEventListener('storage', handleStorage);
    };
  }, [workingMonth, refreshData]);

  return data;
}
