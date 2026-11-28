import { cleanStr, cleanInv, round2, ALL_FY_MONTHS, getFyMonthList } from './recoSyncService';
import { loadRegisterRows, REGISTER_TABS } from './registerService';
import defaultRegisters from '../data/defaultRegisters.json';

export interface OldItcItem {
  id: string;
  originalMonth: string;
  date: string;
  invNo: string;
  gInvNo?: string;
  gstin: string;
  supplierName: string;

  // Books amounts
  bTaxable: number;
  bIgst: number;
  bCgst: number;
  bSgst: number;
  bTotal: number;

  // 2B amounts
  gTaxable: number;
  gIgst: number;
  gCgst: number;
  gSgst: number;
  gTotal: number;

  // Mismatch & Reason
  reason: string;
  mismatchFields: string[];
  itcMatchRatio: string; // e.g. "4/6"

  // Status
  status: 'Pending' | 'Matched - Claim' | 'Not Claim';
  month2b?: string;
  claimMonth?: string;
  pendingMonthsCount: number;

  // Audit notes
  auditStatus?: string;
  auditNote?: string;

  manualNotClaim?: boolean;
  updatedAt: number;
}

export interface SupplierItcRiskSummary {
  supplierName: string;
  gstin: string;
  count: number;
  taxableAtRisk: number;
  itcAtRisk: number;
}

const STORAGE_KEY_OLD_ITC = 'gst_old_itc_pool_v2';
const STORAGE_KEY_AUDIT_NOTES = 'gst_reco_audit_notes_v1';

export function getOldItcPool(): OldItcItem[] {
  try {
    const raw = localStorage.getItem(STORAGE_KEY_OLD_ITC);
    if (raw) {
      const parsed = JSON.parse(raw);
      if (Array.isArray(parsed)) return parsed;
    }
  } catch {}
  return [];
}

export function saveOldItcPool(items: OldItcItem[], emitEvent: boolean = false): void {
  try {
    localStorage.setItem(STORAGE_KEY_OLD_ITC, JSON.stringify(items));
    if (emitEvent) {
      window.dispatchEvent(new CustomEvent('old-itc-updated'));
    }
  } catch (e) {
    console.warn('Failed to save Old ITC pool:', e);
  }
}

/**
 * Calculates how many months have passed between originalMonth and workingMonth
 */
export function calculatePendingMonths(originalMonth: string, workingMonth: string): number {
  const monthList = getFyMonthList(workingMonth);
  const origIdx = monthList.findIndex((m) => m.value === originalMonth);
  const workIdx = monthList.findIndex((m) => m.value === workingMonth);
  if (origIdx === -1 || workIdx === -1 || workIdx < origIdx) return 1;
  return workIdx - origIdx + 1;
}

/**
 * Syncs and updates the Old ITC pool for the current working month
 */
export function syncOldItcPool(workingMonth: string = 'Oct-2026', itcClaimLastDate?: string): {
  oldItcItems: OldItcItem[];
  pendingItems: OldItcItem[];
  matchedClaimItems: OldItcItem[];
  notClaimItems: OldItcItem[];
  supplierRiskSummary: SupplierItcRiskSummary[];
  pendingCount: number;
  pendingTotalItc: number;
  notClaimCount: number;
  notClaimTotalItc: number;
} {
  // 1. Load registers
  let booksRows = loadRegisterRows(REGISTER_TABS.ITC_PURCHASE);
  if (!booksRows || !Array.isArray(booksRows) || booksRows.length === 0) {
    try {
      const raw = localStorage.getItem('itc_register_v1');
      if (raw) booksRows = JSON.parse(raw);
    } catch {}
  }
  if (!booksRows || !Array.isArray(booksRows) || booksRows.length === 0) {
    booksRows = (defaultRegisters as any)?.itc || [];
  }

  let gstr2bRows = loadRegisterRows(REGISTER_TABS.GSTR2B);
  if (!gstr2bRows || !Array.isArray(gstr2bRows) || gstr2bRows.length === 0) {
    try {
      const raw = localStorage.getItem('gstr2b_register_v1');
      if (raw) gstr2bRows = JSON.parse(raw);
    } catch {}
  }
  if (!gstr2bRows || !Array.isArray(gstr2bRows) || gstr2bRows.length === 0) {
    gstr2bRows = (defaultRegisters as any)?.g2b || [];
  }

  // Load audit notes
  let auditNotes: Record<string, { status: string; note: string }> = {};
  try {
    const aRaw = localStorage.getItem(STORAGE_KEY_AUDIT_NOTES);
    if (aRaw) auditNotes = JSON.parse(aRaw);
  } catch {}

  const fyMonthList = getFyMonthList(workingMonth);
  const currentMonthIdx = fyMonthList.findIndex((m) => m.value === workingMonth);
  const effectiveMonthIdx = currentMonthIdx === -1 ? 6 : currentMonthIdx;

  // 2. Identify earlier months' Books invoices (Oct, Sep, Aug, Jul, etc.)
  // Parse month for each book row
  const parseRowMonth = (r: any): string => {
    if (r.month && typeof r.month === 'string' && r.month.includes('-')) return r.month.trim();
    if (r.m && typeof r.m === 'string' && r.m.includes('-')) return r.m.trim();
    const d = String(r.date || r.d || '').trim();
    const parts = d.split(/[-/]/);
    const MONTHS = ['', 'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
    if (parts.length === 3) {
      if (parts[0].length === 4) {
        const m = parseInt(parts[1], 10);
        return `${MONTHS[m] || 'Oct'}-${parts[0]}`;
      } else if (parts[2].length === 4) {
        const m = parseInt(parts[1], 10);
        return `${MONTHS[m] || 'Oct'}-${parts[2]}`;
      }
    }
    const yrMatch = workingMonth.match(/\b(20\d\d)\b/);
    const yr = yrMatch ? yrMatch[1] : '2026';
    return (r.m ? `${r.m}-${yr}` : workingMonth);
  };

  const currentPool = getOldItcPool();
  const poolMap = new Map<string, OldItcItem>(currentPool.map((item) => [item.id, item]));

  // Check if past deadline
  const isPastDeadline = Boolean(
    itcClaimLastDate && new Date().toISOString().slice(0, 10) > itcClaimLastDate.trim()
  );

  // We consider all earlier months up to (and including) previous month of working month
  booksRows.forEach((b: any) => {
    const bMonth = parseRowMonth(b);
    const bMonthIdx = fyMonthList.findIndex((m) => m.value === bMonth);

    // Rule: Only earlier months than workingMonth go into Old ITC carry forward
    const isEarlierMonth = bMonthIdx !== -1 && bMonthIdx < effectiveMonthIdx;
    
    // Also if this invoice was already in the pool, keep it
    if (!isEarlierMonth && !poolMap.has(b.id)) {
      return;
    }

    const bTaxable = round2(Number(b.taxable || b.t || 0));
    const bIgst = round2(Number(b.igst || b.i || 0));
    const bCgst = round2(Number(b.cgst || b.c || 0));
    const bSgst = round2(Number(b.sgst || b.s || 0));
    const bTotal = round2(bTaxable + bIgst + bCgst + bSgst);
    const bInvNo = String(b.invNo || b.n || '').trim();
    const bGstin = String(b.gstin || b.g || '').trim();
    const bSupplier = String(b.supplierName || b.p || '').trim();
    const bDate = String(b.date || b.d || '').trim();

    // Check if Debit Note (Debit notes do NOT go into Old ITC)
    if (bInvNo.toUpperCase().includes('DBN') || bTaxable < 0) {
      return;
    }

    // Check if matched in original month's 2B
    // If not matched, it belongs in Old ITC pool
    let existing = poolMap.get(b.id);
    if (!existing) {
      const pendingCount = calculatePendingMonths(bMonth, workingMonth);
      existing = {
        id: b.id,
        originalMonth: bMonth,
        date: bDate,
        invNo: bInvNo,
        gstin: bGstin,
        supplierName: bSupplier,
        bTaxable,
        bIgst,
        bCgst,
        bSgst,
        bTotal,
        gTaxable: 0,
        gIgst: 0,
        gCgst: 0,
        gSgst: 0,
        gTotal: 0,
        reason: 'Not in 2B: supplier ne file nahi kiya (possible quarterly filer)',
        mismatchFields: ['GSTR-2B Missing'],
        itcMatchRatio: '0/6',
        status: 'Pending',
        pendingMonthsCount: pendingCount,
        auditStatus: auditNotes[b.id]?.status,
        auditNote: auditNotes[b.id]?.note,
        updatedAt: Date.now()
      };
      poolMap.set(b.id, existing);
    } else {
      existing.pendingMonthsCount = calculatePendingMonths(existing.originalMonth, workingMonth);
      existing.bTaxable = bTaxable;
      existing.bIgst = bIgst;
      existing.bCgst = bCgst;
      existing.bSgst = bSgst;
      existing.bTotal = bTotal;
    }
  });

  // 3. Re-match Pending Old ITC records against GSTR-2B across ALL months
  const used2bIds = new Set<string>();
  const poolArray = Array.from(poolMap.values());

  poolArray.forEach((item) => {
    // If manually marked Not Claim or past deadline
    if (item.manualNotClaim || (isPastDeadline && item.status === 'Pending')) {
      item.status = 'Not Claim';
      if (!item.reason.includes('ITC Claim Last Date')) {
        item.reason = isPastDeadline
          ? 'ITC Claim Last Date nikal gayi (Ineligible to claim)'
          : item.reason;
      }
      return;
    }

    // If already claimed, skip re-matching
    if (item.status === 'Matched - Claim' && item.claimMonth) {
      if (!item.month2b) item.month2b = item.claimMonth;
      return;
    }

    // Try 6/6 Match against GSTR-2B rows
    const bGstinClean = cleanStr(item.gstin);
    const bInvClean = cleanInv(item.invNo);

    let perfect2b: any = null;
    let partial2b: any = null;
    let highestFieldMatch = 0;
    let partialMismatches: string[] = [];

    for (const g of gstr2bRows) {
      if (used2bIds.has(g.id)) continue;
      const gGstinClean = cleanStr(g.gstin || g.g);
      const gInvClean = cleanInv(g.invNo || g.n);

      if (bGstinClean !== gGstinClean) continue;

      const gTaxable = round2(Number(g.taxable || g.t || 0));
      const gIgst = round2(Number(g.igst || g.i || 0));
      const gCgst = round2(Number(g.cgst || g.c || 0));
      const gSgst = round2(Number(g.sgst || g.s || 0));

      const invMatch = bInvClean === gInvClean;
      const taxMatch = Math.abs(item.bTaxable - gTaxable) <= 1.00;
      const igstMatch = Math.abs(item.bIgst - gIgst) <= 1.00;
      const cgstMatch = Math.abs(item.bCgst - gCgst) <= 1.00;
      const sgstMatch = Math.abs(item.bSgst - gSgst) <= 1.00;

      let matchCount = 1; // GSTIN matches
      const mismatches: string[] = [];
      if (invMatch) matchCount++; else mismatches.push('Invoice No');
      if (taxMatch) matchCount++; else mismatches.push('Taxable Value');
      if (igstMatch) matchCount++; else mismatches.push('IGST');
      if (cgstMatch) matchCount++; else mismatches.push('CGST');
      if (sgstMatch) matchCount++; else mismatches.push('SGST');

      if (matchCount === 6) {
        perfect2b = g;
        break;
      } else if (matchCount > highestFieldMatch && (invMatch || taxMatch)) {
        highestFieldMatch = matchCount;
        partial2b = g;
        partialMismatches = mismatches;
      }
    }

    if (perfect2b) {
      used2bIds.add(perfect2b.id);
      const gTaxable = round2(Number(perfect2b.taxable || perfect2b.t || 0));
      const gIgst = round2(Number(perfect2b.igst || perfect2b.i || 0));
      const gCgst = round2(Number(perfect2b.cgst || perfect2b.c || 0));
      const gSgst = round2(Number(perfect2b.sgst || perfect2b.s || 0));
      const gMonth = parseRowMonth(perfect2b);

      item.status = 'Matched - Claim';
      item.month2b = gMonth;
      item.claimMonth = gMonth;
      item.itcMatchRatio = '6/6';
      item.reason = `6/6 Match – Reflected in 2B (${gMonth})`;
      item.mismatchFields = [];
      item.gInvNo = perfect2b.invNo || perfect2b.n || item.invNo;
      item.gTaxable = gTaxable;
      item.gIgst = gIgst;
      item.gCgst = gCgst;
      item.gSgst = gSgst;
      item.gTotal = round2(gTaxable + gIgst + gCgst + gSgst);
    } else if (partial2b) {
      const gMonth = parseRowMonth(partial2b);
      item.status = 'Pending';
      item.month2b = gMonth;
      item.claimMonth = undefined;
      item.itcMatchRatio = `${highestFieldMatch}/6`;
      item.reason = `Wrong detail in 2B: ${partialMismatches.join(' / ')}`;
      item.mismatchFields = partialMismatches;
      item.gInvNo = partial2b.invNo || partial2b.n;
      const gTaxable = round2(Number(partial2b.taxable || partial2b.t || 0));
      const gIgst = round2(Number(partial2b.igst || partial2b.i || 0));
      const gCgst = round2(Number(partial2b.cgst || partial2b.c || 0));
      const gSgst = round2(Number(partial2b.sgst || partial2b.s || 0));
      item.gTaxable = gTaxable;
      item.gIgst = gIgst;
      item.gCgst = gCgst;
      item.gSgst = gSgst;
      item.gTotal = round2(gTaxable + gIgst + gCgst + gSgst);
    } else {
      item.status = 'Pending';
      item.month2b = undefined;
      item.claimMonth = undefined;
      item.itcMatchRatio = '0/6';
      item.reason = 'Not in 2B: supplier ne file nahi kiya (possible quarterly filer)';
      item.mismatchFields = ['GSTR-2B Missing'];
      item.gTaxable = 0;
      item.gIgst = 0;
      item.gCgst = 0;
      item.gSgst = 0;
      item.gTotal = 0;
    }
  });

  saveOldItcPool(poolArray, false);

  const pendingItems = poolArray.filter((x) => x.status === 'Pending');
  const matchedClaimItems = poolArray.filter((x) => x.status === 'Matched - Claim');
  const notClaimItems = poolArray.filter((x) => x.status === 'Not Claim');

  const pendingTotalItc = round2(pendingItems.reduce((acc, r) => acc + (r.bIgst + r.bCgst + r.bSgst), 0));
  const notClaimTotalItc = round2(notClaimItems.reduce((acc, r) => acc + (r.bIgst + r.bCgst + r.bSgst), 0));

  // Supplier-wise ITC at risk summary for Not Claim ITC
  const riskMap = new Map<string, SupplierItcRiskSummary>();
  notClaimItems.forEach((r) => {
    const key = r.gstin || r.supplierName;
    const itc = round2(r.bIgst + r.bCgst + r.bSgst);
    const existing = riskMap.get(key);
    if (existing) {
      existing.count += 1;
      existing.taxableAtRisk = round2(existing.taxableAtRisk + r.bTaxable);
      existing.itcAtRisk = round2(existing.itcAtRisk + itc);
    } else {
      riskMap.set(key, {
        supplierName: r.supplierName,
        gstin: r.gstin,
        count: 1,
        taxableAtRisk: r.bTaxable,
        itcAtRisk: itc
      });
    }
  });

  const supplierRiskSummary = Array.from(riskMap.values()).sort((a, b) => b.itcAtRisk - a.itcAtRisk);

  return {
    oldItcItems: poolArray,
    pendingItems,
    matchedClaimItems,
    notClaimItems,
    supplierRiskSummary,
    pendingCount: pendingItems.length,
    pendingTotalItc,
    notClaimCount: notClaimItems.length,
    notClaimTotalItc
  };
}

export function moveToNotClaim(id: string, reason?: string): void {
  const pool = getOldItcPool();
  const target = pool.find((x) => x.id === id);
  if (target) {
    target.status = 'Not Claim';
    target.manualNotClaim = true;
    if (reason) target.reason = reason;
    target.updatedAt = Date.now();
    saveOldItcPool(pool, true);
  }
}

export function moveBackToOldItc(id: string): void {
  const pool = getOldItcPool();
  const target = pool.find((x) => x.id === id);
  if (target) {
    target.status = 'Pending';
    target.manualNotClaim = false;
    target.updatedAt = Date.now();
    saveOldItcPool(pool, true);
  }
}

export function updateOldItcAudit(id: string, auditStatus: string, auditNote: string): void {
  const pool = getOldItcPool();
  const target = pool.find((x) => x.id === id);
  if (target) {
    target.auditStatus = auditStatus;
    target.auditNote = auditNote;
    target.updatedAt = Date.now();
    saveOldItcPool(pool, true);
  }
}
