import React, { useState, useMemo, useEffect } from 'react';
import * as XLSX from 'xlsx';
import {
  syncOldItcPool,
  moveToNotClaim,
  updateOldItcAudit,
  OldItcItem
} from '../services/oldItcService';

interface OldItcDetailSheetProps {
  onBack: () => void;
  showToast: (msg: string) => void;
  workingMonth?: string;
}

function formatInr(val: number): string {
  if (isNaN(val)) return '0.00';
  return val.toLocaleString('en-IN', {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2
  });
}

export default function OldItcDetailSheet({
  onBack,
  showToast,
  workingMonth = 'Oct-2026'
}: OldItcDetailSheetProps) {
  const [searchTerm, setSearchTerm] = useState('');
  const [monthFilter, setMonthFilter] = useState('ALL');
  const [statusFilter, setStatusFilter] = useState<'ALL' | 'Pending' | 'Matched - Claim'>('ALL');
  const [poolVersion, setPoolVersion] = useState(0);

  // Selected invoice for View modal
  const [viewItem, setViewItem] = useState<OldItcItem | null>(null);

  // Selected invoice for Audit modal
  const [auditItem, setAuditItem] = useState<OldItcItem | null>(null);
  const [auditStatus, setAuditStatus] = useState<string>('Pending');
  const [auditNote, setAuditNote] = useState<string>('');

  // Sync Old ITC pool
  const syncResult = useMemo(() => {
    return syncOldItcPool(workingMonth);
  }, [workingMonth, poolVersion]);

  // Listen to old-itc-updated events
  useEffect(() => {
    const handleUpdate = () => setPoolVersion((v) => v + 1);
    window.addEventListener('old-itc-updated', handleUpdate);
    return () => window.removeEventListener('old-itc-updated', handleUpdate);
  }, []);

  // Filter pool items
  const filteredItems = useMemo(() => {
    return syncResult.oldItcItems.filter((item) => {
      // Status filter
      if (statusFilter !== 'ALL' && item.status !== statusFilter) {
        return false;
      }

      // Month filter
      if (monthFilter !== 'ALL' && item.originalMonth !== monthFilter) {
        return false;
      }

      // Search term
      if (searchTerm.trim()) {
        const q = searchTerm.toLowerCase();
        const matchName = item.supplierName.toLowerCase().includes(q);
        const matchGst = item.gstin.toLowerCase().includes(q);
        const matchInv = item.invNo.toLowerCase().includes(q);
        const matchReason = (item.reason || '').toLowerCase().includes(q);
        const matchStatus = (item.status || '').toLowerCase().includes(q);
        const matchBookMonth = (item.originalMonth || '').toLowerCase().includes(q);
        const match2bMonth = (item.month2b || '').toLowerCase().includes(q);
        const matchClaimMonth = (item.claimMonth || '').toLowerCase().includes(q);
        if (!matchName && !matchGst && !matchInv && !matchReason && !matchStatus && !matchBookMonth && !match2bMonth && !matchClaimMonth) {
          return false;
        }
      }

      return true;
    });
  }, [syncResult.oldItcItems, statusFilter, monthFilter, searchTerm]);

  // Unique months available
  const availableMonths = useMemo(() => {
    const set = new Set<string>();
    syncResult.oldItcItems.forEach((i) => {
      if (i.originalMonth) set.add(i.originalMonth);
    });
    return Array.from(set);
  }, [syncResult.oldItcItems]);

  // Totals of filtered items
  const summaryValues = useMemo(() => {
    let taxable = 0;
    let igst = 0;
    let cgst = 0;
    let sgst = 0;
    let totalTax = 0;
    let grandTotal = 0;

    filteredItems.forEach((r) => {
      taxable += r.bTaxable;
      igst += r.bIgst;
      cgst += r.bCgst;
      sgst += r.bSgst;
    });

    totalTax = igst + cgst + sgst;
    grandTotal = taxable + totalTax;

    return { taxable, igst, cgst, sgst, totalTax, grandTotal };
  }, [filteredItems]);

  // Percentage calculations for Pending vs Matched & Claimed (always summing to 100%)
  const { pendingPctStr, matchedPctStr } = useMemo(() => {
    const total = syncResult.oldItcItems.length;
    const pCount = syncResult.pendingCount;
    const mCount = syncResult.matchedClaimItems.length;

    if (total <= 0) {
      return { pendingPctStr: '0%', matchedPctStr: '0%' };
    }
    if (pCount <= 0 && mCount <= 0) {
      return { pendingPctStr: '0%', matchedPctStr: '0%' };
    }
    if (pCount === total) {
      return { pendingPctStr: '100%', matchedPctStr: '0%' };
    }
    if (mCount === total) {
      return { pendingPctStr: '0%', matchedPctStr: '100%' };
    }

    const rawPending = Math.round(((pCount / total) * 100) * 10) / 10;
    const rawMatched = Math.round((100 - rawPending) * 10) / 10;

    const pStr = Number.isInteger(rawPending) ? `${rawPending}%` : `${rawPending.toFixed(1)}%`;
    const mStr = Number.isInteger(rawMatched) ? `${rawMatched}%` : `${rawMatched.toFixed(1)}%`;

    return { pendingPctStr: pStr, matchedPctStr: mStr };
  }, [syncResult.oldItcItems.length, syncResult.pendingCount, syncResult.matchedClaimItems.length]);

  // Move to Not Claim handler
  const handleMoveToNotClaim = (item: OldItcItem) => {
    if (window.confirm(`Move Invoice "${item.invNo}" (${item.supplierName}) to Not Claim ITC?`)) {
      moveToNotClaim(item.id, 'Manually moved by user to Not Claim ITC');
      showToast(`Invoice "${item.invNo}" moved to Not Claim ITC.`);
      setPoolVersion((v) => v + 1);
    }
  };

  // Open Audit Modal
  const handleOpenAudit = (item: OldItcItem) => {
    setAuditItem(item);
    setAuditStatus(item.auditStatus || 'Pending');
    setAuditNote(item.auditNote || '');
  };

  // Save Audit Note
  const handleSaveAudit = () => {
    if (!auditItem) return;
    updateOldItcAudit(auditItem.id, auditStatus, auditNote);
    showToast(`Audit updated for Invoice ${auditItem.invNo}`);
    setAuditItem(null);
    setPoolVersion((v) => v + 1);
  };

  // Export to Excel
  const handleExportExcel = () => {
    try {
      const exportData = filteredItems.map((r, i) => ({
        'S.No': i + 1,
        'Book Month': r.originalMonth,
        '2B Month': r.month2b || '-',
        'Claim Month': r.status === 'Matched - Claim' ? (r.claimMonth || '-') : '-',
        'Invoice No.': r.invNo,
        'Supplier GSTIN': r.gstin,
        'Supplier Name': r.supplierName,
        'Reason': r.reason,
        'Status': r.status,
        'Books Taxable (₹)': r.bTaxable,
        'Books IGST (₹)': r.bIgst,
        'Books CGST (₹)': r.bCgst,
        'Books SGST (₹)': r.bSgst,
        'Books Total (₹)': r.bTotal,
        '2B Taxable (₹)': r.gTaxable,
        '2B IGST (₹)': r.gIgst,
        '2B CGST (₹)': r.gCgst,
        '2B SGST (₹)': r.gSgst,
        '2B Total (₹)': r.gTotal
      }));

      const ws = XLSX.utils.json_to_sheet(exportData);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, 'Old_ITC_Carried_Forward');
      XLSX.writeFile(wb, `Old_ITC_Carried_Forward_${new Date().toISOString().slice(0, 10)}.xlsx`);
      showToast('Old ITC Excel downloaded successfully!');
    } catch (e) {
      console.error(e);
      showToast('Failed to export Excel.');
    }
  };

  return (
    <section className="detail-sheet-page" style={{ padding: '4px 0 24px 0' }}>
      {/* 1. Unified Single Header & Metrics Template */}
      <div
        style={{
          background: '#ffffff',
          border: '1.5px solid #e2e8f0',
          padding: '12px 20px',
          marginBottom: '14px',
          borderRadius: 0,
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'space-between',
          gap: '20px',
          flexWrap: 'wrap'
        }}
      >
        {/* Left: Title & Working Month info */}
        <div style={{ minWidth: '220px', borderRight: '1.5px solid #e2e8f0', paddingRight: '20px' }}>
          <h2 style={{ margin: 0, fontSize: '18px', fontWeight: 800, color: '#241a3d', letterSpacing: '-0.01em', whiteSpace: 'nowrap' }}>
            Old ITC – Carried Forward
          </h2>
          <div style={{ fontSize: '12px', color: '#64748b', marginTop: '3px', fontWeight: 600, whiteSpace: 'nowrap' }}>
            {filteredItems.length} invoices | Working Month: {workingMonth}
          </div>
        </div>

        {/* Right: 4 Metrics in single row */}
        <div
          style={{
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'space-between',
            flex: 1,
            gap: '16px',
            minWidth: '520px'
          }}
        >
          <div>
            <div style={{ fontSize: '11px', fontWeight: 700, color: '#64748b', textTransform: 'uppercase', letterSpacing: '0.04em' }}>
              TOTAL INVOICES
            </div>
            <div style={{ fontSize: '22px', fontWeight: 800, color: '#241a3d', marginTop: '3px' }}>
              {syncResult.oldItcItems.length}
            </div>
          </div>

          <div>
            <div style={{ fontSize: '11px', fontWeight: 700, color: '#64748b', textTransform: 'uppercase', letterSpacing: '0.04em' }}>
              Pending Invoices
            </div>
            <div style={{ fontSize: '22px', fontWeight: 800, color: '#6D28D9', marginTop: '3px', display: 'flex', alignItems: 'baseline', gap: '5px' }}>
              <span>{syncResult.pendingCount}</span>
              <span style={{ fontSize: '12px', fontWeight: 700, color: '#6D28D9' }}>({pendingPctStr})</span>
            </div>
          </div>

          <div>
            <div style={{ fontSize: '11px', fontWeight: 700, color: '#64748b', textTransform: 'uppercase', letterSpacing: '0.04em' }}>
              Matched &amp; Claimed
            </div>
            <div style={{ fontSize: '22px', fontWeight: 800, color: '#15803d', marginTop: '3px', display: 'flex', alignItems: 'baseline', gap: '5px' }}>
              <span>{syncResult.matchedClaimItems.length}</span>
              <span style={{ fontSize: '12px', fontWeight: 700, color: '#15803d' }}>({matchedPctStr})</span>
            </div>
          </div>

          <div>
            <div style={{ fontSize: '11px', fontWeight: 700, color: '#64748b', textTransform: 'uppercase', letterSpacing: '0.04em' }}>
              Total Pending ITC
            </div>
            <div style={{ fontSize: '19px', fontWeight: 800, color: '#6D28D9', marginTop: '4px' }}>
              ₹{formatInr(syncResult.pendingTotalItc)}
            </div>
          </div>
        </div>
      </div>

      {/* 3. 5 Summary Metric Boxes */}
      <div
        style={{
          display: 'grid',
          gridTemplateColumns: 'repeat(auto-fit, minmax(150px, 1fr))',
          gap: '10px',
          marginBottom: '16px'
        }}
      >
        <div style={{ background: '#ffffff', border: '1.5px solid #e2e8f0', padding: '10px 14px', borderRadius: 0 }}>
          <div style={{ fontSize: '11.5px', fontWeight: 700, color: '#64748b' }}>Taxable Value</div>
          <div style={{ fontSize: '16px', fontWeight: 800, color: '#241a3d', marginTop: '4px' }}>
            ₹{formatInr(summaryValues.taxable)}
          </div>
        </div>

        <div style={{ background: '#ffffff', border: '1.5px solid #e2e8f0', padding: '10px 14px', borderRadius: 0 }}>
          <div style={{ fontSize: '11.5px', fontWeight: 700, color: '#64748b' }}>IGST</div>
          <div style={{ fontSize: '16px', fontWeight: 800, color: '#241a3d', marginTop: '4px' }}>
            ₹{formatInr(summaryValues.igst)}
          </div>
        </div>

        <div style={{ background: '#ffffff', border: '1.5px solid #e2e8f0', padding: '10px 14px', borderRadius: 0 }}>
          <div style={{ fontSize: '11.5px', fontWeight: 700, color: '#64748b' }}>CGST</div>
          <div style={{ fontSize: '16px', fontWeight: 800, color: '#241a3d', marginTop: '4px' }}>
            ₹{formatInr(summaryValues.cgst)}
          </div>
        </div>

        <div style={{ background: '#ffffff', border: '1.5px solid #e2e8f0', padding: '10px 14px', borderRadius: 0 }}>
          <div style={{ fontSize: '11.5px', fontWeight: 700, color: '#64748b' }}>SGST</div>
          <div style={{ fontSize: '16px', fontWeight: 800, color: '#241a3d', marginTop: '4px' }}>
            ₹{formatInr(summaryValues.sgst)}
          </div>
        </div>

        <div style={{ background: '#ffffff', border: '1.5px solid #e2e8f0', padding: '10px 14px', borderRadius: 0 }}>
          <div style={{ fontSize: '11.5px', fontWeight: 700, color: '#64748b' }}>Total Tax Amount</div>
          <div style={{ fontSize: '16px', fontWeight: 800, color: '#6D28D9', marginTop: '4px' }}>
            ₹{formatInr(summaryValues.totalTax)}
          </div>
        </div>
      </div>

      {/* 4. Controls: Search, Filters & Excel Export */}
      <div
        style={{
          display: 'flex',
          justifyContent: 'space-between',
          alignItems: 'center',
          gap: '12px',
          background: '#ffffff',
          border: '1.5px solid #e2e8f0',
          padding: '10px 16px',
          marginBottom: '12px',
          borderRadius: 0,
          flexWrap: 'wrap'
        }}
      >
        <div style={{ display: 'flex', alignItems: 'center', gap: '8px', flex: 1, minWidth: '240px', flexWrap: 'wrap' }}>
          <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="#64748b" strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round">
            <circle cx="11" cy="11" r="8"></circle>
            <line x1="21" y1="21" x2="16.65" y2="16.65"></line>
          </svg>
          <input
            type="text"
            value={searchTerm}
            onChange={(e) => setSearchTerm(e.target.value)}
            placeholder="Search Supplier Name, GSTIN, Invoice No..."
            style={{
              width: '100%',
              maxWidth: '300px',
              border: '1px solid #cbd5e1',
              padding: '6px 12px',
              fontSize: '12.5px',
              outline: 'none',
              borderRadius: 0
            }}
          />
          {searchTerm && (
            <button
              type="button"
              onClick={() => setSearchTerm('')}
              style={{ border: 'none', background: 'transparent', cursor: 'pointer', color: '#64748b', fontSize: '13px' }}
            >
              ✕ Clear
            </button>
          )}

          {/* Month Filter */}
          <select
            value={monthFilter}
            onChange={(e) => setMonthFilter(e.target.value)}
            style={{
              border: '1px solid #cbd5e1',
              padding: '6px 10px',
              fontSize: '12px',
              fontWeight: 600,
              color: '#334155',
              borderRadius: 0,
              background: '#f8fafc'
            }}
          >
            <option value="ALL">All Months</option>
            {availableMonths.map((m) => (
              <option key={m} value={m}>{m}</option>
            ))}
          </select>

          {/* Status Filter */}
          <select
            value={statusFilter}
            onChange={(e) => setStatusFilter(e.target.value as any)}
            style={{
              border: '1px solid #cbd5e1',
              padding: '6px 10px',
              fontSize: '12px',
              fontWeight: 600,
              color: '#334155',
              borderRadius: 0,
              background: '#f8fafc'
            }}
          >
            <option value="ALL">All Status</option>
            <option value="Pending">Pending</option>
            <option value="Matched - Claim">Matched - Claim</option>
          </select>
        </div>

        <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
          <span style={{ fontSize: '12px', color: '#64748b', fontWeight: 600 }}>
            Showing {filteredItems.length} of {syncResult.oldItcItems.length} records
          </span>
          <button
            type="button"
            onClick={handleExportExcel}
            style={{
              display: 'inline-flex',
              alignItems: 'center',
              gap: '6px',
              background: '#15803d',
              color: '#ffffff',
              border: 'none',
              padding: '6px 14px',
              fontSize: '12px',
              fontWeight: 700,
              cursor: 'pointer',
              borderRadius: 0
            }}
          >
            <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round">
              <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4" />
              <polyline points="7 10 12 15 17 10" />
              <line x1="12" y1="15" x2="12" y2="3" />
            </svg>
            Export Excel
          </button>
        </div>
      </div>

      {/* 5. Main Old ITC Table */}
      <div className="detail-sheet-table-wrapper" style={{ overflowX: 'auto', border: '1.5px solid #e2e8f0', background: '#ffffff' }}>
        <table className="detail-sheet-table" style={{ width: '100%', borderCollapse: 'collapse', fontSize: '13px', color: '#1E1747' }}>
          <thead>
            {/* Top Multi-Header Row */}
            <tr style={{ background: '#2B0A4F', color: '#ffffff', textAlign: 'center' }}>
              <th
                colSpan={9}
                style={{
                  padding: '7px 10px',
                  fontSize: '13px',
                  fontWeight: 800,
                  letterSpacing: '0.05em',
                  textTransform: 'uppercase',
                  textAlign: 'center',
                  borderRight: '1px solid rgba(255, 255, 255, 0.12)',
                  borderBottom: '1px solid rgba(255, 255, 255, 0.15)',
                  background: '#2B0A4F',
                  fontFamily: "'Bricolage Grotesque', sans-serif"
                }}
              >
                Invoice &amp; Status Details
              </th>
              <th
                colSpan={5}
                style={{
                  padding: '7px 10px',
                  fontSize: '13px',
                  fontWeight: 800,
                  letterSpacing: '0.05em',
                  textTransform: 'uppercase',
                  textAlign: 'center',
                  background: '#240742',
                  borderRight: '1px solid rgba(255, 255, 255, 0.25)',
                  borderBottom: '1px solid rgba(255, 255, 255, 0.15)',
                  fontFamily: "'Bricolage Grotesque', sans-serif"
                }}
              >
                Books (Purchase Register)
              </th>
              <th
                colSpan={5}
                style={{
                  padding: '7px 10px',
                  fontSize: '13px',
                  fontWeight: 800,
                  letterSpacing: '0.05em',
                  textTransform: 'uppercase',
                  textAlign: 'center',
                  background: '#1d0535',
                  borderRight: '1px solid rgba(255, 255, 255, 0.12)',
                  borderBottom: '1px solid rgba(255, 255, 255, 0.15)',
                  fontFamily: "'Bricolage Grotesque', sans-serif"
                }}
              >
                Govt GSTR-2B
              </th>
              <th
                rowSpan={2}
                style={{
                  padding: '10px 12px',
                  fontSize: '12px',
                  fontWeight: 800,
                  letterSpacing: '0.05em',
                  textTransform: 'uppercase',
                  textAlign: 'center',
                  minWidth: '180px',
                  width: '180px',
                  background: '#2B0A4F',
                  color: '#ffffff',
                  borderBottom: '1px solid rgba(255, 255, 255, 0.15)',
                  borderRight: '1px solid rgba(255, 255, 255, 0.12)',
                  fontFamily: "'Bricolage Grotesque', sans-serif"
                }}
              >
                Actions
              </th>
            </tr>

            {/* Column Header Row */}
            <tr
              style={{
                background: '#4A1A7A',
                color: '#ffffff',
                fontSize: '11.5px',
                fontWeight: 700,
                textTransform: 'uppercase',
                letterSpacing: '0.03em',
                whiteSpace: 'nowrap',
                fontFamily: "'Bricolage Grotesque', sans-serif"
              }}
            >
              {/* 1. S.No */}
              <th
                style={{
                  width: '55px',
                  minWidth: '55px',
                  padding: '6px 8px',
                  textAlign: 'center',
                  background: '#4A1A7A',
                  borderRight: '1px solid rgba(255, 255, 255, 0.10)',
                  borderBottom: '1px solid rgba(255, 255, 255, 0.12)'
                }}
              >
                S.No
              </th>

              {/* 2. Book Month */}
              <th
                style={{
                  width: '105px',
                  minWidth: '105px',
                  padding: '6px 10px',
                  textAlign: 'center',
                  background: '#4A1A7A',
                  borderRight: '1px solid rgba(255, 255, 255, 0.10)',
                  borderBottom: '1px solid rgba(255, 255, 255, 0.12)'
                }}
              >
                Book Month
              </th>

              {/* 3. 2B Month */}
              <th
                style={{
                  width: '100px',
                  minWidth: '100px',
                  padding: '6px 10px',
                  textAlign: 'center',
                  borderRight: '1px solid rgba(255, 255, 255, 0.10)',
                  borderBottom: '1px solid rgba(255, 255, 255, 0.12)'
                }}
              >
                2B Month
              </th>

              {/* 4. Claim Month */}
              <th
                style={{
                  width: '105px',
                  minWidth: '105px',
                  padding: '6px 10px',
                  textAlign: 'center',
                  borderRight: '1px solid rgba(255, 255, 255, 0.10)',
                  borderBottom: '1px solid rgba(255, 255, 255, 0.12)'
                }}
              >
                Claim Month
              </th>

              {/* 5. Invoice No. */}
              <th
                style={{
                  width: '140px',
                  minWidth: '140px',
                  padding: '6px 12px',
                  textAlign: 'center',
                  borderRight: '1px solid rgba(255, 255, 255, 0.10)',
                  borderBottom: '1px solid rgba(255, 255, 255, 0.12)'
                }}
              >
                Invoice No.
              </th>

              {/* 6. Supplier GSTIN */}
              <th
                style={{
                  width: '155px',
                  minWidth: '155px',
                  padding: '6px 12px',
                  textAlign: 'center',
                  borderRight: '1px solid rgba(255, 255, 255, 0.10)',
                  borderBottom: '1px solid rgba(255, 255, 255, 0.12)'
                }}
              >
                Supplier GSTIN
              </th>

              {/* 7. Supplier Name */}
              <th
                style={{
                  width: '220px',
                  minWidth: '220px',
                  padding: '6px 14px',
                  textAlign: 'center',
                  borderRight: '1px solid rgba(255, 255, 255, 0.10)',
                  borderBottom: '1px solid rgba(255, 255, 255, 0.12)'
                }}
              >
                Supplier Name
              </th>

              {/* 8. Reason */}
              <th
                style={{
                  width: '200px',
                  minWidth: '200px',
                  padding: '6px 12px',
                  textAlign: 'center',
                  borderRight: '1px solid rgba(255, 255, 255, 0.10)',
                  borderBottom: '1px solid rgba(255, 255, 255, 0.12)'
                }}
              >
                Reason
              </th>

              {/* 9. Status */}
              <th
                style={{
                  width: '130px',
                  minWidth: '130px',
                  padding: '6px 10px',
                  textAlign: 'center',
                  borderRight: '1px solid rgba(255, 255, 255, 0.10)',
                  borderBottom: '1px solid rgba(255, 255, 255, 0.12)'
                }}
              >
                Status
              </th>

              {/* BOOKS (Purchase Register) 5 Columns */}
              <th
                style={{
                  width: '115px',
                  minWidth: '115px',
                  padding: '6px 10px',
                  textAlign: 'center',
                  background: '#3b1263',
                  borderRight: '1px solid rgba(255, 255, 255, 0.10)',
                  borderBottom: '1px solid rgba(255, 255, 255, 0.12)'
                }}
              >
                Taxable
              </th>
              <th
                style={{
                  width: '105px',
                  minWidth: '105px',
                  padding: '6px 10px',
                  textAlign: 'center',
                  background: '#3b1263',
                  borderRight: '1px solid rgba(255, 255, 255, 0.10)',
                  borderBottom: '1px solid rgba(255, 255, 255, 0.12)'
                }}
              >
                IGST
              </th>
              <th
                style={{
                  width: '105px',
                  minWidth: '105px',
                  padding: '6px 10px',
                  textAlign: 'center',
                  background: '#3b1263',
                  borderRight: '1px solid rgba(255, 255, 255, 0.10)',
                  borderBottom: '1px solid rgba(255, 255, 255, 0.12)'
                }}
              >
                CGST
              </th>
              <th
                style={{
                  width: '105px',
                  minWidth: '105px',
                  padding: '6px 10px',
                  textAlign: 'center',
                  background: '#3b1263',
                  borderRight: '1px solid rgba(255, 255, 255, 0.10)',
                  borderBottom: '1px solid rgba(255, 255, 255, 0.12)'
                }}
              >
                SGST
              </th>
              <th
                style={{
                  width: '115px',
                  minWidth: '115px',
                  padding: '6px 10px',
                  textAlign: 'center',
                  background: '#3b1263',
                  borderRight: '1px solid rgba(255, 255, 255, 0.25)',
                  borderBottom: '1px solid rgba(255, 255, 255, 0.12)'
                }}
              >
                Total
              </th>

              {/* GOVT GSTR-2B 5 Columns */}
              <th
                style={{
                  width: '115px',
                  minWidth: '115px',
                  padding: '6px 10px',
                  textAlign: 'center',
                  background: '#2c0d4a',
                  borderRight: '1px solid rgba(255, 255, 255, 0.10)',
                  borderBottom: '1px solid rgba(255, 255, 255, 0.12)'
                }}
              >
                Taxable
              </th>
              <th
                style={{
                  width: '105px',
                  minWidth: '105px',
                  padding: '6px 10px',
                  textAlign: 'center',
                  background: '#2c0d4a',
                  borderRight: '1px solid rgba(255, 255, 255, 0.10)',
                  borderBottom: '1px solid rgba(255, 255, 255, 0.12)'
                }}
              >
                IGST
              </th>
              <th
                style={{
                  width: '105px',
                  minWidth: '105px',
                  padding: '6px 10px',
                  textAlign: 'center',
                  background: '#2c0d4a',
                  borderRight: '1px solid rgba(255, 255, 255, 0.10)',
                  borderBottom: '1px solid rgba(255, 255, 255, 0.12)'
                }}
              >
                CGST
              </th>
              <th
                style={{
                  width: '105px',
                  minWidth: '105px',
                  padding: '6px 10px',
                  textAlign: 'center',
                  background: '#2c0d4a',
                  borderRight: '1px solid rgba(255, 255, 255, 0.10)',
                  borderBottom: '1px solid rgba(255, 255, 255, 0.12)'
                }}
              >
                SGST
              </th>
              <th
                style={{
                  width: '115px',
                  minWidth: '115px',
                  padding: '6px 10px',
                  textAlign: 'center',
                  background: '#2c0d4a',
                  borderRight: '1px solid rgba(255, 255, 255, 0.10)',
                  borderBottom: '1px solid rgba(255, 255, 255, 0.12)'
                }}
              >
                Total
              </th>
            </tr>
          </thead>
          <tbody>
            {filteredItems.length === 0 ? (
              <tr>
                <td colSpan={20} style={{ textAlign: 'center', padding: '36px', color: '#64748b', fontSize: '14px' }}>
                  No Old ITC invoices match the search or filter criteria.
                </td>
              </tr>
            ) : (
              filteredItems.map((item, idx) => {
                const isClaimed = item.status === 'Matched - Claim';

                // Comparison indicators for Taxable & Tax amounts
                const renderAmt = (bAmt: number, gAmt: number, isBooks: boolean) => {
                  const val = isBooks ? bAmt : gAmt;
                  const other = isBooks ? gAmt : bAmt;
                  let arrow = null;
                  if (other > 0 && Math.abs(val - other) > 0.01) {
                    if (val > other) {
                      arrow = <span style={{ color: '#15803d', fontSize: '10px', marginLeft: '3px' }}>▲</span>;
                    } else {
                      arrow = <span style={{ color: '#b91c1c', fontSize: '10px', marginLeft: '3px' }}>▼</span>;
                    }
                  }
                  return (
                    <span>
                      {val === 0 ? <span style={{ color: '#94a3b8' }}>0.00</span> : formatInr(val)}
                      {arrow}
                    </span>
                  );
                };

                return (
                  <tr
                    key={item.id}
                    style={{
                      background: idx % 2 === 0 ? '#FAF8FD' : '#ffffff',
                      borderBottom: '1px solid #e2e8f0',
                      height: '42px'
                    }}
                    onMouseEnter={(e) => (e.currentTarget.style.background = '#F3EBFE')}
                    onMouseLeave={(e) => (e.currentTarget.style.background = idx % 2 === 0 ? '#FAF8FD' : '#ffffff')}
                  >
                    {/* 1. S.No */}
                    <td
                      style={{
                        width: '55px',
                        minWidth: '55px',
                        padding: '8px 8px',
                        textAlign: 'center',
                        color: '#64748b',
                        fontWeight: 600,
                        borderRight: '1px solid #e2e8f0'
                      }}
                    >
                      {idx + 1}
                    </td>

                    {/* 2. Book Month */}
                    <td
                      style={{
                        width: '105px',
                        minWidth: '105px',
                        padding: '8px 10px',
                        textAlign: 'center',
                        fontWeight: 700,
                        color: '#334155',
                        borderRight: '1px solid #e2e8f0'
                      }}
                    >
                      {item.originalMonth}
                    </td>

                    {/* 3. 2B Month */}
                    <td
                      style={{
                        width: '100px',
                        minWidth: '100px',
                        padding: '8px 10px',
                        textAlign: 'center',
                        fontWeight: 600,
                        color: item.month2b ? '#334155' : '#94a3b8',
                        borderRight: '1px solid #e2e8f0'
                      }}
                    >
                      {item.month2b || '-'}
                    </td>

                    {/* 4. Claim Month */}
                    <td
                      style={{
                        width: '105px',
                        minWidth: '105px',
                        padding: '8px 10px',
                        textAlign: 'center',
                        fontWeight: 700,
                        color: isClaimed ? '#15803d' : '#94a3b8',
                        borderRight: '1px solid #e2e8f0'
                      }}
                    >
                      {isClaimed ? (item.claimMonth || '-') : '-'}
                    </td>

                    {/* 5. Invoice No. */}
                    <td
                      style={{
                        width: '140px',
                        minWidth: '140px',
                        padding: '8px 12px',
                        textAlign: 'left',
                        fontWeight: 700,
                        color: '#1E1747',
                        fontSize: '13px',
                        borderRight: '1px solid #e2e8f0'
                      }}
                    >
                      {item.invNo}
                    </td>

                    {/* 6. Supplier GSTIN */}
                    <td
                      style={{
                        width: '155px',
                        minWidth: '155px',
                        padding: '8px 12px',
                        textAlign: 'left',
                        fontFamily: 'monospace',
                        fontSize: '12px',
                        borderRight: '1px solid #e2e8f0'
                      }}
                    >
                      {item.gstin}
                    </td>

                    {/* 7. Supplier Name */}
                    <td
                      style={{
                        width: '220px',
                        minWidth: '220px',
                        padding: '8px 14px',
                        textAlign: 'left',
                        fontWeight: 600,
                        color: '#1E1747',
                        borderRight: '1px solid #e2e8f0'
                      }}
                    >
                      {item.supplierName}
                    </td>

                    {/* 8. Reason */}
                    <td
                      style={{
                        width: '200px',
                        minWidth: '200px',
                        padding: '8px 12px',
                        textAlign: 'left',
                        fontSize: '12px',
                        color: '#475569',
                        borderRight: '1px solid #e2e8f0'
                      }}
                      title={item.reason}
                    >
                      {item.reason}
                    </td>

                    {/* 9. Status */}
                    <td
                      style={{
                        width: '130px',
                        minWidth: '130px',
                        padding: '8px 10px',
                        textAlign: 'center',
                        borderRight: '1px solid #e2e8f0'
                      }}
                    >
                      <span
                        style={{
                          display: 'inline-block',
                          padding: '2px 8px',
                          fontSize: '11px',
                          fontWeight: 700,
                          borderRadius: '2px',
                          background: isClaimed ? '#dcfce7' : '#fef3c7',
                          color: isClaimed ? '#15803d' : '#b45309',
                          border: isClaimed ? '1px solid #86efac' : '1px solid #fde68a'
                        }}
                      >
                        {item.status}
                      </span>
                    </td>

                    {/* BOOKS 5 Columns (Sky Blue tint #F0F7FF) */}
                    <td style={{ width: '115px', minWidth: '115px', padding: '8px 10px', textAlign: 'right', background: '#F0F7FF', fontWeight: 600, borderRight: '1px solid #e2e8f0' }}>
                      {renderAmt(item.bTaxable, item.gTaxable, true)}
                    </td>
                    <td style={{ width: '105px', minWidth: '105px', padding: '8px 10px', textAlign: 'right', background: '#F0F7FF', fontWeight: 600, borderRight: '1px solid #e2e8f0' }}>
                      {renderAmt(item.bIgst, item.gIgst, true)}
                    </td>
                    <td style={{ width: '105px', minWidth: '105px', padding: '8px 10px', textAlign: 'right', background: '#F0F7FF', fontWeight: 600, borderRight: '1px solid #e2e8f0' }}>
                      {renderAmt(item.bCgst, item.gCgst, true)}
                    </td>
                    <td style={{ width: '105px', minWidth: '105px', padding: '8px 10px', textAlign: 'right', background: '#F0F7FF', fontWeight: 600, borderRight: '1px solid #e2e8f0' }}>
                      {renderAmt(item.bSgst, item.gSgst, true)}
                    </td>
                    <td style={{ width: '115px', minWidth: '115px', padding: '8px 10px', textAlign: 'right', background: '#F0F7FF', fontWeight: 700, color: '#1E1747', borderRight: '1.5px solid #cbd5e1' }}>
                      {renderAmt(item.bTotal, item.gTotal, true)}
                    </td>

                    {/* GOVT 2B 5 Columns (Lavender tint #F6F2FC) */}
                    <td style={{ width: '115px', minWidth: '115px', padding: '8px 10px', textAlign: 'right', background: '#F6F2FC', fontWeight: 600, borderRight: '1px solid #e2e8f0' }}>
                      {renderAmt(item.bTaxable, item.gTaxable, false)}
                    </td>
                    <td style={{ width: '105px', minWidth: '105px', padding: '8px 10px', textAlign: 'right', background: '#F6F2FC', fontWeight: 600, borderRight: '1px solid #e2e8f0' }}>
                      {renderAmt(item.bIgst, item.gIgst, false)}
                    </td>
                    <td style={{ width: '105px', minWidth: '105px', padding: '8px 10px', textAlign: 'right', background: '#F6F2FC', fontWeight: 600, borderRight: '1px solid #e2e8f0' }}>
                      {renderAmt(item.bCgst, item.gCgst, false)}
                    </td>
                    <td style={{ width: '105px', minWidth: '105px', padding: '8px 10px', textAlign: 'right', background: '#F6F2FC', fontWeight: 600, borderRight: '1px solid #e2e8f0' }}>
                      {renderAmt(item.bSgst, item.gSgst, false)}
                    </td>
                    <td style={{ width: '115px', minWidth: '115px', padding: '8px 10px', textAlign: 'right', background: '#F6F2FC', fontWeight: 700, color: '#350b5e', borderRight: '1px solid #e2e8f0' }}>
                      {renderAmt(item.bTotal, item.gTotal, false)}
                    </td>

                    {/* Action Column: View, Audit, Move to Not Claim */}
                    <td
                      style={{
                        width: '180px',
                        minWidth: '180px',
                        padding: '8px 12px',
                        textAlign: 'center',
                        whiteSpace: 'nowrap',
                        borderLeft: '1px solid #e2e8f0',
                        borderRight: '1px solid #e2e8f0'
                      }}
                    >
                      <div style={{ display: 'inline-flex', gap: '4px' }}>
                        <button
                          type="button"
                          onClick={() => setViewItem(item)}
                          style={{
                            background: '#f1f5f9',
                            color: '#334155',
                            border: '1px solid #cbd5e1',
                            padding: '3px 7px',
                            fontSize: '11px',
                            fontWeight: 700,
                            cursor: 'pointer',
                            borderRadius: 0
                          }}
                          title="View Invoice Details"
                        >
                          View
                        </button>
                        <button
                          type="button"
                          onClick={() => handleOpenAudit(item)}
                          style={{
                            background: item.auditStatus === 'Verified' ? '#dcfce7' : '#f1f5f9',
                            color: item.auditStatus === 'Verified' ? '#15803d' : '#334155',
                            border: '1px solid #cbd5e1',
                            padding: '3px 7px',
                            fontSize: '11px',
                            fontWeight: 700,
                            cursor: 'pointer',
                            borderRadius: 0
                          }}
                          title="Add / View Audit Note"
                        >
                          Audit
                        </button>
                        {!isClaimed && (
                          <button
                            type="button"
                            onClick={() => handleMoveToNotClaim(item)}
                            style={{
                              background: '#fef2f2',
                              color: '#b91c1c',
                              border: '1px solid #fecaca',
                              padding: '3px 7px',
                              fontSize: '11px',
                              fontWeight: 700,
                              cursor: 'pointer',
                              borderRadius: 0
                            }}
                            title="Move to Not Claim ITC"
                          >
                            Not Claim
                          </button>
                        )}
                      </div>
                    </td>
                  </tr>
                );
              })
            )}
          </tbody>
        </table>
      </div>

      {/* VIEW MODAL */}
      {viewItem && (
        <div
          style={{
            position: 'fixed',
            inset: 0,
            background: 'rgba(0,0,0,0.5)',
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center',
            zIndex: 1000
          }}
          onClick={() => setViewItem(null)}
        >
          <div
            style={{
              background: '#ffffff',
              width: '90%',
              maxWidth: '650px',
              padding: '24px',
              borderRadius: 0,
              boxShadow: '0 10px 25px rgba(0,0,0,0.2)'
            }}
            onClick={(e) => e.stopPropagation()}
          >
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', borderBottom: '1.5px solid #e2e8f0', paddingBottom: '12px' }}>
              <h3 style={{ margin: 0, fontSize: '18px', color: '#241a3d', fontWeight: 800 }}>
                Invoice Detail: {viewItem.invNo}
              </h3>
              <button
                type="button"
                onClick={() => setViewItem(null)}
                style={{ background: 'transparent', border: 'none', fontSize: '18px', cursor: 'pointer', color: '#64748b' }}
              >
                ✕
              </button>
            </div>

            <div style={{ marginTop: '16px', display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '12px', fontSize: '13px' }}>
              <div><strong>Supplier Name:</strong> {viewItem.supplierName}</div>
              <div><strong>Supplier GSTIN:</strong> <code style={{ color: '#4c1d95' }}>{viewItem.gstin}</code></div>
              <div><strong>Invoice Date:</strong> {viewItem.date}</div>
              <div><strong>Original Month:</strong> {viewItem.originalMonth}</div>
              <div><strong>Current Status:</strong> <span style={{ fontWeight: 700, color: viewItem.status === 'Matched - Claim' ? '#15803d' : '#b45309' }}>{viewItem.status}</span></div>
              <div><strong>Claim Month:</strong> {viewItem.claimMonth || '-'}</div>
              <div><strong>ITC Match Score:</strong> {viewItem.itcMatchRatio}</div>
              <div><strong>Pending Months:</strong> {viewItem.pendingMonthsCount} month(s)</div>
            </div>

            <div style={{ marginTop: '14px', background: '#f8fafc', padding: '10px 12px', border: '1px solid #e2e8f0', fontSize: '12.5px' }}>
              <strong>Mismatch / Reason:</strong>
              <div style={{ color: '#b91c1c', marginTop: '4px', fontWeight: 600 }}>{viewItem.reason}</div>
            </div>

            {/* Side by side Books vs 2B */}
            <div style={{ marginTop: '16px', border: '1px solid #e2e8f0' }}>
              <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', background: '#4A1A7A', color: '#ffffff', padding: '8px 12px', fontWeight: 700, fontSize: '12px' }}>
                <div>BOOKS (PURCHASE REGISTER)</div>
                <div>GOVT GSTR-2B</div>
              </div>
              <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', padding: '12px', gap: '16px', fontSize: '13px' }}>
                <div style={{ background: '#F0F7FF', padding: '10px' }}>
                  <div>Taxable: <strong>₹{formatInr(viewItem.bTaxable)}</strong></div>
                  <div>IGST: <strong>₹{formatInr(viewItem.bIgst)}</strong></div>
                  <div>CGST: <strong>₹{formatInr(viewItem.bCgst)}</strong></div>
                  <div>SGST: <strong>₹{formatInr(viewItem.bSgst)}</strong></div>
                  <div style={{ marginTop: '6px', borderTop: '1px solid #cbd5e1', paddingTop: '4px', color: '#1E1747', fontWeight: 800 }}>
                    Total: ₹{formatInr(viewItem.bTotal)}
                  </div>
                </div>
                <div style={{ background: '#F6F2FC', padding: '10px' }}>
                  <div>Taxable: <strong>₹{formatInr(viewItem.gTaxable)}</strong></div>
                  <div>IGST: <strong>₹{formatInr(viewItem.gIgst)}</strong></div>
                  <div>CGST: <strong>₹{formatInr(viewItem.gCgst)}</strong></div>
                  <div>SGST: <strong>₹{formatInr(viewItem.gSgst)}</strong></div>
                  <div style={{ marginTop: '6px', borderTop: '1px solid #cbd5e1', paddingTop: '4px', color: '#350b5e', fontWeight: 800 }}>
                    Total: ₹{formatInr(viewItem.gTotal)}
                  </div>
                </div>
              </div>
            </div>

            <div style={{ marginTop: '20px', textAlign: 'right' }}>
              <button
                type="button"
                onClick={() => setViewItem(null)}
                style={{
                  background: '#4A1A7A',
                  color: '#ffffff',
                  border: 'none',
                  padding: '8px 20px',
                  fontSize: '13px',
                  fontWeight: 700,
                  cursor: 'pointer',
                  borderRadius: 0
                }}
              >
                Close
              </button>
            </div>
          </div>
        </div>
      )}

      {/* AUDIT MODAL */}
      {auditItem && (
        <div
          style={{
            position: 'fixed',
            inset: 0,
            background: 'rgba(0,0,0,0.5)',
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center',
            zIndex: 1000
          }}
          onClick={() => setAuditItem(null)}
        >
          <div
            style={{
              background: '#ffffff',
              width: '90%',
              maxWidth: '500px',
              padding: '24px',
              borderRadius: 0,
              boxShadow: '0 10px 25px rgba(0,0,0,0.2)'
            }}
            onClick={(e) => e.stopPropagation()}
          >
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', borderBottom: '1.5px solid #e2e8f0', paddingBottom: '12px' }}>
              <h3 style={{ margin: 0, fontSize: '18px', color: '#241a3d', fontWeight: 800 }}>
                Audit Note: {auditItem.invNo}
              </h3>
              <button
                type="button"
                onClick={() => setAuditItem(null)}
                style={{ background: 'transparent', border: 'none', fontSize: '18px', cursor: 'pointer', color: '#64748b' }}
              >
                ✕
              </button>
            </div>

            <div style={{ marginTop: '16px' }}>
              <label style={{ display: 'block', fontSize: '12px', fontWeight: 700, color: '#334155', marginBottom: '6px' }}>
                Audit Status
              </label>
              <select
                value={auditStatus}
                onChange={(e) => setAuditStatus(e.target.value)}
                style={{
                  width: '100%',
                  border: '1px solid #cbd5e1',
                  padding: '8px 12px',
                  fontSize: '13px',
                  borderRadius: 0,
                  background: '#f8fafc'
                }}
              >
                <option value="Pending">Pending Review</option>
                <option value="Verified">Verified / Approved</option>
                <option value="Query">Query / Follow-up Sent</option>
              </select>
            </div>

            <div style={{ marginTop: '16px' }}>
              <label style={{ display: 'block', fontSize: '12px', fontWeight: 700, color: '#334155', marginBottom: '6px' }}>
                Auditor Remark / Note
              </label>
              <textarea
                value={auditNote}
                onChange={(e) => setAuditNote(e.target.value)}
                rows={4}
                placeholder="Write observation, supplier communication details, etc..."
                style={{
                  width: '100%',
                  border: '1px solid #cbd5e1',
                  padding: '8px 12px',
                  fontSize: '13px',
                  borderRadius: 0,
                  outline: 'none',
                  fontFamily: 'inherit'
                }}
              />
            </div>

            <div style={{ marginTop: '20px', display: 'flex', justifyContent: 'flex-end', gap: '10px' }}>
              <button
                type="button"
                onClick={() => setAuditItem(null)}
                style={{
                  background: '#f1f5f9',
                  color: '#334155',
                  border: '1px solid #cbd5e1',
                  padding: '8px 16px',
                  fontSize: '13px',
                  fontWeight: 700,
                  cursor: 'pointer',
                  borderRadius: 0
                }}
              >
                Cancel
              </button>
              <button
                type="button"
                onClick={handleSaveAudit}
                style={{
                  background: '#6D28D9',
                  color: '#ffffff',
                  border: 'none',
                  padding: '8px 20px',
                  fontSize: '13px',
                  fontWeight: 700,
                  cursor: 'pointer',
                  borderRadius: 0
                }}
              >
                Save Audit
              </button>
            </div>
          </div>
        </div>
      )}
    </section>
  );
}
