import React, { useState, useMemo, useEffect } from 'react';
import * as XLSX from 'xlsx';
import {
  syncOldItcPool,
  moveBackToOldItc,
  updateOldItcAudit,
  OldItcItem
} from '../services/oldItcService';

interface NotClaimItcDetailSheetProps {
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

export default function NotClaimItcDetailSheet({
  onBack,
  showToast,
  workingMonth = 'Oct-2026'
}: NotClaimItcDetailSheetProps) {
  const [searchTerm, setSearchTerm] = useState('');
  const [monthFilter, setMonthFilter] = useState('ALL');
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

  // Filter Not Claim items
  const filteredItems = useMemo(() => {
    return syncResult.notClaimItems.filter((item) => {
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
        if (!matchName && !matchGst && !matchInv && !matchReason) {
          return false;
        }
      }

      return true;
    });
  }, [syncResult.notClaimItems, monthFilter, searchTerm]);

  // Unique months available
  const availableMonths = useMemo(() => {
    const set = new Set<string>();
    syncResult.notClaimItems.forEach((i) => {
      if (i.originalMonth) set.add(i.originalMonth);
    });
    return Array.from(set);
  }, [syncResult.notClaimItems]);

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

  // Move back to Old ITC handler
  const handleMoveBack = (item: OldItcItem) => {
    if (window.confirm(`Move Invoice "${item.invNo}" back to Old ITC pool for re-checking?`)) {
      moveBackToOldItc(item.id);
      showToast(`Invoice "${item.invNo}" moved back to Old ITC.`);
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

  // Helper function to format Reason according to user specification
  const getDisplayReason = (item: OldItcItem) => {
    if (item.mismatchFields && item.mismatchFields.length > 0 && !item.mismatchFields.includes('GSTR-2B Missing')) {
      return `Wrong detail in 2B: ${item.mismatchFields.join(', ')}`;
    }
    if (item.reason && item.reason.includes('Wrong detail in 2B:')) {
      return item.reason.replace(/\s*\/\s*/g, ', ');
    }
    if (!item.month2b || (item.reason && item.reason.includes('Not in 2B'))) {
      return 'Not in 2B: supplier ne return file nahi kiya (possible quarterly filer)';
    }
    return item.reason || 'Not in 2B: supplier ne return file nahi kiya (possible quarterly filer)';
  };

  // Export to Excel
  const handleExportExcel = () => {
    try {
      const exportData = filteredItems.map((r, i) => ({
        'S.No': i + 1,
        'Original Month': r.originalMonth,
        'Pending': `${r.pendingMonthsCount} Months`,
        'Reason': getDisplayReason(r),
        'Date': r.date,
        'Invoice No.': r.invNo,
        'Supplier GSTIN': r.gstin,
        'Supplier Name': r.supplierName,
        'Books Taxable (₹)': r.bTaxable,
        'Books IGST (₹)': r.bIgst,
        'Books CGST (₹)': r.bCgst,
        'Books SGST (₹)': r.bSgst,
        'Books Total (₹)': r.bTotal
      }));

      const ws = XLSX.utils.json_to_sheet(exportData);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, 'Not_Claim_ITC');
      XLSX.writeFile(wb, `Not_Claim_ITC_${new Date().toISOString().slice(0, 10)}.xlsx`);
      showToast('Not Claim ITC Excel downloaded successfully!');
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
            Not Claim ITC
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
              Invoices At Risk
            </div>
            <div style={{ fontSize: '22px', fontWeight: 800, color: '#7F1D1D', marginTop: '3px' }}>
              {syncResult.notClaimCount}
            </div>
          </div>

          <div>
            <div style={{ fontSize: '11px', fontWeight: 700, color: '#64748b', textTransform: 'uppercase', letterSpacing: '0.04em' }}>
              Taxable At Risk
            </div>
            <div style={{ fontSize: '19px', fontWeight: 800, color: '#241a3d', marginTop: '4px' }}>
              ₹{formatInr(summaryValues.taxable)}
            </div>
          </div>

          <div>
            <div style={{ fontSize: '11px', fontWeight: 700, color: '#64748b', textTransform: 'uppercase', letterSpacing: '0.04em' }}>
              Total ITC At Risk
            </div>
            <div style={{ fontSize: '21px', fontWeight: 800, color: '#7F1D1D', marginTop: '3px' }}>
              ₹{formatInr(syncResult.notClaimTotalItc)}
            </div>
          </div>

          <div>
            <div style={{ fontSize: '11px', fontWeight: 700, color: '#64748b', textTransform: 'uppercase', letterSpacing: '0.04em' }}>
              Suppliers At Risk
            </div>
            <div style={{ fontSize: '22px', fontWeight: 800, color: '#241a3d', marginTop: '3px' }}>
              {syncResult.supplierRiskSummary.length}
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
          <div style={{ fontSize: '11.5px', fontWeight: 700, color: '#64748b' }}>IGST at Risk</div>
          <div style={{ fontSize: '16px', fontWeight: 800, color: '#241a3d', marginTop: '4px' }}>
            ₹{formatInr(summaryValues.igst)}
          </div>
        </div>

        <div style={{ background: '#ffffff', border: '1.5px solid #e2e8f0', padding: '10px 14px', borderRadius: 0 }}>
          <div style={{ fontSize: '11.5px', fontWeight: 700, color: '#64748b' }}>CGST at Risk</div>
          <div style={{ fontSize: '16px', fontWeight: 800, color: '#241a3d', marginTop: '4px' }}>
            ₹{formatInr(summaryValues.cgst)}
          </div>
        </div>

        <div style={{ background: '#ffffff', border: '1.5px solid #e2e8f0', padding: '10px 14px', borderRadius: 0 }}>
          <div style={{ fontSize: '11.5px', fontWeight: 700, color: '#64748b' }}>SGST at Risk</div>
          <div style={{ fontSize: '16px', fontWeight: 800, color: '#241a3d', marginTop: '4px' }}>
            ₹{formatInr(summaryValues.sgst)}
          </div>
        </div>

        <div style={{ background: '#ffffff', border: '1.5px solid #e2e8f0', padding: '10px 14px', borderRadius: 0 }}>
          <div style={{ fontSize: '11.5px', fontWeight: 700, color: '#64748b' }}>Total ITC at Risk</div>
          <div style={{ fontSize: '16px', fontWeight: 800, color: '#7F1D1D', marginTop: '4px' }}>
            ₹{formatInr(summaryValues.totalTax)}
          </div>
        </div>
      </div>

      {/* 4. Controls: Search, Filter & Excel Export */}
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
            <option value="ALL">All Original Months</option>
            {availableMonths.map((m) => (
              <option key={m} value={m}>{m}</option>
            ))}
          </select>
        </div>

        <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
          <span style={{ fontSize: '12px', color: '#64748b', fontWeight: 600 }}>
            Showing {filteredItems.length} of {syncResult.notClaimItems.length} records
          </span>
          <button
            type="button"
            onClick={handleExportExcel}
            style={{
              display: 'inline-flex',
              alignItems: 'center',
              gap: '6px',
              background: '#7F1D1D',
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

      {/* 5. Main Not Claim Table */}
      <div className="detail-sheet-table-wrapper" style={{ overflowX: 'auto', border: '1.5px solid #e2e8f0', background: '#ffffff' }}>
        <table className="detail-sheet-table" style={{ width: '100%', borderCollapse: 'collapse', fontSize: '13px', color: '#1E1747' }}>
          <thead>
            {/* Top Multi-Header Row */}
            <tr style={{ background: '#2B0A4F', color: '#ffffff', textAlign: 'center' }}>
              <th
                colSpan={8}
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
                Invoice &amp; Variance Details
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

              {/* 2. Original Month */}
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
                Original Month
              </th>

              {/* 3. Pending */}
              <th
                style={{
                  width: '90px',
                  minWidth: '90px',
                  padding: '6px 8px',
                  textAlign: 'center',
                  background: '#4A1A7A',
                  borderRight: '1px solid rgba(255, 255, 255, 0.10)',
                  borderBottom: '1px solid rgba(255, 255, 255, 0.12)'
                }}
              >
                Pending
              </th>

              {/* 4. Reason */}
              <th
                style={{
                  width: '250px',
                  minWidth: '250px',
                  padding: '6px 12px',
                  textAlign: 'left',
                  background: '#4A1A7A',
                  borderRight: '1px solid rgba(255, 255, 255, 0.10)',
                  borderBottom: '1px solid rgba(255, 255, 255, 0.12)'
                }}
              >
                Reason
              </th>

              {/* 5. Date */}
              <th
                style={{
                  width: '100px',
                  minWidth: '100px',
                  padding: '6px 10px',
                  textAlign: 'center',
                  background: '#4A1A7A',
                  borderRight: '1px solid rgba(255, 255, 255, 0.10)',
                  borderBottom: '1px solid rgba(255, 255, 255, 0.12)'
                }}
              >
                Date
              </th>

              {/* 6. Invoice No. */}
              <th
                style={{
                  width: '140px',
                  minWidth: '140px',
                  padding: '6px 12px',
                  textAlign: 'left',
                  background: '#4A1A7A',
                  borderRight: '1px solid rgba(255, 255, 255, 0.10)',
                  borderBottom: '1px solid rgba(255, 255, 255, 0.12)'
                }}
              >
                Invoice No.
              </th>

              {/* 7. Supplier GSTIN */}
              <th
                style={{
                  width: '155px',
                  minWidth: '155px',
                  padding: '6px 12px',
                  textAlign: 'left',
                  background: '#4A1A7A',
                  borderRight: '1px solid rgba(255, 255, 255, 0.10)',
                  borderBottom: '1px solid rgba(255, 255, 255, 0.12)'
                }}
              >
                Supplier GSTIN
              </th>

              {/* 8. Supplier Name */}
              <th
                style={{
                  width: '220px',
                  minWidth: '220px',
                  padding: '6px 14px',
                  textAlign: 'left',
                  background: '#4A1A7A',
                  borderRight: '1px solid rgba(255, 255, 255, 0.12)',
                  borderBottom: '1px solid rgba(255, 255, 255, 0.12)'
                }}
              >
                Supplier Name
              </th>

              {/* BOOKS (Purchase Register) 5 Columns */}
              {/* 9. Taxable */}
              <th
                style={{
                  width: '115px',
                  minWidth: '115px',
                  padding: '6px 10px',
                  textAlign: 'right',
                  background: '#3b1263',
                  borderRight: '1px solid rgba(255, 255, 255, 0.10)',
                  borderBottom: '1px solid rgba(255, 255, 255, 0.12)'
                }}
              >
                Taxable
              </th>

              {/* 10. IGST */}
              <th
                style={{
                  width: '105px',
                  minWidth: '105px',
                  padding: '6px 10px',
                  textAlign: 'right',
                  background: '#3b1263',
                  borderRight: '1px solid rgba(255, 255, 255, 0.10)',
                  borderBottom: '1px solid rgba(255, 255, 255, 0.12)'
                }}
              >
                IGST
              </th>

              {/* 11. CGST */}
              <th
                style={{
                  width: '105px',
                  minWidth: '105px',
                  padding: '6px 10px',
                  textAlign: 'right',
                  background: '#3b1263',
                  borderRight: '1px solid rgba(255, 255, 255, 0.10)',
                  borderBottom: '1px solid rgba(255, 255, 255, 0.12)'
                }}
              >
                CGST
              </th>

              {/* 12. SGST */}
              <th
                style={{
                  width: '105px',
                  minWidth: '105px',
                  padding: '6px 10px',
                  textAlign: 'right',
                  background: '#3b1263',
                  borderRight: '1px solid rgba(255, 255, 255, 0.10)',
                  borderBottom: '1px solid rgba(255, 255, 255, 0.12)'
                }}
              >
                SGST
              </th>

              {/* 13. Total */}
              <th
                style={{
                  width: '115px',
                  minWidth: '115px',
                  padding: '6px 10px',
                  textAlign: 'right',
                  background: '#3b1263',
                  borderRight: '1px solid rgba(255, 255, 255, 0.25)',
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
                <td colSpan={14} style={{ textAlign: 'center', padding: '36px', color: '#64748b', fontSize: '14px' }}>
                  No Not Claim ITC invoices found for this period.
                </td>
              </tr>
            ) : (
              filteredItems.map((item, idx) => {
                const displayReason = getDisplayReason(item);

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

                    {/* 2. Original Month */}
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

                    {/* 3. Pending */}
                    <td
                      style={{
                        width: '90px',
                        minWidth: '90px',
                        padding: '8px 8px',
                        textAlign: 'center',
                        color: '#b91c1c',
                        fontWeight: 700,
                        fontSize: '12px',
                        borderRight: '1px solid #e2e8f0'
                      }}
                    >
                      {item.pendingMonthsCount} mos
                    </td>

                    {/* 4. Reason */}
                    <td
                      style={{
                        width: '250px',
                        minWidth: '250px',
                        padding: '8px 12px',
                        textAlign: 'left',
                        fontSize: '12px',
                        color: '#b91c1c',
                        fontWeight: 600,
                        whiteSpace: 'normal',
                        wordBreak: 'break-word',
                        lineHeight: 1.4,
                        borderRight: '1px solid #e2e8f0'
                      }}
                      title={displayReason}
                    >
                      {displayReason}
                    </td>

                    {/* 5. Date */}
                    <td
                      style={{
                        width: '100px',
                        minWidth: '100px',
                        padding: '8px 10px',
                        textAlign: 'center',
                        whiteSpace: 'nowrap',
                        color: '#475569',
                        fontSize: '12px',
                        borderRight: '1px solid #e2e8f0'
                      }}
                    >
                      {item.date}
                    </td>

                    {/* 6. Invoice No. */}
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

                    {/* 7. Supplier GSTIN */}
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

                    {/* 8. Supplier Name */}
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

                    {/* BOOKS (Purchase Register) 5 Columns (Sky Blue tint #F0F7FF) */}
                    {/* 9. Taxable */}
                    <td
                      style={{
                        width: '115px',
                        minWidth: '115px',
                        padding: '8px 10px',
                        textAlign: 'right',
                        background: '#F0F7FF',
                        fontWeight: 600,
                        color: '#1E1747',
                        borderRight: '1px solid #e2e8f0'
                      }}
                    >
                      {item.bTaxable === 0 ? <span style={{ color: '#94a3b8' }}>0.00</span> : formatInr(item.bTaxable)}
                    </td>

                    {/* 10. IGST */}
                    <td
                      style={{
                        width: '105px',
                        minWidth: '105px',
                        padding: '8px 10px',
                        textAlign: 'right',
                        background: '#F0F7FF',
                        fontWeight: 600,
                        color: '#1E1747',
                        borderRight: '1px solid #e2e8f0'
                      }}
                    >
                      {item.bIgst === 0 ? <span style={{ color: '#94a3b8' }}>0.00</span> : formatInr(item.bIgst)}
                    </td>

                    {/* 11. CGST */}
                    <td
                      style={{
                        width: '105px',
                        minWidth: '105px',
                        padding: '8px 10px',
                        textAlign: 'right',
                        background: '#F0F7FF',
                        fontWeight: 600,
                        color: '#1E1747',
                        borderRight: '1px solid #e2e8f0'
                      }}
                    >
                      {item.bCgst === 0 ? <span style={{ color: '#94a3b8' }}>0.00</span> : formatInr(item.bCgst)}
                    </td>

                    {/* 12. SGST */}
                    <td
                      style={{
                        width: '105px',
                        minWidth: '105px',
                        padding: '8px 10px',
                        textAlign: 'right',
                        background: '#F0F7FF',
                        fontWeight: 600,
                        color: '#1E1747',
                        borderRight: '1px solid #e2e8f0'
                      }}
                    >
                      {item.bSgst === 0 ? <span style={{ color: '#94a3b8' }}>0.00</span> : formatInr(item.bSgst)}
                    </td>

                    {/* 13. Total */}
                    <td
                      style={{
                        width: '115px',
                        minWidth: '115px',
                        padding: '8px 10px',
                        textAlign: 'right',
                        background: '#F0F7FF',
                        fontWeight: 700,
                        color: '#7F1D1D',
                        borderRight: '1px solid #e2e8f0'
                      }}
                    >
                      {formatInr(item.bTotal)}
                    </td>

                    {/* 14. Action Column: View, Audit, Move back to Old ITC */}
                    <td
                      style={{
                        width: '180px',
                        minWidth: '180px',
                        padding: '8px 12px',
                        textAlign: 'center',
                        whiteSpace: 'nowrap'
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
                        <button
                          type="button"
                          onClick={() => handleMoveBack(item)}
                          style={{
                            background: '#f5f3ff',
                            color: '#6D28D9',
                            border: '1px solid #ddd6fe',
                            padding: '3px 7px',
                            fontSize: '11px',
                            fontWeight: 700,
                            cursor: 'pointer',
                            borderRadius: 0
                          }}
                          title="Restore back to Old ITC"
                        >
                          To Old ITC
                        </button>
                      </div>
                    </td>
                  </tr>
                );
              })
            )}
          </tbody>
          {filteredItems.length > 0 && (
            <tfoot>
              <tr
                style={{
                  background: '#f8fafc',
                  fontWeight: 800,
                  borderTop: '2px solid #cbd5e1',
                  fontFamily: "'Bricolage Grotesque', sans-serif"
                }}
              >
                <td
                  colSpan={8}
                  style={{
                    padding: '10px 12px',
                    textAlign: 'left',
                    color: '#1E1747',
                    letterSpacing: '0.05em',
                    borderRight: '1px solid #cbd5e1'
                  }}
                >
                  TOTAL ({filteredItems.length} INVOICES)
                </td>
                <td
                  style={{
                    padding: '10px 10px',
                    textAlign: 'right',
                    background: '#e0f2fe',
                    color: '#0369a1',
                    borderRight: '1px solid #cbd5e1'
                  }}
                >
                  {formatInr(summaryValues.taxable)}
                </td>
                <td
                  style={{
                    padding: '10px 10px',
                    textAlign: 'right',
                    background: '#e0f2fe',
                    color: '#0369a1',
                    borderRight: '1px solid #cbd5e1'
                  }}
                >
                  {formatInr(summaryValues.igst)}
                </td>
                <td
                  style={{
                    padding: '10px 10px',
                    textAlign: 'right',
                    background: '#e0f2fe',
                    color: '#0369a1',
                    borderRight: '1px solid #cbd5e1'
                  }}
                >
                  {formatInr(summaryValues.cgst)}
                </td>
                <td
                  style={{
                    padding: '10px 10px',
                    textAlign: 'right',
                    background: '#e0f2fe',
                    color: '#0369a1',
                    borderRight: '1px solid #cbd5e1'
                  }}
                >
                  {formatInr(summaryValues.sgst)}
                </td>
                <td
                  style={{
                    padding: '10px 10px',
                    textAlign: 'right',
                    background: '#e0f2fe',
                    color: '#b91c1c',
                    borderRight: '1px solid #cbd5e1'
                  }}
                >
                  {formatInr(summaryValues.grandTotal)}
                </td>
                <td style={{ padding: '10px 12px' }}></td>
              </tr>
            </tfoot>
          )}
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
              <div><strong>Pending Duration:</strong> {viewItem.pendingMonthsCount} month(s)</div>
              <div><strong>ITC at Risk:</strong> <span style={{ fontWeight: 800, color: '#7F1D1D' }}>₹{formatInr(viewItem.bIgst + viewItem.bCgst + viewItem.bSgst)}</span></div>
            </div>

            <div style={{ marginTop: '14px', background: '#fef2f2', padding: '10px 12px', border: '1px solid #fecaca', fontSize: '12.5px' }}>
              <strong>Risk Reason:</strong>
              <div style={{ color: '#b91c1c', marginTop: '4px', fontWeight: 700 }}>{getDisplayReason(viewItem)}</div>
            </div>

            {/* Books Tax Details */}
            <div style={{ marginTop: '16px', border: '1px solid #e2e8f0' }}>
              <div style={{ background: '#2B0A4F', color: '#ffffff', padding: '8px 12px', fontWeight: 700, fontSize: '12px', letterSpacing: '0.04em' }}>
                BOOKS (PURCHASE REGISTER) TAX DETAILS
              </div>
              <div style={{ padding: '14px', fontSize: '13px', background: '#F0F7FF' }}>
                <div style={{ display: 'grid', gridTemplateColumns: 'repeat(4, 1fr)', gap: '12px' }}>
                  <div>Taxable: <strong>₹{formatInr(viewItem.bTaxable)}</strong></div>
                  <div>IGST: <strong>₹{formatInr(viewItem.bIgst)}</strong></div>
                  <div>CGST: <strong>₹{formatInr(viewItem.bCgst)}</strong></div>
                  <div>SGST: <strong>₹{formatInr(viewItem.bSgst)}</strong></div>
                </div>
                <div style={{ marginTop: '10px', borderTop: '1px solid #cbd5e1', paddingTop: '8px', color: '#7F1D1D', fontWeight: 800, fontSize: '14px' }}>
                  Total Invoice Value: ₹{formatInr(viewItem.bTotal)}
                </div>
              </div>
            </div>

            <div style={{ marginTop: '20px', textAlign: 'right' }}>
              <button
                type="button"
                onClick={() => setViewItem(null)}
                style={{
                  background: '#7F1D1D',
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
                <option value="Verified">Verified / Supplier Contacted</option>
                <option value="Query">Query Raised with Supplier</option>
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
                placeholder="Write follow-up notes, date supplier promised amendment, etc..."
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
                  background: '#7F1D1D',
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
