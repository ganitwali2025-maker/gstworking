import React, { useState, useRef } from 'react';
import * as XLSX from 'xlsx';
import { fm, fmtVal, dmySlash, inr, pp, iso } from '../utils/gstFormat';

export interface ItcPurchaseRecord {
  id: string;
  m: string;
  q: string;
  fy: string;
  d: string;
  p: string; // Name of Supplier
  g: string; // GST No.
  n: string; // Invoice No
  t: number; // BESIC AS PER BOOK (Basic / Taxable value in paise or num)
  i: number; // Integrated Tax
  c: number; // Central Tax
  s: number; // State Tax
  cess: number; // Cess
}

const SEED_ITC_RECORDS: ItcPurchaseRecord[] = [
  {
    id: 'itc-1',
    m: 'Aug',
    q: 'Q2 (Jul-Sep)',
    fy: '2026-27',
    d: '2026-08-02',
    p: 'Bharat Steel & Alloys Pvt Ltd',
    g: '22AABCB4567M1Z2',
    n: 'BS-4421',
    t: 42000000,
    i: 0,
    c: 3780000,
    s: 3780000,
    cess: 0
  },
  {
    id: 'itc-2',
    m: 'Aug',
    q: 'Q2 (Jul-Sep)',
    fy: '2026-27',
    d: '2026-08-05',
    p: 'Shree Krishna Logistics & Spares',
    g: '22AAFCK8812D1Z4',
    n: 'SK-1092',
    t: 28000000,
    i: 0,
    c: 2520000,
    s: 2520000,
    cess: 0
  },
  {
    id: 'itc-3',
    m: 'Aug',
    q: 'Q2 (Jul-Sep)',
    fy: '2026-27',
    d: '2026-08-08',
    p: 'Tata Steel Downstream Products Ltd',
    g: '21AABCT1332N1Z8',
    n: 'TS-9901',
    t: 185000000,
    i: 33300000,
    c: 0,
    s: 0,
    cess: 0
  },
  {
    id: 'itc-4',
    m: 'Aug',
    q: 'Q2 (Jul-Sep)',
    fy: '2026-27',
    d: '2026-08-14',
    p: 'Hindalco Industries Limited',
    g: '23AAACH1245G1Z5',
    n: 'HND-5519',
    t: 96000000,
    i: 17280000,
    c: 0,
    s: 0,
    cess: 0
  }
];

interface ItcRegisterViewProps {
  showToast: (msg: string) => void;
  fy: string;
}

export default function ItcRegisterView({ showToast, fy }: ItcRegisterViewProps) {
  const [records, setRecords] = useState<ItcPurchaseRecord[]>(() => {
    try {
      const saved = localStorage.getItem('itc_register_v1');
      return saved ? JSON.parse(saved) : SEED_ITC_RECORDS;
    } catch {
      return SEED_ITC_RECORDS;
    }
  });

  const [isLocked, setIsLocked] = useState(false);
  const [editingRecord, setEditingRecord] = useState<ItcPurchaseRecord | null>(null);
  const [modalOpen, setModalOpen] = useState(false);
  const [pastRecords, setPastRecords] = useState<ItcPurchaseRecord[][]>([]);
  const [futureRecords, setFutureRecords] = useState<ItcPurchaseRecord[][]>([]);
  const fileInputRef = useRef<HTMLInputElement>(null);

  const pushHistory = (prev: ItcPurchaseRecord[]) => {
    setPastRecords((p) => [...p.slice(-25), prev]);
    setFutureRecords([]);
  };

  const handleUndo = () => {
    if (pastRecords.length === 0) {
      showToast('No more actions to undo');
      return;
    }
    const prev = pastRecords[pastRecords.length - 1];
    setFutureRecords((f) => [records, ...f]);
    setPastRecords((p) => p.slice(0, p.length - 1));
    setRecords(prev);
    try { localStorage.setItem('itc_register_v1', JSON.stringify(prev)); } catch {}
    showToast('Undo successful');
  };

  const handleRedo = () => {
    if (futureRecords.length === 0) {
      showToast('No more actions to redo');
      return;
    }
    const next = futureRecords[0];
    setPastRecords((p) => [...p, records]);
    setFutureRecords((f) => f.slice(1));
    setRecords(next);
    try { localStorage.setItem('itc_register_v1', JSON.stringify(next)); } catch {}
    showToast('Redo successful');
  };

  // Totals for top KPI cards
  const totalBasic = records.reduce((acc, r) => acc + (r.t || 0), 0);
  const totalIgst = records.reduce((acc, r) => acc + (r.i || 0), 0);
  const totalCgst = records.reduce((acc, r) => acc + (r.c || 0), 0);
  const totalSgst = records.reduce((acc, r) => acc + (r.s || 0), 0);
  const totalTax = totalIgst + totalCgst + totalSgst + records.reduce((acc, r) => acc + (r.cess || 0), 0);

  const persistRecords = (recs: ItcPurchaseRecord[]) => {
    pushHistory(records);
    setRecords(recs);
    try {
      localStorage.setItem('itc_register_v1', JSON.stringify(recs));
    } catch {}
  };

  const handleOpenEdit = (rec: ItcPurchaseRecord) => {
    if (isLocked) {
      showToast('Sheet is locked. Unlock to make changes.');
      return;
    }
    setEditingRecord(rec);
    setModalOpen(true);
  };

  const handleSaveRecord = (updated: ItcPurchaseRecord) => {
    persistRecords(records.map((r) => (r.id === updated.id ? updated : r)));
    setModalOpen(false);
    showToast('Record updated successfully');
  };

  const handleClearData = () => {
    if (isLocked) {
      showToast('Sheet is locked. Unlock to clear.');
      return;
    }
    if (window.confirm('Are you sure you want to clear all ITC purchase records?')) {
      persistRecords([]);
      showToast('Data cleared');
    }
  };

  const handleDelete = (id: string) => {
    if (isLocked) {
      showToast('Sheet is locked.');
      return;
    }
    persistRecords(records.filter((r) => r.id !== id));
    showToast('Record deleted');
  };

  const handleFileUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    e.target.value = '';
    if (!file) return;

    try {
      const data = await file.arrayBuffer();
      const wb = XLSX.read(data, { type: 'array' });
      const firstSheet = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json<any>(firstSheet, { defval: '' });

      const imported: ItcPurchaseRecord[] = [];
      for (const row of rows) {
        const d = iso(row['Invoice Date'] || row['Date'] || row['d'] || '');
        const p = String(row['Name of Supplier'] || row['Supplier'] || row['Party'] || '').trim();
        const n = String(row['Invoice No'] || row['Invoice'] || '').trim();
        const g = String(row['GST No.'] || row['GST No'] || row['GSTIN'] || '').trim();
        const t = pp(row['BESIC AS PER BOOK'] || row['Basic'] || row['Taxable Value'] || 0);
        const i = pp(row['Integrated Tax (₹)'] || row['IGST'] || 0);
        const c = pp(row['Central Tax (₹)'] || row['CGST'] || 0);
        const s = pp(row['State Tax (₹)'] || row['SGST'] || 0);
        const cess = pp(row['Cess'] || 0);

        if (!p || !n) continue;

        imported.push({
          id: 'itc-' + Date.now().toString(36) + Math.random().toString(36).slice(2, 6),
          m: 'Aug',
          q: 'Q2 (Jul-Sep)',
          fy: '2026-27',
          d: d || '2026-08-01',
          p,
          g,
          n,
          t: t || 0,
          i: i || 0,
          c: c || 0,
          s: s || 0,
          cess: cess || 0
        });
      }

      if (imported.length > 0) {
        persistRecords([...imported, ...records]);
        showToast(`${imported.length} purchase records imported`);
      } else {
        showToast('No valid rows found to import.');
      }
    } catch (err: any) {
      showToast('Import error: ' + err.message);
    }
  };

  return (
    <section id="v-g1">
      <input
        type="file"
        ref={fileInputRef}
        accept=".xlsx,.xls,.csv"
        style={{ display: 'none' }}
        onChange={handleFileUpload}
      />

      {/* 5 KPI Cards Grid */}
      <div className="gstr1-kpis-grid">
        {/* Card 1: Total Basic */}
        <div className="gstr1-kpi-card">
          <div className="gstr1-kpi-top">
            <div className="gstr1-kpi-icon">
              <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path>
                <polyline points="14 2 14 8 20 8"></polyline>
                <line x1="16" y1="13" x2="8" y2="13"></line>
                <line x1="16" y1="17" x2="8" y2="17"></line>
              </svg>
            </div>
            <div className="gstr1-kpi-chart-icon">
              <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="#9ca3af" strokeWidth="2">
                <path d="M18 20V10M12 20V4M6 20v-6"></path>
              </svg>
            </div>
          </div>
          <div className="gstr1-kpi-label">Total Basic</div>
          <div className="gstr1-kpi-val">{inr(totalBasic / 100)}</div>
        </div>

        {/* Card 2: Total IGST */}
        <div className="gstr1-kpi-card">
          <div className="gstr1-kpi-top">
            <div className="gstr1-kpi-icon">
              <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                <circle cx="12" cy="12" r="10"></circle>
                <line x1="2" y1="12" x2="22" y2="12"></line>
                <path d="M12 2a15.3 15.3 0 0 1 4 10 15.3 15.3 0 0 1-4 10 15.3 15.3 0 0 1-4-10 15.3 15.3 0 0 1 4-10z"></path>
              </svg>
            </div>
            <div className="gstr1-kpi-chart-icon">
              <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="#9ca3af" strokeWidth="2">
                <path d="M18 20V10M12 20V4M6 20v-6"></path>
              </svg>
            </div>
          </div>
          <div className="gstr1-kpi-label">Total IGST</div>
          <div className="gstr1-kpi-val">{inr(totalIgst / 100)}</div>
        </div>

        {/* Card 3: Total CGST */}
        <div className="gstr1-kpi-card">
          <div className="gstr1-kpi-top">
            <div className="gstr1-kpi-icon">
              <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                <line x1="3" y1="21" x2="21" y2="21"></line>
                <line x1="3" y1="10" x2="21" y2="10"></line>
                <polyline points="5 6 12 3 19 6"></polyline>
                <line x1="4" y1="10" x2="4" y2="21"></line>
                <line x1="20" y1="10" x2="20" y2="21"></line>
                <line x1="8" y1="14" x2="8" y2="17"></line>
                <line x1="12" y1="14" x2="12" y2="17"></line>
                <line x1="16" y1="14" x2="16" y2="17"></line>
              </svg>
            </div>
            <div className="gstr1-kpi-chart-icon">
              <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="#9ca3af" strokeWidth="2">
                <path d="M18 20V10M12 20V4M6 20v-6"></path>
              </svg>
            </div>
          </div>
          <div className="gstr1-kpi-label">Total CGST</div>
          <div className="gstr1-kpi-val">{inr(totalCgst / 100)}</div>
        </div>

        {/* Card 4: Total SGST */}
        <div className="gstr1-kpi-card">
          <div className="gstr1-kpi-top">
            <div className="gstr1-kpi-icon">
              <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
                <path d="M21 10c0 7-9 13-9 13s-9-6-9-13a9 9 0 0 1 18 0z"></path>
                <circle cx="12" cy="10" r="3"></circle>
              </svg>
            </div>
            <div className="gstr1-kpi-chart-icon">
              <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="#9ca3af" strokeWidth="2">
                <path d="M18 20V10M12 20V4M6 20v-6"></path>
              </svg>
            </div>
          </div>
          <div className="gstr1-kpi-label">Total SGST</div>
          <div className="gstr1-kpi-val">{inr(totalSgst / 100)}</div>
        </div>

        {/* Card 5: Total Tax */}
        <div className="gstr1-kpi-card">
          <div className="gstr1-kpi-top">
            <div className="gstr1-kpi-icon">
              <span style={{ fontSize: '16px', fontWeight: 700 }}>₹</span>
            </div>
            <div className="gstr1-kpi-chart-icon">
              <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="#9ca3af" strokeWidth="2">
                <path d="M18 20V10M12 20V4M6 20v-6"></path>
              </svg>
            </div>
          </div>
          <div className="gstr1-kpi-label">Total Tax</div>
          <div className="gstr1-kpi-val">{inr(totalTax / 100)}</div>
        </div>
      </div>

      <div className="card gstr1-card">
        {/* Top Header Card */}
        <div className="gstr1-header">
          <div className="gstr1-title-box">
            <div className="gstr1-icon-badge">
              <svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path>
                <polyline points="14 2 14 8 20 8"></polyline>
                <line x1="16" y1="13" x2="8" y2="13"></line>
                <line x1="16" y1="17" x2="8" y2="17"></line>
                <polyline points="10 9 9 9 8 9"></polyline>
              </svg>
            </div>
            <div>
              <h2 className="gstr1-title">Input Tax Credit (ITC) Purchase Register</h2>
            </div>
          </div>

          {/* Action Buttons */}
          <div className="gstr1-btn-group">
            <button
              type="button"
              className="btn-gstr1 btn-lock"
              onClick={() => {
                setIsLocked(!isLocked);
                showToast(isLocked ? 'Sheet unlocked' : 'Sheet locked');
              }}
            >
              <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round">
                {isLocked ? (
                  <>
                    <rect x="3" y="11" width="18" height="11" rx="2" ry="2"></rect>
                    <path d="M7 11V7a5 5 0 0 1 9.9-1"></path>
                  </>
                ) : (
                  <>
                    <rect x="3" y="11" width="18" height="11" rx="2" ry="2"></rect>
                    <path d="M7 11V7a5 5 0 0 1 10 0v4"></path>
                  </>
                )}
              </svg>
              {isLocked ? 'Unlock Sheet' : 'Lock Sheet'}
            </button>

            <button
              type="button"
              className="btn-gstr1 btn-clear"
              disabled={isLocked}
              onClick={handleClearData}
            >
              <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round">
                <polyline points="3 6 5 6 21 6"></polyline>
                <path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"></path>
              </svg>
              Clear Data
            </button>

            <button
              type="button"
              className="btn-gstr1 btn-refresh"
              onClick={() => {
                showToast('Register refreshed');
              }}
            >
              <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round">
                <path d="M21.5 2v6h-6M21.34 15.57a10 10 0 1 1-.57-8.38l5.67-5.67"></path>
              </svg>
              Refresh
            </button>

            <button
              type="button"
              className="btn-gstr1 btn-import"
              disabled={isLocked}
              onClick={() => fileInputRef.current?.click()}
            >
              <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round">
                <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path>
                <polyline points="17 8 12 3 7 8"></polyline>
                <line x1="12" y1="3" x2="12" y2="15"></line>
              </svg>
              Import
            </button>
          </div>
        </div>

        {/* Table matching requested heading row */}
        <div className="gstr1-table-wrapper">
          <table className="gstr1-table">
            <thead>
              <tr>
                <th style={{ textAlign: 'center' }}>MONTH</th>
                <th style={{ textAlign: 'center' }}>QUARTER</th>
                <th style={{ textAlign: 'center' }}>INVOICE DATE</th>
                <th className="col-party" style={{ minWidth: '220px', textAlign: 'left' }}>NAME OF SUPPLIER</th>
                <th style={{ textAlign: 'center' }}>GST NO.</th>
                <th className="col-inv" style={{ textAlign: 'left' }}>INVOICE NO</th>
                <th style={{ textAlign: 'center' }}>BESIC AS PER BOOK</th>
                <th style={{ textAlign: 'center' }}>INTEGRATED TAX (₹)</th>
                <th style={{ textAlign: 'center' }}>CENTRAL TAX (₹)</th>
                <th style={{ textAlign: 'center' }}>STATE TAX (₹)</th>
                <th style={{ textAlign: 'center' }}>CESS</th>
                <th style={{ textAlign: 'center' }}>TOTAL TAX AMOUNT (₹)</th>
                <th style={{ textAlign: 'center' }}>TOTAL INVOICE VALUE (₹)</th>
                <th style={{ textAlign: 'center' }}>ACTION</th>
              </tr>
            </thead>
            <tbody>
              {records.length > 0 ? (
                records.map((r) => {
                  const totalTax = (r.i || 0) + (r.c || 0) + (r.s || 0) + (r.cess || 0);
                  const totalInvVal = (r.t || 0) + totalTax;

                  return (
                    <tr key={r.id}>
                      <td style={{ textAlign: 'center' }}>{r.m}</td>
                      <td style={{ textAlign: 'center' }}>{r.q}</td>
                      <td style={{ textAlign: 'center' }}>{dmySlash(r.d)}</td>
                      <td className="party-name" style={{ textAlign: 'left' }} title={r.p}>
                        {r.p && r.p.length > 34 ? r.p.slice(0, 34) + '...' : r.p}
                      </td>
                      <td className="gst-no" style={{ textAlign: 'center' }}>{r.g || '—'}</td>
                      <td className="inv-no" style={{ textAlign: 'left' }}>{r.n}</td>
                      <td className="num-val" style={{ textAlign: 'center' }}>{fmtVal(r.t)}</td>
                      <td className="num-val" style={{ textAlign: 'center' }}>{fmtVal(r.i)}</td>
                      <td className="num-val" style={{ textAlign: 'center' }}>{fmtVal(r.c)}</td>
                      <td className="num-val" style={{ textAlign: 'center' }}>{fmtVal(r.s)}</td>
                      <td className="num-val" style={{ textAlign: 'center' }}>{fmtVal(r.cess)}</td>
                      <td className="num-val total-tax" style={{ textAlign: 'center' }}>{fmtVal(totalTax)}</td>
                      <td className="num-val" style={{ fontWeight: 600, textAlign: 'center' }}>{fmtVal(totalInvVal)}</td>
                      <td className="act-col" style={{ textAlign: 'center', whiteSpace: 'nowrap' }}>
                        <button
                          type="button"
                          disabled={isLocked}
                          onClick={() => handleOpenEdit(r)}
                          className="row-edit-btn"
                          title="Edit record"
                        >
                          Edit
                        </button>
                        <button
                          type="button"
                          disabled={isLocked}
                          onClick={() => handleDelete(r.id)}
                          className="row-delete-btn"
                          title="Delete invoice"
                        >
                          Delete
                        </button>
                      </td>
                    </tr>
                  );
                })
              ) : (
                <tr>
                  <td colSpan={14} style={{ textAlign: 'center', padding: '36px', color: '#64748b' }}>
                    No rows found. Click “Import” to load records.
                  </td>
                </tr>
              )}
            </tbody>
          </table>
        </div>
      </div>

      {/* Edit ITC Modal */}
      {modalOpen && editingRecord && (
        <ItcEditModal
          isOpen={modalOpen}
          record={editingRecord}
          isLocked={isLocked}
          onClose={() => setModalOpen(false)}
          onSave={handleSaveRecord}
        />
      )}
    </section>
  );
}

interface ItcEditModalProps {
  isOpen: boolean;
  record: ItcPurchaseRecord;
  isLocked: boolean;
  onClose: () => void;
  onSave: (record: ItcPurchaseRecord) => void;
}

function ItcEditModal({ isOpen, record, isLocked, onClose, onSave }: ItcEditModalProps) {
  if (!isOpen) return null;

  const [d, setD] = useState(record.d);
  const [p, setP] = useState(record.p);
  const [g, setG] = useState(record.g);
  const [n, setN] = useState(record.n);
  const [t, setT] = useState((record.t / 100).toString());
  const [i, setI] = useState((record.i / 100).toString());
  const [c, setC] = useState((record.c / 100).toString());
  const [s, setS] = useState((record.s / 100).toString());
  const [cess, setCess] = useState((record.cess / 100).toString());
  const [error, setError] = useState('');

  const tVal = pp(t) || 0;
  const iVal = pp(i) || 0;
  const cVal = pp(c) || 0;
  const sVal = pp(s) || 0;
  const cessVal = pp(cess) || 0;
  const totalTax = iVal + cVal + sVal + cessVal;
  const totalInv = tVal + totalTax;

  const handleSubmit = () => {
    if (isLocked) {
      setError('Sheet is locked.');
      return;
    }
    if (!d) return setError('Enter invoice date.');
    if (!n.trim()) return setError('Enter invoice number.');
    if (!p.trim()) return setError('Enter supplier name.');

    onSave({
      ...record,
      d,
      p: p.trim(),
      g: (g || '').trim().toUpperCase(),
      n: n.trim(),
      t: tVal,
      i: iVal,
      c: cVal,
      s: sVal,
      cess: cessVal
    });
  };

  return (
    <div className="mo">
      <div className="mb">
        <h3>Edit ITC Purchase Record</h3>
        <div style={{ color: 'var(--mute)', fontSize: '13px', marginBottom: '12px' }}>
          Period: {record.m} | {record.fy}
        </div>

        <div className="fg">
          <div>
            <label>Invoice Date</label>
            <input type="date" value={d} onChange={(e) => setD(e.target.value)} />
          </div>
          <div>
            <label>Invoice No</label>
            <input value={n} onChange={(e) => setN(e.target.value)} />
          </div>
          <div>
            <label>Name of Supplier</label>
            <input value={p} onChange={(e) => setP(e.target.value)} />
          </div>
          <div>
            <label>GST No.</label>
            <input maxLength={15} value={g} onChange={(e) => setG(e.target.value.toUpperCase())} />
          </div>
          <div>
            <label>BESIC AS PER BOOK (₹)</label>
            <input inputMode="decimal" value={t} onChange={(e) => setT(e.target.value)} />
          </div>
          <div>
            <label>Integrated Tax (₹)</label>
            <input inputMode="decimal" value={i} onChange={(e) => setI(e.target.value)} />
          </div>
          <div>
            <label>Central Tax (₹)</label>
            <input inputMode="decimal" value={c} onChange={(e) => setC(e.target.value)} />
          </div>
          <div>
            <label>State Tax (₹)</label>
            <input inputMode="decimal" value={s} onChange={(e) => setS(e.target.value)} />
          </div>
          <div>
            <label>Cess (₹)</label>
            <input inputMode="decimal" value={cess} onChange={(e) => setCess(e.target.value)} />
          </div>
        </div>

        <div style={{ marginTop: '14px', display: 'flex', gap: '24px', flexWrap: 'wrap' }}>
          <span>Total Tax: <b>{fm(totalTax)}</b></span>
          <span>Total Invoice Value: <b>{fm(totalInv)}</b></span>
        </div>

        {error && <div className="fe">{error}</div>}

        <div className="fr">
          <button type="button" onClick={onClose}>Cancel</button>
          <button type="button" className="pri" onClick={handleSubmit}>Save Changes</button>
        </div>
      </div>
    </div>
  );
}
