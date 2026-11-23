import { useState } from 'react';
import { InvoiceRecord } from '../types/gst';
import { pp, fm } from '../utils/gstFormat';

interface ModalProps {
  isOpen: boolean;
  invoice: InvoiceRecord | null;
  isLocked: boolean;
  fy: string;
  defaultMonth?: string;
  defaultFy?: string;
  onClose: () => void;
  onSave: (record: InvoiceRecord) => void;
}

export default function Modal({
  isOpen,
  invoice,
  isLocked,
  fy,
  defaultMonth,
  defaultFy,
  onClose,
  onSave
}: ModalProps) {
  if (!isOpen) return null;

  const initialDate = invoice
    ? invoice.d
    : defaultMonth && defaultMonth !== 'ALL'
      ? (() => {
          const MN = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
          const mi = MN.findIndex(x => x.toLowerCase() === defaultMonth.toLowerCase().slice(0, 3));
          const mm = mi !== -1 ? String(mi + 1).padStart(2, '0') : '08';
          const yr = (defaultFy && defaultFy !== 'ALL' ? defaultFy.replace(/[^0-9]/g, '').slice(0, 4) : '2026') || '2026';
          return `${yr}-${mm}-01`;
        })()
      : new Date().toISOString().slice(0, 10);

  const [date, setDate] = useState(initialDate);
  const [invoiceNo, setInvoiceNo] = useState(invoice ? invoice.n : '');
  const [partyName, setPartyName] = useState(invoice ? invoice.p : '');
  const [gstNo, setGstNo] = useState(invoice ? invoice.g : '');

  const [taxable, setTaxable] = useState(invoice ? (invoice.t / 100).toFixed(2) : '');
  const [igst, setIgst] = useState(invoice ? (invoice.i / 100).toFixed(2) : '');
  const [cgst, setCgst] = useState(invoice ? (invoice.c / 100).toFixed(2) : '');
  const [sgst, setSgst] = useState(invoice ? (invoice.s / 100).toFixed(2) : '');

  const [error, setError] = useState('');

  const tVal = pp(taxable) || 0;
  const iVal = pp(igst) || 0;
  const cVal = pp(cgst) || 0;
  const sVal = pp(sgst) || 0;

  const totalTax = iVal + cVal + sVal;
  const totalInvoice = tVal + totalTax;

  const handleSubmit = () => {
    if (isLocked) {
      setError('Sheet is locked. Unlock to make changes.');
      return;
    }

    if (!date) return setError('Enter the invoice date.');
    if (!invoiceNo.trim()) return setError('Enter the invoice number.');
    if (!partyName.trim()) return setError('Enter the supplier / party name.');

    const cleanGst = (gstNo || '').trim().toUpperCase();
    if (cleanGst && !/^[0-9]{2}[A-Z]{5}[0-9]{4}[A-Z][1-9A-Z]Z[0-9A-Z]$/.test(cleanGst)) {
      return setError('Enter a valid 15-character GST No, or leave blank for unregistered.');
    }

    if (isNaN(tVal) || isNaN(iVal) || isNaN(cVal) || isNaN(sVal)) {
      return setError('Amounts must be valid numbers (e.g. 12500.50).');
    }

    if (iVal > 0 && (cVal > 0 || sVal > 0)) {
      return setError('An invoice cannot have IGST together with CGST/SGST.');
    }

    const MN = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
    const QM_LABEL: Record<string, string> = {
      Q1: 'Apr-Jun',
      Q2: 'Jul-Sep',
      Q3: 'Oct-Dec',
      Q4: 'Jan-Mar'
    };
    const qOf = (mi: number): string => (mi >= 3 && mi <= 5 ? 'Q1' : mi >= 6 && mi <= 8 ? 'Q2' : mi >= 9 && mi <= 11 ? 'Q3' : 'Q4');

    const dateParts = String(date || '').split('-').map(Number);
    const dateMi = (dateParts.length >= 2 && !isNaN(dateParts[1])) ? dateParts[1] - 1 : 7;
    const computedMonth = MN[dateMi] || 'Aug';
    const computedQ = qOf(dateMi);
    const computedQFormatted = `${computedQ} (${QM_LABEL[computedQ] || ''})`;

    const saveFy = invoice ? invoice.fy : (defaultFy && defaultFy !== 'ALL' ? defaultFy : fy);
    const saveMonth = invoice ? invoice.m : (defaultMonth && defaultMonth !== 'ALL' ? defaultMonth.slice(0, 3) : computedMonth);
    const saveQ = invoice ? invoice.q : computedQFormatted;

    const newRecord: InvoiceRecord = {
      id: invoice ? invoice.id : 'inv-' + Date.now().toString(36) + Math.random().toString(36).slice(2, 6),
      fy: saveFy,
      q: saveQ,
      m: saveMonth,
      d: date,
      n: invoiceNo.trim(),
      p: partyName.trim(),
      g: cleanGst,
      t: tVal,
      i: iVal,
      c: cVal,
      s: sVal
    };

    onSave(newRecord);
  };

  return (
    <div className="mo">
      <div className="mb">
        <h3>{invoice ? 'Edit Invoice' : 'Add New Invoice'}</h3>
        <div style={{ color: 'var(--mute)', fontSize: '13px', marginBottom: '12px' }}>
          Period: Financial Year {invoice ? invoice.fy : (defaultFy && defaultFy !== 'ALL' ? defaultFy : fy)} | Month: {invoice ? invoice.m : (defaultMonth && defaultMonth !== 'ALL' ? defaultMonth : 'Auto')}
        </div>

        <div className="fg">
          <div>
            <label htmlFor="f_d">Invoice Date</label>
            <input
              id="f_d"
              type="date"
              value={date}
              onChange={(e) => setDate(e.target.value)}
            />
          </div>
          <div>
            <label htmlFor="f_n">Invoice No</label>
            <input
              id="f_n"
              value={invoiceNo}
              placeholder="e.g. INV-001"
              onChange={(e) => setInvoiceNo(e.target.value)}
            />
          </div>
          <div>
            <label htmlFor="f_p">Supplier / Party Name</label>
            <input
              id="f_p"
              value={partyName}
              placeholder="e.g. Acme Corp"
              onChange={(e) => setPartyName(e.target.value)}
            />
          </div>
          <div>
            <label htmlFor="f_g">GST No</label>
            <input
              id="f_g"
              maxLength={15}
              value={gstNo}
              placeholder="Blank for unregistered"
              onChange={(e) => setGstNo(e.target.value.toUpperCase())}
            />
          </div>
          <div>
            <label htmlFor="f_t">Taxable Value (₹)</label>
            <input
              id="f_t"
              inputMode="decimal"
              value={taxable}
              placeholder="0.00"
              onChange={(e) => setTaxable(e.target.value)}
            />
          </div>
          <div>
            <label htmlFor="f_i">IGST (₹)</label>
            <input
              id="f_i"
              inputMode="decimal"
              value={igst}
              placeholder="0.00"
              onChange={(e) => setIgst(e.target.value)}
            />
          </div>
          <div>
            <label htmlFor="f_c">CGST (₹)</label>
            <input
              id="f_c"
              inputMode="decimal"
              value={cgst}
              placeholder="0.00"
              onChange={(e) => setCgst(e.target.value)}
            />
          </div>
          <div>
            <label htmlFor="f_s">SGST (₹)</label>
            <input
              id="f_s"
              inputMode="decimal"
              value={sgst}
              placeholder="0.00"
              onChange={(e) => setSgst(e.target.value)}
            />
          </div>
        </div>

        <div style={{ marginTop: '14px', display: 'flex', gap: '24px', flexWrap: 'wrap' }}>
          <span>
            Total Tax: <b>{fm(totalTax)}</b>
          </span>
          <span>
            Total Invoice Value: <b>{fm(totalInvoice)}</b>
          </span>
        </div>

        {error && <div className="fe">{error}</div>}

        <div className="fr">
          <button type="button" onClick={onClose}>
            Cancel
          </button>
          <button type="button" className="pri" onClick={handleSubmit}>
            Save Invoice
          </button>
        </div>
      </div>
    </div>
  );
}
