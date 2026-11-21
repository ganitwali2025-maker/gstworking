import React, { useState } from 'react';
import { InvoiceRecord } from '../types/gst';
import { pp, fm } from '../utils/gstFormat';

interface ModalProps {
  isOpen: boolean;
  invoice: InvoiceRecord | null;
  isLocked: boolean;
  fy: string;
  onClose: () => void;
  onSave: (record: InvoiceRecord) => void;
}

export default function Modal({
  isOpen,
  invoice,
  isLocked,
  fy,
  onClose,
  onSave
}: ModalProps) {
  if (!isOpen) return null;

  const [date, setDate] = useState(invoice ? invoice.d : new Date().toISOString().slice(0, 10));
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

    const newRecord: InvoiceRecord = {
      id: invoice ? invoice.id : 'inv-' + Date.now().toString(36) + Math.random().toString(36).slice(2, 6),
      fy,
      q: 'Q3',
      m: 'Oct 2026',
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
          Period: Financial Year {fy}
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
