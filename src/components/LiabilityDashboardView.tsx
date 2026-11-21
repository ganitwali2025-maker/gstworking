import React from 'react';
import { inr } from '../utils/gstFormat';
import { MonthlyStats } from '../types/gst';

interface LiabilityDashboardViewProps {
  month: string;
  monthlyData: Record<string, MonthlyStats>;
}

export default function LiabilityDashboardView({ month, monthlyData }: LiabilityDashboardViewProps) {
  const months = Object.keys(monthlyData);
  const cur = monthlyData[month] || monthlyData[months[months.length - 1]];

  const sum = (a: number[]) => a.reduce((x, y) => x + y, 0);

  const l = cur.liability;
  const i = cur.itc;
  const totalLiab = sum(l);
  const totalItc = sum(i);
  const cashPayable = sum(l.map((v, k) => Math.max(0, v - i[k])));
  const itcAdjusted = totalLiab - cashPayable;
  const itcCarryForward = sum(i.map((v, k) => Math.max(0, v - l[k])));

  const pct = totalLiab > 0 ? Math.round((itcAdjusted / totalLiab) * 100) : 0;
  const cashPct = 100 - pct;

  const heads = ['IGST', 'CGST', 'SGST'];

  const dueDates = [
    { day: '11', month: 'Nov', form: 'GSTR-1', description: 'Outward supplies file karein' },
    { day: '13', month: 'Nov', form: 'IFF / GSTR-1A', description: 'Invoice furnishing (QRMP)' },
    { day: '20', month: 'Nov', form: 'GSTR-3B', description: 'Tax payment aur return' },
    { day: '25', month: 'Nov', form: 'PMT-06', description: 'QRMP challan payment' }
  ];

  return (
    <section id="v-liab">
      {/* 4 Clean KPI Cards without collar */}
      <div className="grid kpis">
        <div className="card kpi">
          <small>Output GST liability</small>
          <b>{inr(totalLiab)}</b>
          <span>Is month ki total tax</span>
        </div>
        <div className="card kpi">
          <small>ITC adjusted</small>
          <b>{inr(itcAdjusted)}</b>
          <span>ITC se cover hua</span>
        </div>
        <div className="card kpi">
          <small>Cash payable</small>
          <b>{inr(cashPayable)}</b>
          <span>Challan se dena hai</span>
        </div>
        <div className="card kpi">
          <small>Effective cash rate</small>
          <b>{pct < 100 ? `${cashPct}%` : '0%'}</b>
          <span>{pct < 100 ? 'Liability ka cash hissa' : 'Pura ITC se covered'}</span>
        </div>
      </div>

      {/* Coverage Bar */}
      <div className="card" style={{ marginBottom: '14px' }}>
        <h3>Liability kaise cover hui</h3>
        <div className="bar">
          <div style={{ width: `${pct}%`, background: 'var(--itc-soft)', color: 'var(--itc)' }}>
            ITC {inr(itcAdjusted)} ({pct}%)
          </div>
          <div style={{ width: `${cashPct}%`, background: 'var(--cash-soft)', color: 'var(--cash)' }}>
            Cash {inr(cashPayable)} ({cashPct}%)
          </div>
        </div>
        <div className="legend">
          <span>
            <i style={{ background: 'var(--itc)' }} />
            ITC se adjust
          </span>
          <span>
            <i style={{ background: 'var(--cash)' }} />
            Cash se payable
          </span>
        </div>
      </div>

      {/* Breakup & Dues */}
      <div className="grid two">
        <div className="card">
          <h3>Tax head ke hisaab se breakup</h3>
          <div className="tw" style={{ marginTop: '10px' }}>
            <table>
              <thead>
                <tr>
                  <th>Head</th>
                  <th>Liability</th>
                  <th>ITC</th>
                  <th>Cash payable</th>
                  <th>Carry forward</th>
                </tr>
              </thead>
              <tbody>
                {heads.map((h, k) => {
                  const c = Math.max(0, l[k] - i[k]);
                  const f = Math.max(0, i[k] - l[k]);
                  return (
                    <tr key={h}>
                      <td>{h}</td>
                      <td>{inr(l[k])}</td>
                      <td>{inr(i[k])}</td>
                      <td className="pos">{inr(c)}</td>
                      <td>{inr(f)}</td>
                    </tr>
                  );
                })}
              </tbody>
              <tfoot>
                <tr>
                  <td>Total</td>
                  <td>{inr(totalLiab)}</td>
                  <td>{inr(totalItc)}</td>
                  <td className="pos">{inr(cashPayable)}</td>
                  <td>{inr(itcCarryForward)}</td>
                </tr>
              </tfoot>
            </table>
          </div>
        </div>

        <div className="card">
          <h3>Aane wali due dates</h3>
          <div style={{ marginTop: '10px' }}>
            {dueDates.map((due, idx) => (
              <div key={idx} className="due">
                <div className="date">
                  {due.day}
                  <small>{due.month}</small>
                </div>
                <div>
                  <b>{due.form}</b>
                  <br />
                  <small style={{ color: 'var(--mute)' }}>{due.description}</small>
                </div>
              </div>
            ))}
          </div>
        </div>
      </div>
    </section>
  );
}
