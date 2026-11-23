import { inr } from '../utils/gstFormat';
import { useMemo } from 'react';
import { MonthlyStats } from '../types/gst';

interface LiabilityDashboardViewProps {
  month: string;
  monthlyData: Record<string, MonthlyStats>;
}

export default function LiabilityDashboardView({ month, monthlyData }: LiabilityDashboardViewProps) {
  const months = Object.keys(monthlyData || {});
  const fallbackStats: MonthlyStats = {
    liability: [0, 0, 0],
    itc: [0, 0, 0],
    rec: { matchRate: 0, booksItc: 0, gstr2bItc: 0, ineligible: 0 }
  };

  const allMonthsSum: MonthlyStats = useMemo(() => {
    const list = Object.values(monthlyData || {});
    if (!list.length) return fallbackStats;
    const l0 = list.reduce((acc, x) => acc + (x.liability?.[0] || 0), 0);
    const l1 = list.reduce((acc, x) => acc + (x.liability?.[1] || 0), 0);
    const l2 = list.reduce((acc, x) => acc + (x.liability?.[2] || 0), 0);
    const i0 = list.reduce((acc, x) => acc + (x.itc?.[0] || 0), 0);
    const i1 = list.reduce((acc, x) => acc + (x.itc?.[1] || 0), 0);
    const i2 = list.reduce((acc, x) => acc + (x.itc?.[2] || 0), 0);
    const bItc = list.reduce((acc, x) => acc + (x.rec?.booksItc || 0), 0);
    const gItc = list.reduce((acc, x) => acc + (x.rec?.gstr2bItc || 0), 0);
    const inel = list.reduce((acc, x) => acc + (x.rec?.ineligible || 0), 0);
    const avgMatch = Math.round(list.reduce((acc, x) => acc + (x.rec?.matchRate || 0), 0) / list.length);
    return {
      liability: [l0, l1, l2],
      itc: [i0, i1, i2],
      rec: { matchRate: avgMatch, booksItc: bItc, gstr2bItc: gItc, ineligible: inel }
    };
  }, [monthlyData]);

  const cur = useMemo(() => {
    if (!monthlyData) return fallbackStats;
    if (!month || month.toUpperCase() === 'ALL') {
      return allMonthsSum;
    }
    const cleanMonth = month.trim().toLowerCase();
    const exact = monthlyData[month];
    if (exact) return exact;
    const matched = Object.entries(monthlyData).find(([k]) =>
      k.toLowerCase().startsWith(cleanMonth) || cleanMonth.startsWith(k.toLowerCase().slice(0, 3))
    );
    return matched ? matched[1] : fallbackStats;
  }, [month, monthlyData, allMonthsSum]);

  const sum = (a: number[]) => (a || []).reduce((x, y) => x + y, 0);

  const l = cur.liability || [0, 0, 0];
  const i = cur.itc || [0, 0, 0];
  const totalLiab = sum(l);
  const totalItc = sum(i);
  const cashPayable = sum(l.map((v, k) => Math.max(0, v - (i[k] || 0))));
  const itcAdjusted = totalLiab - cashPayable;
  const itcCarryForward = sum(i.map((v, k) => Math.max(0, v - (l[k] || 0))));

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
