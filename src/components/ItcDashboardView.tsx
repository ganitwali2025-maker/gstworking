import React from 'react';
import { inr } from '../utils/gstFormat';
import { MonthlyStats } from '../types/gst';

interface ItcDashboardViewProps {
  month: string;
  setMonth: (m: string) => void;
  monthlyData: Record<string, MonthlyStats>;
}

export default function ItcDashboardView({ month, setMonth, monthlyData }: ItcDashboardViewProps) {
  const months = Object.keys(monthlyData);
  const cur = monthlyData[month] || monthlyData[months[months.length - 1]];

  const sum = (a: number[]) => a.reduce((x, y) => x + y, 0);

  const l = cur.liability;
  const i = cur.itc;
  const totalLiab = sum(l);
  const totalItc = sum(i);
  const cashPayable = sum(l.map((v, k) => Math.max(0, v - i[k])));
  const itcCarryForward = sum(i.map((v, k) => Math.max(0, v - l[k])));
  const itcAdjusted = totalLiab - cashPayable;

  const trend = months.map((m) => {
    const data = monthlyData[m];
    return {
      month: m,
      liability: sum(data.liability),
      itc: sum(data.itc)
    };
  });

  const maxVal = Math.max(...trend.map((t) => Math.max(t.liability, t.itc)), 1);

  return (
    <section id="v-itc">
      {/* 4 Clean KPI Cards without collar */}
      <div className="grid kpis">
        <div className="card kpi">
          <small>ITC available</small>
          <b>{inr(totalItc)}</b>
          <span>2B ke hisaab se</span>
        </div>
        <div className="card kpi">
          <small>ITC used</small>
          <b>{inr(itcAdjusted)}</b>
          <span>Liability se adjust</span>
        </div>
        <div className="card kpi">
          <small>ITC carry forward</small>
          <b>{inr(itcCarryForward)}</b>
          <span>Agle month ke liye</span>
        </div>
        <div className="card kpi">
          <small>Ineligible ITC</small>
          <b>{inr(cur.rec.ineligible)}</b>
          <span>Reverse karna hai</span>
        </div>
      </div>

      {/* Chart & Reconciliation */}
      <div className="grid two">
        <div className="card">
          <h3>Month-wise ITC vs liability</h3>
          <div className="legend" style={{ marginTop: '6px' }}>
            <span>
              <i style={{ background: 'var(--liab)' }} />
              Liability
            </span>
            <span>
              <i style={{ background: 'var(--itc)' }} />
              ITC
            </span>
          </div>

          <div className="chart">
            {trend.map((t) => {
              const isSelected = t.month === month;
              const liabHeight = (t.liability / maxVal) * 100;
              const itcHeight = (t.itc / maxVal) * 100;

              return (
                <button
                  key={t.month}
                  type="button"
                  className={`col ${isSelected ? 'on' : ''}`}
                  onClick={() => setMonth(t.month)}
                  aria-label={t.month}
                >
                  <div className="pair">
                    <span style={{ height: `${liabHeight}%`, background: 'var(--liab)' }} />
                    <span style={{ height: `${itcHeight}%`, background: 'var(--itc)' }} />
                  </div>
                  {t.month.slice(0, 3)}
                </button>
              );
            })}
          </div>
        </div>

        <div className="card">
          <h3>ITC reconciliation (2B vs Books)</h3>
          <div style={{ marginTop: '10px' }}>
            <div className="rec">
              <span>Books mein ITC</span>
              <b>{inr(cur.rec.booksItc)}</b>
            </div>
            <div className="rec">
              <span>GSTR-2B mein ITC</span>
              <b>{inr(cur.rec.gstr2bItc)}</b>
            </div>
            <div className="rec">
              <span>Difference</span>
              <b style={{ color: 'var(--warn)' }}>{inr(cur.rec.booksItc - cur.rec.gstr2bItc)}</b>
            </div>
            <div className="rec">
              <span>Ineligible (Sec 17(5))</span>
              <b>{inr(cur.rec.ineligible)}</b>
            </div>
          </div>

          <div className="meter">
            <div style={{ width: `${cur.rec.matchRate}%` }} />
          </div>
          <small style={{ color: 'var(--mute)', display: 'block', marginTop: '4px' }}>
            {cur.rec.matchRate}% invoices match hue
          </small>
        </div>
      </div>
    </section>
  );
}
