import { getDb } from '../data/store.js';

const sum = (arr) => arr.reduce((a, b) => a + b, 0);

export const getDashboardSummary = (req, res) => {
  try {
    const db = getDb();
    const { month } = req.query;

    const monthlyData = db.monthlyData || {};
    const months = Object.keys(monthlyData);
    const selectedMonth = month && monthlyData[month] ? month : months[months.length - 1];

    const currentStats = monthlyData[selectedMonth] || {
      liability: [0, 0, 0],
      itc: [0, 0, 0],
      rec: { matchRate: 0, booksItc: 0, gstr2bItc: 0, ineligible: 0 }
    };

    const l = currentStats.liability;
    const i = currentStats.itc;

    const totalLiab = sum(l);
    const totalItc = sum(i);

    const cashPayable = sum(l.map((val, idx) => Math.max(0, val - i[idx])));
    const itcCarryForward = sum(i.map((val, idx) => Math.max(0, val - l[idx])));
    const itcAdjusted = totalLiab - cashPayable;
    const coverPercent = totalLiab > 0 ? Math.round((itcAdjusted / totalLiab) * 100) : 0;

    // Month-wise trend for chart
    const trend = months.map(m => {
      const data = monthlyData[m];
      const liabSum = sum(data.liability);
      const itcSum = sum(data.itc);
      return {
        month: m,
        liability: liabSum,
        itc: itcSum
      };
    });

    // Tax heads breakup
    const heads = ['IGST', 'CGST', 'SGST'];
    const taxHeadsBreakup = heads.map((head, idx) => ({
      head,
      liability: l[idx],
      itc: i[idx],
      cashPayable: Math.max(0, l[idx] - i[idx]),
      carryForward: Math.max(0, i[idx] - l[idx])
    }));

    // Upcoming dues
    const dueDates = [
      { day: '11', month: 'Nov', form: 'GSTR-1', description: 'Outward supplies file karein' },
      { day: '13', month: 'Nov', form: 'IFF / GSTR-1A', description: 'Invoice furnishing (QRMP)' },
      { day: '20', month: 'Nov', form: 'GSTR-3B', description: 'Tax payment aur return' },
      { day: '25', month: 'Nov', form: 'PMT-06', description: 'QRMP challan payment' }
    ];

    res.json({
      success: true,
      selectedMonth,
      months,
      kpis: {
        totalLiability: totalLiab,
        totalItc: totalItc,
        itcAdjusted,
        cashPayable,
        itcCarryForward,
        ineligibleItc: currentStats.rec.ineligible,
        coverPercent
      },
      rec: currentStats.rec,
      trend,
      taxHeadsBreakup,
      dueDates,
      rcmDetails: db.rcmDetails || []
    });
  } catch (error) {
    res.status(500).json({ success: false, message: error.message });
  }
};
