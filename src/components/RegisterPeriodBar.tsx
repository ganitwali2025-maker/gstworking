import React from 'react';
import { ViewType } from '../types/gst';

interface RegisterPeriodBarProps {
  view: ViewType;
}

const TITLES: Record<string, [string, string]> = {
  itc: ['GST ITC', 'Input tax credit available, reconcile aur use kaise hua'],
  liab: ['Liability', 'Output tax, ITC adjustment aur cash payable'],
  steps: ['GST Working Steps', 'FY 2026–27 ke GST working steps, ek-ek karke tick karein'],
  set: ['Settings', 'Business profile, filing aur alerts ki settings'],
  rcm: ['Reverse Charge Mechanism', 'RCM liability and eligible ITC tracking'],
  rec: ['ITC Reconciliation', 'GSTR-2B vs Books reconciliation statement']
};

export default function RegisterPeriodBar({
  view
}: RegisterPeriodBarProps) {
  if (['g1', 'itc-reg', 'rcm', 'g2b', 'rec'].includes(view)) return null;

  const titleInfo = TITLES[view] || TITLES.itc;

  return (
    <div className="banner">
      <div>
        <h1>{titleInfo[0]}</h1>
        <p>{titleInfo[1]}</p>
      </div>
    </div>
  );
}
