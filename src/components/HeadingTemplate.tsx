import React from 'react';
import { updateQueryParams } from '../utils/router';

export const MONTH_OPTIONS = [
  { value: 'ALL', label: 'All Months' },
  { value: 'Apr', label: 'April' },
  { value: 'May', label: 'May' },
  { value: 'Jun', label: 'June' },
  { value: 'Jul', label: 'July' },
  { value: 'Aug', label: 'August' },
  { value: 'Sep', label: 'September' },
  { value: 'Oct', label: 'October' },
  { value: 'Nov', label: 'November' },
  { value: 'Dec', label: 'December' },
  { value: 'Jan', label: 'January' },
  { value: 'Feb', label: 'February' },
  { value: 'Mar', label: 'March' }
];

export const DEFAULT_FY_OPTIONS = ['FY 2026–27', 'FY 2025–26'];

interface HeadingTemplateProps {
  fy?: string;
  setFy?: (fy: string) => void;
  month?: string;
  setMonth?: (month: string) => void;
  fyOptions?: string[];
  className?: string;
}

export default function HeadingTemplate({
  fy = 'ALL',
  setFy,
  month = 'ALL',
  setMonth,
  fyOptions = DEFAULT_FY_OPTIONS,
  className = ''
}: HeadingTemplateProps) {
  const allFyOptions = Array.from(new Set(['FY 2026–27', 'FY 2025–26', ...(fyOptions || [])]));

  return (
    <div className={`heading-template-wrapper ${className}`} style={{ display: 'inline-flex', alignItems: 'center', gap: '8px' }}>
      {/* 1. FY Select Dropdown Button */}
      <label className="gstr1-select-wrapper" title="Select Financial Year">
        <div className="gstr1-select-icon">
          <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round">
            <rect x="3" y="4" width="18" height="18" rx="2" ry="2"></rect>
            <line x1="16" y1="2" x2="16" y2="6"></line>
            <line x1="8" y1="2" x2="8" y2="6"></line>
            <line x1="3" y1="10" x2="21" y2="10"></line>
          </svg>
        </div>
        <span className="gstr1-select-label">FY:</span>
        <select
          className="gstr1-select-field"
          value={fy || 'ALL'}
          onChange={(e) => {
            const val = e.target.value;
            if (setFy) setFy(val);
            updateQueryParams({ fy: val });
          }}
        >
          <option value="ALL">All FY</option>
          {allFyOptions.map((y) => (
            <option key={y} value={y}>{y}</option>
          ))}
        </select>
      </label>

      {/* 2. Month Select Dropdown Button */}
      <label className="gstr1-select-wrapper" title="Select Month">
        <div className="gstr1-select-icon">
          <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round">
            <circle cx="12" cy="12" r="10"></circle>
            <polyline points="12 6 12 12 16 14"></polyline>
          </svg>
        </div>
        <span className="gstr1-select-label">Month:</span>
        <select
          className="gstr1-select-field"
          value={month || 'ALL'}
          onChange={(e) => {
            const val = e.target.value;
            if (setMonth) setMonth(val);
            updateQueryParams({ month: val });
          }}
        >
          {MONTH_OPTIONS.map((m) => (
            <option key={m.value} value={m.value}>{m.label}</option>
          ))}
        </select>
      </label>
    </div>
  );
}
