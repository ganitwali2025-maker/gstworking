import React, { useMemo, useState, useEffect } from 'react';
import { updateQueryParams } from '../utils/router';
import { normalizeFy } from '../services/registerService';
import { isMonthLocked, GST_LOCK_CHANGED_EVENT } from '../services/lockService';

export const MONTH_OPTIONS = [
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

export interface HeadingSearchBarProps {
  value: string;
  onChange: (val: string) => void;
  placeholder?: string;
  className?: string;
}

/**
 * Common Sheet Search Bar:
 * Displays a professional search icon, input field, and clear (×) button.
 */
export function HeadingSearchBar({
  value,
  onChange,
  placeholder = 'Search GSTIN, Party, Invoice No...',
  className = ''
}: HeadingSearchBarProps) {
  return (
    <div className={`gstr1-search-wrapper ${className}`} title="Search by GSTIN, Party Name, Invoice No, etc.">
      <div className="gstr1-search-icon">
        <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round">
          <circle cx="11" cy="11" r="8"></circle>
          <line x1="21" y1="21" x2="16.65" y2="16.65"></line>
        </svg>
      </div>
      <input
        type="text"
        className="gstr1-search-input"
        value={value}
        onChange={(e) => onChange(e.target.value)}
        placeholder={placeholder}
        aria-label="Search records"
      />
      {value ? (
        <button
          type="button"
          className="gstr1-search-clear"
          onClick={() => onChange('')}
          title="Clear search"
          aria-label="Clear search"
        >
          <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round">
            <line x1="18" y1="6" x2="6" y2="18"></line>
            <line x1="6" y1="6" x2="18" y2="18"></line>
          </svg>
        </button>
      ) : null}
    </div>
  );
}

interface HeadingTemplateProps {
  fy?: string;
  setFy?: (fy: string) => void;
  month?: string;
  setMonth?: (month: string) => void;
  fyOptions?: string[];
  searchQuery?: string;
  setSearchQuery?: (q: string) => void;
  searchPlaceholder?: string;
  showSearch?: boolean;
  className?: string;
}

export default function HeadingTemplate({
  fy = 'FY 2026–27',
  setFy,
  month = 'Oct',
  setMonth,
  fyOptions = DEFAULT_FY_OPTIONS,
  searchQuery,
  setSearchQuery,
  searchPlaceholder,
  showSearch = false,
  className = ''
}: HeadingTemplateProps) {
  const [lockVersion, setLockVersion] = useState(0);

  useEffect(() => {
    const onLockChange = () => setLockVersion((v) => v + 1);
    window.addEventListener(GST_LOCK_CHANGED_EVENT, onLockChange);
    return () => window.removeEventListener(GST_LOCK_CHANGED_EVENT, onLockChange);
  }, []);

  const allFyOptions = Array.from(new Set(['FY 2026–27', 'FY 2025–26', ...(fyOptions || [])]));

  const selectedFyValue = useMemo(() => {
    if (!fy || fy.toUpperCase() === 'ALL') return 'FY 2026–27';
    const clean = fy.trim();
    const matched = allFyOptions.find(
      (opt) => opt.toLowerCase() === clean.toLowerCase() || normalizeFy(opt) === normalizeFy(clean)
    );
    return matched || clean;
  }, [fy, allFyOptions]);

  const selectedMonthValue = useMemo(() => {
    if (!month || month.toUpperCase() === 'ALL') return 'Oct';
    const clean = month.trim().toLowerCase();
    const matched = MONTH_OPTIONS.find(
      (m) => m.value.toLowerCase() === clean || m.label.toLowerCase() === clean || clean.startsWith(m.value.toLowerCase())
    );
    return matched ? matched.value : clean;
  }, [month]);

  // Dynamic 12 month options with Year and Lock icon (e.g. "October 2026 🔒")
  const dynamicMonthOptions = useMemo(() => {
    const yrMatch = selectedFyValue.match(/\b(20\d\d)\b/);
    const startYr = yrMatch ? parseInt(yrMatch[1], 10) : 2026;
    const endYr = startYr + 1;

    return MONTH_OPTIONS.map((m) => {
      const isLateQuarter = ['Jan', 'Feb', 'Mar'].includes(m.value);
      const year = isLateQuarter ? endYr : startYr;
      const locked = isMonthLocked(m.value, selectedFyValue);
      return {
        value: m.value,
        label: `${m.label} ${year}${locked ? ' 🔒' : ''}`
      };
    });
  }, [selectedFyValue, lockVersion]);

  const handleFyChange = (val: string) => {
    if (setFy) setFy(val);
    const defaultMon = val.includes('2026') ? 'Oct' : 'Apr';
    if (setMonth) setMonth(defaultMon);
    updateQueryParams({ fy: val, month: defaultMon });
  };

  const handleMonthChange = (val: string) => {
    if (setMonth) setMonth(val);
    updateQueryParams({ month: val });
  };

  return (
    <div className={`heading-template-wrapper ${className}`} style={{ display: 'inline-flex', alignItems: 'center', gap: '8px' }}>
      {/* Optional Search Bar when used in sheets or header */}
      {(showSearch || searchQuery !== undefined) && setSearchQuery && (
        <HeadingSearchBar
          value={searchQuery || ''}
          onChange={setSearchQuery}
          placeholder={searchPlaceholder}
        />
      )}

      {/* 1. FY Select Dropdown Button (White pill with purple calendar icon) */}
      {setFy && (
        <label className="gstr1-select-wrapper" title="Select Financial Year">
          <div className="gstr1-select-icon">
            <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="#7c3aed" strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round">
              <rect x="3" y="4" width="18" height="18" rx="2" ry="2"></rect>
              <line x1="16" y1="2" x2="16" y2="6"></line>
              <line x1="8" y1="2" x2="8" y2="6"></line>
              <line x1="3" y1="10" x2="21" y2="10"></line>
            </svg>
          </div>
          <select
            className="gstr1-select-field"
            value={selectedFyValue}
            onChange={(e) => handleFyChange(e.target.value)}
          >
            {allFyOptions.map((y) => (
              <option key={y} value={y}>{y}</option>
            ))}
          </select>
        </label>
      )}

      {/* 2. Month Select Dropdown Button (White pill with purple calendar icon) */}
      {setMonth && (
        <label className="gstr1-select-wrapper" title="Select Month">
          <div className="gstr1-select-icon">
            <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="#7c3aed" strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round">
              <rect x="3" y="4" width="18" height="18" rx="2" ry="2"></rect>
              <line x1="16" y1="2" x2="16" y2="6"></line>
              <line x1="8" y1="2" x2="8" y2="6"></line>
              <line x1="3" y1="10" x2="21" y2="10"></line>
            </svg>
          </div>
          <select
            className="gstr1-select-field"
            value={selectedMonthValue}
            onChange={(e) => handleMonthChange(e.target.value)}
          >
            {dynamicMonthOptions.map((m) => (
              <option key={m.value} value={m.value}>{m.label}</option>
            ))}
          </select>
        </label>
      )}
    </div>
  );
}
