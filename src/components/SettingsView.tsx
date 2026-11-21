import React, { useState, useEffect } from 'react';
import { BusinessSettings } from '../types/gst';

interface SettingsViewProps {
  settings: BusinessSettings;
  onSave: (newSettings: BusinessSettings) => void;
}

export default function SettingsView({ settings, onSave }: SettingsViewProps) {
  const [formData, setFormData] = useState<BusinessSettings>(settings);
  const [saving, setSaving] = useState(false);
  const [toastText, setToastText] = useState('');

  useEffect(() => {
    setFormData(settings);
  }, [settings]);

  const handleChange = (field: keyof BusinessSettings, value: any) => {
    setFormData((prev) => ({ ...prev, [field]: value }));
  };

  const handleSave = () => {
    setSaving(true);
    onSave(formData);
    setToastText('Saved');
    setTimeout(() => {
      setToastText('');
      setSaving(false);
    }, 2000);
  };

  return (
    <section id="v-set">
      <div className="card" style={{ marginBottom: '14px' }}>
        <h3 style={{ marginBottom: '14px' }}>Business profile</h3>
        <div className="form">
          <div className="f">
            <label htmlFor="s1">Business name</label>
            <input
              id="s1"
              value={formData.businessName}
              onChange={(e) => handleChange('businessName', e.target.value)}
            />
          </div>
          <div className="f">
            <label htmlFor="s2">GSTIN</label>
            <input
              id="s2"
              value={formData.gstin}
              onChange={(e) => handleChange('gstin', e.target.value.toUpperCase())}
            />
          </div>
          <div className="f">
            <label htmlFor="s3">Filing frequency</label>
            <select
              id="s3"
              value={formData.filingFrequency}
              onChange={(e) => handleChange('filingFrequency', e.target.value)}
            >
              <option value="Monthly">Monthly</option>
              <option value="Quarterly (QRMP)">Quarterly (QRMP)</option>
            </select>
          </div>
          <div className="f">
            <label htmlFor="s4">State</label>
            <input
              id="s4"
              value={formData.state}
              onChange={(e) => handleChange('state', e.target.value)}
            />
          </div>
        </div>
      </div>

      <div className="card">
        <h3>Preferences</h3>
        <div className="tog">
          <span>Due date reminder (3 din pehle)</span>
          <input
            type="checkbox"
            checked={formData.dueDateReminder}
            onChange={(e) => handleChange('dueDateReminder', e.target.checked)}
            aria-label="Due date reminder"
          />
        </div>
        <div className="tog">
          <span>2B mismatch alert</span>
          <input
            type="checkbox"
            checked={formData.mismatchAlert}
            onChange={(e) => handleChange('mismatchAlert', e.target.checked)}
            aria-label="Mismatch alert"
          />
        </div>
        <div className="tog">
          <span>Monthly summary email</span>
          <input
            type="checkbox"
            checked={formData.monthlySummaryEmail}
            onChange={(e) => handleChange('monthlySummaryEmail', e.target.checked)}
            aria-label="Summary email"
          />
        </div>

        <div style={{ marginTop: '16px', display: 'flex', alignItems: 'center' }}>
          <button type="button" className="btn" onClick={handleSave} disabled={saving}>
            {saving ? 'Saving...' : 'Save changes'}
          </button>
          {toastText && <span className="toast">{toastText}</span>}
        </div>
      </div>
    </section>
  );
}
