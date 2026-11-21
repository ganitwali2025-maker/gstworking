# Unified GST ITC Liability Dashboard & GSTR-1 Application

Ek hi single unified project jisme **React (Vite) Frontend** aur **Node.js (Express) Backend** ek saath integraded hain.

- **Sirf 1 command run karni hai:** `npm run dev`
- **Sirf 1 URL open karna hai:** `http://localhost:3000`
- **Frontend aur Backend dono ek hi port par chalte hain!**

---

## 🚀 Quick Start (Localhost)

Terminal / Command Prompt open karein aur run karein:

```bash
npm run dev
```

Browser mein open karein:
👉 **[http://localhost:3000](http://localhost:3000)**

---

## 📁 Single Unified Folder Structure

```
GST APP/
├── src/                      # React Application Components & UI
│   ├── assets/               # CSS & Styling
│   ├── components/
│   │   ├── common/           # Toast, ConfirmModal, LoginModal
│   │   ├── gstr1/            # SalesRegister, InvoiceModal
│   │   ├── itc/              # ItcDashboard, Charts, Rec Meter
│   │   ├── layout/           # Sidebar, TopHeader, Banner
│   │   ├── liability/        # LiabilityDashboard, Coverage Bar
│   │   ├── settings/         # SettingsView
│   │   └── steps/            # WorkingSteps (5 Steps & Notes)
│   ├── context/              # GstContext (Global State & Sync)
│   ├── services/             # api.js (Fetch client for /api)
│   ├── utils/                # formatters.js (INR, Dates, Math)
│   ├── App.jsx               # Main React Component
│   ├── index.css             # Main Design Stylesheet
│   └── main.jsx              # React Entry Point
│
├── server/                   # Backend API Controllers & Database
│   ├── controllers/          # Business logic (Dashboard, Invoices, Settings, Steps)
│   ├── data/
│   │   ├── db.json           # File-backed Database Store
│   │   └── store.js          # DB Read / Write Helper
│   └── routes/               # Express API Route Handlers
│
├── server.js                 # Unified Server (Serves Vite Frontend + Express /api)
├── index.html                # App HTML Entry
├── vite.config.js            # Vite Configuration
├── vercel.json               # Vercel Deployment Configuration
├── package.json              # All Dependencies & Single Dev Script
└── README.md
```

---

## 📡 Live API Endpoints (Port 3000)

- **Health Check**: `http://localhost:3000/api/health`
- **Dashboard Summary**: `http://localhost:3000/api/dashboard/summary`
- **Invoices**: `http://localhost:3000/api/invoices`
- **Steps & Notes**: `http://localhost:3000/api/steps`
- **Settings**: `http://localhost:3000/api/settings`

---

## ☁️ Vercel Deployment (1-Click)

1. Repository ko GitHub par push karein.
2. Vercel mein **"Add New Project"** karein.
3. Vercel build settings:
   - **Framework Preset**: Vite
   - **Build Command**: `npm run build`
   - **Output Directory**: `dist`
4. Deploy! `vercel.json` automatically `/api/*` requests ko backend par aur baki sabko React frontend par route karega.
