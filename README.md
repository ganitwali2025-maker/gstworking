# Unified GST ITC Liability Dashboard & GSTR-1 Application

Unified full-stack GST compliance web application combining a **React + TypeScript (Vite)** frontend with an **Express (Node.js)** backend in a single clean architecture.

---

## 🚀 Quick Start (Local Setup)

### 1. Install Dependencies
```bash
npm install
```

### 2. Start Development Server
```bash
npm run dev
```
Open your browser at: **[http://localhost:3000](http://localhost:3000)**

### 3. Production Build
```bash
npm run build
```
Generates production-optimized static assets in `dist/`.

### 4. Production Start
```bash
npm start
```

---

## ⚙️ Environment Variables

Create a `.env` file from `.env.example`:

| Variable | Default Value | Description |
| :--- | :--- | :--- |
| `PORT` | `3000` | Port on which the Express server listens |
| `NODE_ENV` | `development` | Environment mode (`development` or `production`) |
| `DATABASE_URL` | *(optional)* | PostgreSQL connection string for Prisma ORM |

---

## 📁 Project Architecture

```text
GST APP/
├── src/                          # React + TypeScript Frontend
│   ├── components/               # UI Views and Modals
│   │   ├── Gstr2bRegisterView.tsx
│   │   ├── ItcDashboardView.tsx
│   │   ├── ItcRegisterView.tsx
│   │   ├── LiabilityDashboardView.tsx
│   │   ├── Modal.tsx
│   │   ├── RcmRegisterView.tsx
│   │   ├── ReconciliationView.tsx
│   │   ├── RegisterPeriodBar.tsx
│   │   ├── SettingsView.tsx
│   │   ├── Sidebar.tsx
│   │   ├── TaxRegisterView.tsx
│   │   ├── TopBar.tsx
│   │   └── WorkingStepsView.tsx
│   ├── data/
│   │   └── seedData.ts           # Frontend fallback seed data
│   ├── hooks/
│   │   └── useGstStore.ts        # Global state management & local storage sync
│   ├── types/
│   │   └── gst.ts                # TypeScript interfaces and types
│   ├── utils/
│   │   └── gstFormat.ts          # INR currency formatters, date parsing & helpers
│   ├── App.tsx                   # Main React Application
│   ├── main.tsx                  # React DOM Entry
│   ├── index.css                 # Application styles
│   └── vite-env.d.ts             # Vite client type definitions
│
├── server/                       # Express Backend
│   ├── controllers/              # API Controllers
│   │   ├── dashboardController.js
│   │   ├── invoiceController.js
│   │   ├── settingsController.js
│   │   └── stepsController.js
│   ├── data/
│   │   ├── db.json               # File-based JSON database
│   │   └── store.js              # Safe Database read/write helpers
│   └── routes/                   # API Route definitions
│
├── api/
│   └── index.js                  # Vercel serverless function entry
├── prisma/
│   └── schema.prisma             # Prisma ORM schema
├── public/                       # Static assets
├── _archive/                     # Archived legacy HTML and duplicate files
├── server.js                     # Unified server (Express API + Vite middleware)
├── tsconfig.json                 # TypeScript strict configuration
├── vite.config.ts                # Vite React configuration
├── vercel.json                   # Vercel deployment configuration
├── package.json                  # Dependencies & scripts
└── .gitignore                    # Git ignore rules
```

---

## 📡 Backend API Endpoints (Port 3000)

- `GET /api/health` - Server health status and environment mode
- `GET /api/dashboard/summary` - ITC & Liability summary stats, monthly trend, and due dates
- `GET /api/invoices` - Fetch GSTR-1 invoices (supports `?fy=`, `?q=`, `?m=`)
- `POST /api/invoices` - Create a new sales invoice
- `PUT /api/invoices/:id` - Update existing invoice
- `DELETE /api/invoices/:id` - Delete an invoice
- `POST /api/invoices/actions/toggle-lock` - Lock/Unlock sales register
- `DELETE /api/invoices/actions/clear` - Clear all invoices for period
- `GET /api/steps` - Working steps checklist and notes
- `PUT /api/steps/:id/status` - Mark step complete/incomplete
- `POST /api/steps/:id/notes` - Add note to a step
- `GET /api/settings` - Retrieve business profile settings
- `PUT /api/settings` - Update business profile settings
