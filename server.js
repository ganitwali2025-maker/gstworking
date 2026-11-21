import express from 'express';
import cors from 'cors';
import path from 'path';
import fs from 'fs';
import { fileURLToPath } from 'url';
import { createServer as createViteServer } from 'vite';

import invoiceRoutes from './server/routes/invoiceRoutes.js';
import dashboardRoutes from './server/routes/dashboardRoutes.js';
import stepsRoutes from './server/routes/stepsRoutes.js';
import settingsRoutes from './server/routes/settingsRoutes.js';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

export const app = express();
const PORT = process.env.PORT || 3000;
const isProd = process.env.NODE_ENV === 'production';

// Common middleware
app.use(cors());
app.use(express.json({ limit: '10mb' }));
app.use(express.urlencoded({ extended: true, limit: '10mb' }));

// API Routes
app.use('/api/invoices', invoiceRoutes);
app.use('/api/dashboard', dashboardRoutes);
app.use('/api/steps', stepsRoutes);
app.use('/api/settings', settingsRoutes);

app.get('/api/health', (req, res) => {
  res.json({
    status: 'ok',
    service: 'Unified GST ITC Liability Dashboard',
    mode: isProd ? 'production' : 'development',
    timestamp: new Date().toISOString()
  });
});

async function setupApp() {
  if (!isProd) {
    // Development mode: Vite dev server middleware
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: 'spa'
    });
    app.use(vite.middlewares);
  } else {
    // Production mode: Serve built static files
    const distPath = path.join(__dirname, 'dist');
    if (fs.existsSync(distPath)) {
      app.use(express.static(distPath));
      app.get('*', (req, res) => {
        if (!req.path.startsWith('/api')) {
          res.sendFile(path.join(distPath, 'index.html'));
        } else {
          res.status(404).json({ error: 'Endpoint not found' });
        }
      });
    }
  }

  // Only start listening if run directly (not as imported module on Vercel)
  if (process.env.VERCEL !== '1') {
    app.listen(PORT, () => {
      console.log(`\n==================================================`);
      console.log(`🚀 GST App live on: http://localhost:${PORT}`);
      console.log(`📡 Unified Backend & React Frontend ek saath chal rahe hain!`);
      console.log(`==================================================\n`);
    });
  }
}

setupApp();

export default app;
