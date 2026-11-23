// Ensure TLS certificate verification does not fail on Windows proxy / corporate environments
if (!process.env.NODE_TLS_REJECT_UNAUTHORIZED) {
  process.env.NODE_TLS_REJECT_UNAUTHORIZED = '0';
}

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
import sheetsRoutes from './server/routes/sheetsRoutes.js';
import registerRoutes from './server/routes/registerRoutes.js';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// Load .env variables if present
if (fs.existsSync(path.join(__dirname, '.env'))) {
  try {
    if (typeof process.loadEnvFile === 'function') {
      process.loadEnvFile(path.join(__dirname, '.env'));
    } else {
      const envRaw = fs.readFileSync(path.join(__dirname, '.env'), 'utf-8');
      envRaw.split('\n').forEach((line) => {
        const trimmed = line.trim();
        if (trimmed && !trimmed.startsWith('#')) {
          const eqIdx = trimmed.indexOf('=');
          if (eqIdx !== -1) {
            const k = trimmed.slice(0, eqIdx).trim();
            const v = trimmed.slice(eqIdx + 1).trim().replace(/^["']|["']$/g, '');
            if (!process.env[k]) process.env[k] = v;
          }
        }
      });
    }
  } catch (_) {}
}

export const app = express();
const PORT = process.env.PORT || 3000;
const isProd = process.env.NODE_ENV === 'production';

// Common middleware
app.use(cors());
app.use(express.json({ limit: '10mb' }));
app.use(express.urlencoded({ extended: true, limit: '10mb' }));

app.use((req, res, next) => {
  if (req.path.startsWith('/api')) {
    console.log(`[API Request] ${req.method} ${req.originalUrl}`);
  }
  next();
});

// Safe error handler for invalid JSON bodies
app.use((err, req, res, next) => {
  if (err instanceof SyntaxError && err.status === 400 && 'body' in err) {
    return res.status(400).json({ success: false, message: 'Invalid JSON payload' });
  }
  next(err);
});

// API Routes
app.use('/api/invoices', invoiceRoutes);
app.use('/api/dashboard', dashboardRoutes);
app.use('/api/steps', stepsRoutes);
app.use('/api/settings', settingsRoutes);
app.use('/api/sheets', sheetsRoutes);
app.use('/api/registers', registerRoutes);

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

    // Explicit SPA fallback for any non-API GET in development
    app.use('*', async (req, res, next) => {
      if (req.method !== 'GET' || req.originalUrl.startsWith('/api')) {
        return next();
      }
      try {
        const indexPath = path.join(__dirname, 'index.html');
        if (fs.existsSync(indexPath)) {
          const template = fs.readFileSync(indexPath, 'utf-8');
          const html = await vite.transformIndexHtml(req.originalUrl, template);
          res.status(200).set({ 'Content-Type': 'text/html' }).end(html);
        } else {
          next();
        }
      } catch (e) {
        vite.ssrFixStacktrace(e);
        next(e);
      }
    });
  } else {
    // Production mode: Serve built static files
    const distPath = path.join(__dirname, 'dist');
    if (fs.existsSync(distPath)) {
      app.use(express.static(distPath));
    }
    app.get('*', (req, res) => {
      if (!req.path.startsWith('/api')) {
        const indexFile = path.join(distPath, 'index.html');
        if (fs.existsSync(indexFile)) {
          res.sendFile(indexFile);
        } else {
          res.status(404).send('Please run npm run build first');
        }
      } else {
        res.status(404).json({ error: 'Endpoint not found' });
      }
    });
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
