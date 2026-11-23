import express from 'express';
import {
  getInvoices,
  createInvoice,
  updateInvoice,
  deleteInvoice,
  clearInvoices,
  bulkImport,
  getLockStatus,
  toggleLock,
  syncToGoogleSheets
} from '../controllers/invoiceController.js';

const router = express.Router();

router.get('/', getInvoices);
router.post('/', createInvoice);
router.put('/:id', updateInvoice);
router.delete('/actions/clear', clearInvoices);
router.delete('/:id', deleteInvoice);
router.post('/actions/bulk-import', bulkImport);
router.get('/actions/lock-status', getLockStatus);
router.post('/actions/toggle-lock', toggleLock);
router.post('/actions/sync-sheets', syncToGoogleSheets);

export default router;
