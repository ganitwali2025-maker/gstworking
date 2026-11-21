import express from 'express';
import {
  getInvoices,
  createInvoice,
  updateInvoice,
  deleteInvoice,
  clearInvoices,
  bulkImport,
  getLockStatus,
  toggleLock
} from '../controllers/invoiceController.js';

const router = express.Router();

router.get('/', getInvoices);
router.post('/', createInvoice);
router.put('/:id', updateInvoice);
router.delete('/:id', deleteInvoice);
router.delete('/actions/clear', clearInvoices);
router.post('/actions/bulk-import', bulkImport);
router.get('/actions/lock-status', getLockStatus);
router.post('/actions/toggle-lock', toggleLock);

export default router;
