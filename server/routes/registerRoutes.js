import express from 'express';
import {
  getRegisterRows,
  createRegisterRow,
  updateRegisterRow,
  deleteRegisterRow,
  bulkImportRegister,
  clearRegister,
  syncRegister
} from '../controllers/registerController.js';

const router = express.Router();

// Register-specific CRUD & actions
router.get('/:tabName', getRegisterRows);
router.post('/:tabName', createRegisterRow);
router.put('/:tabName/:id', updateRegisterRow);
router.delete('/:tabName/:id', deleteRegisterRow);

router.post('/:tabName/actions/bulk-import', bulkImportRegister);
router.delete('/:tabName/actions/clear', clearRegister);
router.post('/:tabName/actions/sync', syncRegister);

export default router;
