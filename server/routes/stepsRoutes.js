import express from 'express';
import {
  getSteps,
  updateStepStatus,
  addStepNote,
  deleteStepNote
} from '../controllers/stepsController.js';

const router = express.Router();

router.get('/', getSteps);
router.put('/:id/status', updateStepStatus);
router.post('/:id/notes', addStepNote);
router.delete('/:id/notes/:noteIndex', deleteStepNote);

export default router;
