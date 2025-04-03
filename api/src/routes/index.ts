import { Router } from 'express';
import documentRoutes from './documentRoutes';

const router = Router();

// Health check endpoint
router.get('/health', (req, res) => {
  res.json({ status: 'ok' });
});

// Document routes
router.use('/documents', documentRoutes);

// Example route
router.get('/hello', (req, res) => {
  res.json({ message: 'Hello from the API!' });
});

export default router; 