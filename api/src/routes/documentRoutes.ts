import { Router } from 'express';
import { DocumentService } from '../services/documentService';

const router = Router();
const documentService = new DocumentService();

router.get('/generate-certificate', async (req, res) => {
  try {
    const buffer = await documentService.generateCertificateOfAnalysis();
    
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', 'attachment; filename=certificate.docx');
    res.send(buffer);
  } catch (error) {
    console.error('Error generating certificate:', error);
    res.status(500).json({ error: 'Failed to generate certificate' });
  }
});

export default router; 