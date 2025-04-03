import { DocumentService } from './services/documentService';
import * as fs from 'fs';
import * as path from 'path';
import { Document, Packer, Paragraph, ImageRun, AlignmentType, TextRun, UnderlineType } from 'docx';
import sharp from 'sharp';

async function createNewDocument() {
  try {
    // Create a new document service instance
    const documentService = new DocumentService();
    
    // Get image dimensions
    const imagePath = path.join(__dirname, '..', 'image.png');
    const imageBuffer = fs.readFileSync(imagePath);
    const metadata = await sharp(imageBuffer).metadata();
    
    // Calculate width proportionally for 50px height
    const targetHeight = 50;
    const aspectRatio = metadata.width! / metadata.height!;
    const targetWidth = Math.round(targetHeight * aspectRatio);
    
    // Create a new document with proper structure
    const doc = new Document({
      sections: [
        {
          children: [
            // Add centered image
            new Paragraph({
              alignment: AlignmentType.CENTER,
              children: [
                new ImageRun({
                  data: imageBuffer,
                  transformation: {
                    width: targetWidth,
                    height: targetHeight,
                  },
                  type: 'png'
                }),
              ],
            }),
            // Add some spacing
            new Paragraph({}),
            // Add title
            new Paragraph({
              alignment: AlignmentType.CENTER,
              children: [
                new TextRun({
                  text: "CERTIFICATE OF ANALYSIS",
                  bold: true,
                  size: 36,
                  underline: {
                    type: UnderlineType.SINGLE,
                    color: "000000"
                  },
                }),
              ],
            }),
            // Add some spacing
            new Paragraph({}),
            // Add all tables from document service
            documentService.createProductInfoTable(),
            new Paragraph({}),
            documentService.createTestTable(documentService.getTestTables()[0])
          ]
        }
      ]
    });

    // Generate the document
    const buffer = await Packer.toBuffer(doc);
    
    // Create output directory if it doesn't exist
    const outputDir = path.join(__dirname, '..', 'output');
    if (!fs.existsSync(outputDir)) {
      fs.mkdirSync(outputDir);
    }

    // Save the document
    const outputPath = path.join(outputDir, 'certificate.docx');
    fs.writeFileSync(outputPath, buffer);
    
    console.log(`Document saved at: ${outputPath}`);
    console.log(`Image dimensions: ${targetWidth}x${targetHeight} pixels`);
  } catch (error) {
    console.error('Error generating document:', error);
    process.exit(1);
  }
}

createNewDocument(); 