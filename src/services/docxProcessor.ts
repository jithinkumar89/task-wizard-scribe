
import * as mammoth from 'mammoth';
import { Document, Packer, Paragraph, Table, TableRow, TableCell, TextRun } from 'docx';
import JSZip from 'jszip'; // Changed import statement
import { saveAs } from 'file-saver';
import { Task } from '@/components/TaskPreview';

// Interface for extracted content
interface ExtractedContent {
  docTitle: string;
  tasks: Task[];
  images: Array<{
    taskNumber: string;
    imageData: Blob;
    contentType: string;
  }>;
}

// Regex patterns for parsing
const stepNumberRegex = /^(\d+(?:\.\d+)*)\.?\s+/;
const imageRelationshipRegex = /<img[^>]+src="([^"]+)"[^>]*>/g;

// Process the uploaded Word document
export const processDocument = async (file: File): Promise<ExtractedContent> => {
  try {
    // Extract HTML content from the docx file
    const result = await mammoth.extractRawText({ 
      arrayBuffer: await file.arrayBuffer() 
    });
    
    // Also extract images for separate processing
    const imageResult = await mammoth.convertToHtml({
      arrayBuffer: await file.arrayBuffer()
    });

    // Process raw document to extract docTitle from first line
    const lines = result.value.split('\n').filter(line => line.trim() !== '');
    const docTitle = lines[0].trim();
    
    // Extract images and their relationships
    const images = await extractImages(file, imageResult.value);
    
    // Extract tasks from the document content
    const tasks = extractTasks(lines.slice(1), docTitle, images);
    
    return {
      docTitle,
      tasks,
      images
    };
  } catch (error) {
    console.error('Error processing document:', error);
    throw new Error('Failed to process the document. Please check the file format.');
  }
};

// Extract tasks from text content
const extractTasks = (
  lines: string[], 
  docTitle: string, 
  images: Array<{ taskNumber: string; imageData: Blob; contentType: string }>
): Task[] => {
  const tasks: Task[] = [];
  let currentStepNumber = '';
  let currentTask = '';
  
  for (const line of lines) {
    const trimmedLine = line.trim();
    
    // Skip empty lines
    if (trimmedLine === '') continue;
    
    // Check if line starts with a step number pattern (e.g., "1.", "1.1.", "1.1.1.")
    const stepMatch = trimmedLine.match(stepNumberRegex);
    
    if (stepMatch) {
      // If we were processing a previous task, save it
      if (currentStepNumber && currentTask) {
        // Find if this step has an associated image
        const hasImage = images.some(img => img.taskNumber === formatTaskNumber(currentStepNumber));
        
        tasks.push({
          taskNumber: formatTaskNumber(currentStepNumber),
          type: 'Operation',
          etaSec: '',
          description: docTitle,
          activity: currentTask,
          specification: '',
          attachment: hasImage ? formatTaskNumber(currentStepNumber) : '',
          hasImage: hasImage
        });
      }
      
      // Start a new task
      currentStepNumber = stepMatch[1];
      currentTask = trimmedLine.substring(stepMatch[0].length).trim();
    } else {
      // Append to current task description
      currentTask += ' ' + trimmedLine;
    }
  }
  
  // Add the last task if there is one
  if (currentStepNumber && currentTask) {
    const hasImage = images.some(img => img.taskNumber === formatTaskNumber(currentStepNumber));
    
    tasks.push({
      taskNumber: formatTaskNumber(currentStepNumber),
      type: 'Operation',
      etaSec: '',
      description: docTitle,
      activity: currentTask,
      specification: '',
      attachment: hasImage ? formatTaskNumber(currentStepNumber) : '',
      hasImage: hasImage
    });
  }
  
  return tasks;
};

// Format the task number as required (e.g., 1.2.3 -> 1-2-03)
const formatTaskNumber = (stepNumber: string): string => {
  const parts = stepNumber.split('.');
  
  if (parts.length === 1) {
    return `${parts[0]}-0-00`;
  } else if (parts.length === 2) {
    return `${parts[0]}-${parts[1]}-00`;
  } else if (parts.length >= 3) {
    // Pad the last number with leading zeros if needed
    const lastPart = parts[2].padStart(2, '0');
    return `${parts[0]}-${parts[1]}-${lastPart}`;
  }
  
  return stepNumber;
};

// Extract images from the document
const extractImages = async (file: File, htmlContent: string): Promise<Array<{ taskNumber: string; imageData: Blob; contentType: string }>> => {
  const zip = new JSZip(); // Fixed constructor usage
  await zip.loadAsync(await file.arrayBuffer());
  
  // Load the document.xml to identify image relationships
  const documentXml = await zip.file('word/document.xml')?.async('text');
  const relationshipsXml = await zip.file('word/_rels/document.xml.rels')?.async('text');
  
  if (!documentXml || !relationshipsXml) {
    throw new Error('Invalid DOCX file structure');
  }
  
  // Extract image relationships
  const relationshipsMap = new Map<string, string>();
  const relationshipRegex = /<Relationship[^>]+Id="([^"]+)"[^>]+Target="([^"]+)"[^>]+Type="[^"]+"[^>]*>/g;
  let relationshipMatch;
  
  while ((relationshipMatch = relationshipRegex.exec(relationshipsXml)) !== null) {
    relationshipsMap.set(relationshipMatch[1], relationshipMatch[2]);
  }
  
  // Extract images and their associated step numbers from HTML content
  const images: Array<{ taskNumber: string; imageData: Blob; contentType: string }> = [];
  const imageSections = htmlContent.split('<p>');
  
  for (let i = 0; i < imageSections.length; i++) {
    const section = imageSections[i];
    
    // Look for step numbers before images
    const stepMatch = section.match(stepNumberRegex);
    if (!stepMatch) continue;
    
    const stepNumber = stepMatch[1];
    const formattedTaskNumber = formatTaskNumber(stepNumber);
    
    // Look for image tags
    let imgMatch;
    while ((imgMatch = imageRelationshipRegex.exec(section)) !== null) {
      const imgSrc = imgMatch[1];
      const imgRelId = imgSrc.split('/').pop()?.replace('rId', '');
      
      if (imgRelId) {
        // Find the image file in the zip
        for (const [filePath, fileObj] of Object.entries(zip.files)) {
          if (filePath.startsWith('word/media/') && filePath.includes(imgRelId)) {
            const imageData = await fileObj.async('blob');
            const contentType = getContentTypeFromPath(filePath);
            
            images.push({
              taskNumber: formattedTaskNumber,
              imageData,
              contentType
            });
            break;
          }
        }
      }
    }
  }
  
  return images;
};

// Get content type from file path
const getContentTypeFromPath = (path: string): string => {
  const extension = path.split('.').pop()?.toLowerCase();
  switch (extension) {
    case 'png':
      return 'image/png';
    case 'jpg':
    case 'jpeg':
      return 'image/jpeg';
    case 'gif':
      return 'image/gif';
    case 'bmp':
      return 'image/bmp';
    case 'svg':
      return 'image/svg+xml';
    default:
      return 'application/octet-stream';
  }
};

// Generate a task master document from extracted tasks
export const generateTaskMasterDocument = (docTitle: string, tasks: Task[]): Blob => {
  const doc = new Document({
    sections: [
      {
        properties: {},
        children: [
          new Paragraph({
            children: [
              new TextRun({
                text: docTitle,
                bold: true,
                size: 28
              })
            ]
          }),
          new Paragraph({
            children: [
              new TextRun({
                text: "Task Master",
                bold: true,
                size: 24
              })
            ],
            spacing: { after: 400 }
          }),
          new Table({
            rows: [
              new TableRow({
                tableHeader: true,
                children: [
                  new TableCell({ children: [new Paragraph("Task No")] }),
                  new TableCell({ children: [new Paragraph("Type")] }),
                  new TableCell({ children: [new Paragraph("ETA (sec)")] }),
                  new TableCell({ children: [new Paragraph("Description")] }),
                  new TableCell({ children: [new Paragraph("Activity")] }),
                  new TableCell({ children: [new Paragraph("Specification")] }),
                  new TableCell({ children: [new Paragraph("Attachment")] })
                ]
              }),
              ...tasks.map(task => 
                new TableRow({
                  children: [
                    new TableCell({ children: [new Paragraph(task.taskNumber)] }),
                    new TableCell({ children: [new Paragraph(task.type)] }),
                    new TableCell({ children: [new Paragraph(task.etaSec)] }),
                    new TableCell({ children: [new Paragraph(task.description)] }),
                    new TableCell({ children: [new Paragraph(task.activity)] }),
                    new TableCell({ children: [new Paragraph(task.specification)] }),
                    new TableCell({ children: [new Paragraph(task.hasImage ? task.attachment : "")] })
                  ]
                })
              )
            ]
          })
        ]
      }
    ]
  });
  
  return Packer.toBlob(doc);
};

// Create a ZIP file containing task master document and extracted images
export const createDownloadPackage = async (
  docBlob: Blob, 
  images: Array<{ taskNumber: string; imageData: Blob; contentType: string }>,
  docTitle: string
): Promise<Blob> => {
  const zip = new JSZip();
  
  // Add the generated document
  zip.file(`${docTitle} - Task Master.docx`, docBlob);
  
  // Create images folder
  const imagesFolder = zip.folder("images");
  
  // Add each image with the appropriate name
  if (imagesFolder) {
    images.forEach(image => {
      const extension = image.contentType.split('/')[1] || 'png';
      imagesFolder.file(`${image.taskNumber}.${extension}`, image.imageData);
    });
  }
  
  // Generate the zip file and return it directly as a Blob
  return await zip.generateAsync({ type: "blob" });
};
