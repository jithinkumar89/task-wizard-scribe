
import * as mammoth from 'mammoth';
import { Document, Packer, Paragraph, Table, TableRow, TableCell, TextRun } from 'docx';
import JSZip from 'jszip';
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

// Multiple regex patterns for parsing different formats of step numbers
const stepNumberPatterns = [
  /^(\d+(?:\.\d+)*)\.?\s+/, // Standard format: 1., 1.1., etc.
  /^Step\s+(\d+(?:\.\d+)*)\.?\s+/i, // Format with "Step" prefix: Step 1., Step 1.1., etc.
  /^(\d+(?:\.\d+)*)\)\s+/, // Format with parenthesis: 1), 1.1), etc.
  /^[a-zA-Z]?\s*(\d+(?:\.\d+)*)[\.\)]\s+/, // Format with optional letter prefix: a 1), A. 1.1, etc.
  /^Task\s+(\d+(?:\.\d+)*)\.?\s+/i // Format with "Task" prefix: Task 1., Task 1.1., etc.
];

// Process the uploaded Word document
export const processDocument = async (file: File): Promise<ExtractedContent> => {
  try {
    console.log("Processing document started...");
    // Extract HTML content from the docx file
    const result = await mammoth.extractRawText({ 
      arrayBuffer: await file.arrayBuffer() 
    });
    
    // Also extract images for separate processing
    const imageResult = await mammoth.convertToHtml({
      arrayBuffer: await file.arrayBuffer()
    });

    console.log("Document text extracted successfully");
    
    // Process raw document to extract docTitle from first line
    let lines = result.value.split('\n').filter(line => line.trim() !== '');
    
    // Make sure we have some content
    if (lines.length === 0) {
      throw new Error("The document appears to be empty. Please check the file content.");
    }
    
    const docTitle = lines[0].trim();
    console.log("Document title:", docTitle);
    
    // Extract images and their relationships
    const images = await extractImages(file, imageResult.value);
    console.log(`Extracted ${images.length} images from the document`);
    
    // Try to detect if the document has a table structure
    const hasTableStructure = detectTableStructure(result.value);
    
    // Extract tasks based on document structure
    let tasks: Task[] = [];
    if (hasTableStructure) {
      console.log("Detected table structure, extracting tasks from table...");
      tasks = extractTasksFromTable(result.value, docTitle, images);
    } else {
      console.log("Using paragraph-based task extraction...");
      tasks = extractTasks(lines.slice(1), docTitle, images);
    }
    
    console.log(`Extracted ${tasks.length} tasks from document`);
    
    // If no tasks were found, try a more aggressive approach
    if (tasks.length === 0) {
      console.log("No tasks found with primary method, trying alternative extraction...");
      tasks = extractTasksAggressively(result.value, docTitle, images);
      console.log(`Alternative extraction found ${tasks.length} tasks`);
    }
    
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

// Detect if the document likely has a table structure
const detectTableStructure = (content: string): boolean => {
  // Simple heuristic: check for repeated tab or multiple space patterns
  const tableIndicators = [
    /\t[^\t]+\t[^\t]+/g,  // Tab separated content
    /\s{2,}[^\s]+\s{2,}[^\s]+/g  // Space separated (2+ spaces)
  ];
  
  return tableIndicators.some(pattern => pattern.test(content));
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
    
    // Check if line starts with any of the step number patterns
    let stepMatch = null;
    for (const pattern of stepNumberPatterns) {
      stepMatch = trimmedLine.match(pattern);
      if (stepMatch) break;
    }
    
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
          activity: currentTask.trim(),
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
      if (currentTask) {
        currentTask += ' ' + trimmedLine;
      }
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
      activity: currentTask.trim(),
      specification: '',
      attachment: hasImage ? formatTaskNumber(currentStepNumber) : '',
      hasImage: hasImage
    });
  }
  
  return tasks;
};

// More aggressive task extraction method as a fallback
const extractTasksAggressively = (
  content: string,
  docTitle: string,
  images: Array<{ taskNumber: string; imageData: Blob; contentType: string }>
): Task[] => {
  const tasks: Task[] = [];
  
  // Split by potential paragraph markers
  const paragraphs = content.split(/\n\n|\r\n\r\n/).filter(p => p.trim().length > 0);
  let taskIndex = 1;
  
  for (const paragraph of paragraphs) {
    // Skip very short paragraphs and likely headers
    if (paragraph.trim().length < 10 || paragraph.trim() === docTitle) continue;
    
    // Attempt to find a number at the start of the paragraph
    const numberMatch = paragraph.match(/^\s*(\d+(?:\.\d+)*)/);
    const stepNumber = numberMatch ? numberMatch[1] : `${taskIndex}`;
    
    const formattedNumber = formatTaskNumber(stepNumber);
    const hasImage = images.some(img => img.taskNumber === formattedNumber);
    
    tasks.push({
      taskNumber: formattedNumber,
      type: 'Operation',
      etaSec: '',
      description: docTitle,
      activity: paragraph.trim(),
      specification: '',
      attachment: hasImage ? formattedNumber : '',
      hasImage: hasImage
    });
    
    taskIndex++;
  }
  
  return tasks;
};

// Extract tasks from table-structured content
const extractTasksFromTable = (
  content: string,
  docTitle: string,
  images: Array<{ taskNumber: string; imageData: Blob; contentType: string }>
): Task[] => {
  const tasks: Task[] = [];
  const lines = content.split('\n').filter(line => line.trim().length > 0);
  
  for (let i = 1; i < lines.length; i++) {  // Skip the first line as it's likely a header
    const line = lines[i].trim();
    
    // Look for a number at the beginning of the line
    const match = line.match(/^\s*(\d+(?:\.\d+)*)/);
    if (match) {
      const stepNumber = match[1];
      const restOfLine = line.substring(match[0].length).trim();
      
      const formattedNumber = formatTaskNumber(stepNumber);
      const hasImage = images.some(img => img.taskNumber === formattedNumber);
      
      tasks.push({
        taskNumber: formattedNumber,
        type: 'Operation',
        etaSec: '',
        description: docTitle,
        activity: restOfLine,
        specification: '',
        attachment: hasImage ? formattedNumber : '',
        hasImage: hasImage
      });
    }
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
  try {
    const zip = new JSZip();
    await zip.loadAsync(await file.arrayBuffer());
    
    console.log("ZIP file loaded successfully");
    
    // Load the document.xml to identify image relationships
    const documentXml = await zip.file('word/document.xml')?.async('text');
    const relationshipsXml = await zip.file('word/_rels/document.xml.rels')?.async('text');
    
    if (!documentXml || !relationshipsXml) {
      console.warn("Could not find document.xml or relationships file");
      return [];
    }
    
    // Extract image relationships
    const relationshipsMap = new Map<string, string>();
    const relationshipRegex = /<Relationship[^>]+Id="([^"]+)"[^>]+Target="([^"]+)"[^>]+Type="[^"]+"[^>]*>/g;
    let relationshipMatch;
    
    while ((relationshipMatch = relationshipRegex.exec(relationshipsXml)) !== null) {
      relationshipsMap.set(relationshipMatch[1], relationshipMatch[2]);
    }
    
    // Look for images in the ZIP structure directly
    const images: Array<{ taskNumber: string; imageData: Blob; contentType: string }> = [];
    const imageFiles: { [key: string]: { data: Blob, contentType: string } } = {};
    
    // First collect all images from word/media
    for (const [filePath, fileObj] of Object.entries(zip.files)) {
      if (filePath.startsWith('word/media/') && !fileObj.dir) {
        try {
          const imageData = await fileObj.async('blob');
          const contentType = getContentTypeFromPath(filePath);
          const imageName = filePath.split('/').pop() || '';
          imageFiles[imageName] = { data: imageData, contentType };
        } catch (err) {
          console.warn(`Failed to extract image from ${filePath}:`, err);
        }
      }
    }
    
    console.log(`Found ${Object.keys(imageFiles).length} image files in the document`);
    
    // Extract image references and their associated step numbers from HTML content
    const imageRelationshipRegex = /<img[^>]+src="([^"]+)"[^>]*>/g;
    const imageSections = htmlContent.split('<p>');
    
    // Map for storing the last encountered step number
    let lastStepNumber = "1";
    
    for (let i = 0; i < imageSections.length; i++) {
      const section = imageSections[i];
      
      // Look for step numbers before images
      for (const pattern of stepNumberPatterns) {
        const stepMatch = section.match(pattern);
        if (stepMatch) {
          lastStepNumber = stepMatch[1];
          break;
        }
      }
      
      const formattedTaskNumber = formatTaskNumber(lastStepNumber);
      
      // Look for image tags
      let imgMatch;
      while ((imgMatch = imageRelationshipRegex.exec(section)) !== null) {
        const imgSrc = imgMatch[1];
        // Extract the file name or relationship ID
        const imgRelId = imgSrc.split('/').pop()?.replace('rId', '') || '';
        
        // Try to find the actual image file
        for (const [imageName, imageInfo] of Object.entries(imageFiles)) {
          if (imageName.includes(imgRelId)) {
            images.push({
              taskNumber: formattedTaskNumber,
              imageData: imageInfo.data,
              contentType: imageInfo.contentType
            });
            break;
          }
        }
      }
    }
    
    // If we haven't found any images using the relationship method, 
    // just assign images sequentially to tasks
    if (images.length === 0 && Object.keys(imageFiles).length > 0) {
      console.log("Using sequential image assignment as fallback");
      let taskIndex = 1;
      for (const [imageName, imageInfo] of Object.entries(imageFiles)) {
        images.push({
          taskNumber: formatTaskNumber(String(taskIndex)),
          imageData: imageInfo.data,
          contentType: imageInfo.contentType
        });
        taskIndex++;
      }
    }
    
    return images;
  } catch (error) {
    console.error("Error extracting images:", error);
    return [];
  }
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
  
  // Generate the zip file and return it as a Blob
  return await zip.generateAsync({ type: "blob" }) as Blob;
};
