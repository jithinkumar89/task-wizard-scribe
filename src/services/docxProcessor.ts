
import * as mammoth from 'mammoth';
import { Document, Packer, Paragraph, Table, TableRow, TableCell, TextRun } from 'docx';
import JSZip from 'jszip';
import { saveAs } from 'file-saver';
import { Task } from '@/components/TaskPreview';
import * as XLSX from 'xlsx';

// Interface for extracted content
interface ExtractedContent {
  docTitle: string;
  tasks: Task[];
  images: Array<{
    task_no: string;
    imageData: Blob;
    contentType: string;
  }>;
}

// Multiple regex patterns for parsing different formats of step numbers
const stepNumberPatterns = [
  /^(\d+)\.?\s+/, // Standard format: 1., 2., etc.
  /^Step\s+(\d+)\.?\s+/i, // Format with "Step" prefix: Step 1., Step 2., etc.
  /^(\d+)\)\s+/, // Format with parenthesis: 1), 2), etc.
  /^[a-zA-Z]?\s*(\d+)[\.\)]\s+/, // Format with optional letter prefix: a 1), A. 2, etc.
  /^Task\s+(\d+)\.?\s+/i, // Format with "Task" prefix: Task 1., Task 2., etc.
  /^Sl\.\s*No\.\s*(\d+)/i, // Format with "Sl. No." prefix: Sl. No. 1, etc.
  /^(\d+)\s*[\.\)]\s+/ // Simple number with dot or parenthesis: "1) ", "2. "
];

// Process the uploaded Word document
export const processDocument = async (file: File, assemblySequenceId: string = '1'): Promise<ExtractedContent> => {
  try {
    console.log("Processing document started with assembly sequence ID:", assemblySequenceId);
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
    const images = await extractImages(file, imageResult.value, assemblySequenceId);
    console.log(`Extracted ${images.length} images from the document`);
    
    // Try to detect if the document has a table structure
    const hasTableStructure = detectTableStructure(result.value);
    
    // Extract tasks based on document structure
    let tasks: Task[] = [];
    if (hasTableStructure) {
      console.log("Detected table structure, extracting tasks from table...");
      tasks = extractTasksFromTable(result.value, docTitle, images, assemblySequenceId);
    } else {
      console.log("Using paragraph-based task extraction...");
      tasks = extractTasks(lines.slice(1), docTitle, images, assemblySequenceId);
    }
    
    console.log(`Extracted ${tasks.length} tasks from document`);
    
    // If no tasks were found, try a more aggressive approach
    if (tasks.length === 0) {
      console.log("No tasks found with primary method, trying alternative extraction...");
      tasks = extractTasksAggressively(result.value, docTitle, images, assemblySequenceId);
      console.log(`Alternative extraction found ${tasks.length} tasks`);
    }
    
    // Map images to tasks based on figure references
    tasks = mapImagesToTasks(tasks, images, result.value);
    
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

// Map images to tasks based on figure references in the text
const mapImagesToTasks = (
  tasks: Task[], 
  images: Array<{ task_no: string; imageData: Blob; contentType: string }>,
  documentText: string
): Task[] => {
  // Create a mapping of tasks to image references
  const taskImageMapping: Record<string, string[]> = {};
  
  // Extract figure references from document text
  const figurePattern = /figure\s+(\d+)/gi;
  let figureMatch;
  const figureReferences: number[] = [];
  
  while ((figureMatch = figurePattern.exec(documentText)) !== null) {
    figureReferences.push(parseInt(figureMatch[1], 10));
  }
  
  // Map each task to the images it references
  tasks.forEach(task => {
    // Extract the task content
    const taskContent = task.activity;
    const taskFigures: string[] = [];
    
    // Look for figure references in this task
    const taskFigurePattern = /figure\s+(\d+)/gi;
    let taskFigureMatch;
    
    while ((taskFigureMatch = taskFigurePattern.exec(taskContent)) !== null) {
      const figureNum = parseInt(taskFigureMatch[1], 10);
      const imageId = formatImageId(figureNum, task.task_no?.split('.')[0] || '1');
      taskFigures.push(imageId);
    }
    
    // Store the mapping
    if (taskFigures.length > 0) {
      taskImageMapping[task.task_no || ''] = taskFigures;
    }
  });
  
  // Update tasks with their image references
  return tasks.map(task => {
    const taskImages = taskImageMapping[task.task_no || ''] || [];
    return {
      ...task,
      attachment: taskImages.join(', '),
      hasImage: taskImages.length > 0
    };
  });
};

// Format image ID from figure number and assembly ID
const formatImageId = (figureNumber: number, assemblyId: string): string => {
  return `${assemblyId}-0-${figureNumber.toString().padStart(3, '0')}`;
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
  images: Array<{ task_no: string; imageData: Blob; contentType: string }>,
  assemblySequenceId: string = '1'
): Task[] => {
  const tasks: Task[] = [];
  let currentTaskIndex = 0;
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
      if (currentTask) {
        currentTaskIndex++;
        const formatted = formatTaskNumber(currentTaskIndex.toString(), assemblySequenceId);
        
        tasks.push({
          task_no: formatted, // Updated to use task_no instead of taskNumber
          type: 'Operation',
          eta_sec: '',
          description: trimmedLine.substring(stepMatch[0].length).trim(),
          activity: currentTask.trim(),
          specification: '',
          attachment: '',
          hasImage: false
        });
      }
      
      // Start a new task
      currentTask = trimmedLine.substring(stepMatch[0].length).trim();
    } else {
      // Append to current task description
      if (currentTask) {
        currentTask += ' ' + trimmedLine;
      }
    }
  }
  
  // Add the last task if there is one
  if (currentTask) {
    currentTaskIndex++;
    const formatted = formatTaskNumber(currentTaskIndex.toString(), assemblySequenceId);
    
    tasks.push({
      task_no: formatted,
      type: 'Operation',
      eta_sec: '',
      description: currentTask.trim(),
      activity: currentTask.trim(),
      specification: '',
      attachment: '',
      hasImage: false
    });
  }
  
  return tasks;
};

// More aggressive task extraction method as a fallback
const extractTasksAggressively = (
  content: string,
  docTitle: string,
  images: Array<{ task_no: string; imageData: Blob; contentType: string }>,
  assemblySequenceId: string = '1'
): Task[] => {
  const tasks: Task[] = [];
  
  // Split by potential paragraph markers
  const paragraphs = content.split(/\n\n|\r\n\r\n/).filter(p => p.trim().length > 0);
  let taskIndex = 1;
  
  for (const paragraph of paragraphs) {
    // Skip very short paragraphs and likely headers
    if (paragraph.trim().length < 10 || paragraph.trim() === docTitle) continue;
    
    // Attempt to find a number at the start of the paragraph
    const numberMatch = paragraph.match(/^\s*(\d+)/);
    
    // If we found a number, use it as the task index, otherwise use incremental index
    if (numberMatch) {
      taskIndex = parseInt(numberMatch[1], 10);
    }
    
    const formatted = formatTaskNumber(taskIndex.toString(), assemblySequenceId);
    
    tasks.push({
      task_no: formatted,
      type: 'Operation',
      eta_sec: '',
      description: paragraph.trim(),
      activity: paragraph.trim(),
      specification: '',
      attachment: '',
      hasImage: false
    });
    
    taskIndex++;
  }
  
  return tasks;
};

// Extract tasks from table-structured content
const extractTasksFromTable = (
  content: string,
  docTitle: string,
  images: Array<{ task_no: string; imageData: Blob; contentType: string }>,
  assemblySequenceId: string = '1'
): Task[] => {
  const tasks: Task[] = [];
  const lines = content.split('\n').filter(line => line.trim().length > 0);
  let taskIndex = 1;
  
  for (let i = 1; i < lines.length; i++) {  // Skip the first line as it's likely a header
    const line = lines[i].trim();
    
    // Look for a number at the beginning of the line
    const match = line.match(/^\s*(\d+)/);
    if (match) {
      const stepNumber = parseInt(match[1], 10);
      taskIndex = stepNumber; // Use the found step number as the task index
      
      const restOfLine = line.substring(match[0].length).trim();
      
      const formatted = formatTaskNumber(taskIndex.toString(), assemblySequenceId);
      
      tasks.push({
        task_no: formatted,
        type: 'Operation',
        eta_sec: '',
        description: restOfLine,
        activity: restOfLine,
        specification: '',
        attachment: '',
        hasImage: false
      });
    }
  }
  
  return tasks;
};

// Format the task number as required (e.g., for assembly ID 1, task 1 becomes 1.0.001)
const formatTaskNumber = (stepNumber: string, assemblySequenceId: string = '1'): string => {
  // Convert the step number to a three-digit format with leading zeros
  const formattedStepNumber = stepNumber.padStart(3, '0');
  
  // Return in the format: assemblySequenceId.0.formattedStepNumber
  return `${assemblySequenceId}.0.${formattedStepNumber}`;
};

// Extract images from the document
const extractImages = async (
  file: File, 
  htmlContent: string, 
  assemblySequenceId: string = '1'
): Promise<Array<{ task_no: string; imageData: Blob; contentType: string }>> => {
  try {
    // Create a new JSZip instance - fixed constructor issue
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
    const images: Array<{ task_no: string; imageData: Blob; contentType: string }> = [];
    const imageFiles: { [key: string]: { data: Blob, contentType: string } } = {};
    
    // First collect all images from word/media
    for (const filePath in zip.files) {
      const fileObj = zip.files[filePath];
      
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
    
    // Extract figure references from HTML content
    const figurePattern = /Figure\s+(\d+)/gi;
    let figureMatch;
    const figureNumbers: number[] = [];
    
    // Find all figure references in the document
    while ((figureMatch = figurePattern.exec(htmlContent)) !== null) {
      const figNum = parseInt(figureMatch[1], 10);
      if (!figureNumbers.includes(figNum)) {
        figureNumbers.push(figNum);
      }
    }
    
    // Sort figure numbers
    figureNumbers.sort((a, b) => a - b);
    
    // Map figure numbers to images
    if (figureNumbers.length > 0 && Object.keys(imageFiles).length > 0) {
      // For each figure reference, assign an image
      const imageEntries = Object.entries(imageFiles);
      figureNumbers.forEach((figNum, index) => {
        if (index < imageEntries.length) {
          const [imageName, imageInfo] = imageEntries[index];
          // Use the figure number in the image ID
          const imageId = `${assemblySequenceId}-0-${figNum.toString().padStart(3, '0')}`;
          images.push({
            task_no: imageId,
            imageData: imageInfo.data,
            contentType: imageInfo.contentType
          });
        }
      });
      
      // Handle any remaining images
      if (imageEntries.length > figureNumbers.length) {
        for (let i = figureNumbers.length; i < imageEntries.length; i++) {
          const [imageName, imageInfo] = imageEntries[i];
          const nextFigNum = (figureNumbers.length > 0 ? Math.max(...figureNumbers) : 0) + (i - figureNumbers.length + 1);
          const imageId = `${assemblySequenceId}-0-${nextFigNum.toString().padStart(3, '0')}`;
          images.push({
            task_no: imageId,
            imageData: imageInfo.data,
            contentType: imageInfo.contentType
          });
        }
      }
    } else {
      // If no figure references, just assign sequential IDs
      let imgIndex = 1;
      for (const [imageName, imageInfo] of Object.entries(imageFiles)) {
        const imageId = `${assemblySequenceId}-0-${imgIndex.toString().padStart(3, '0')}`;
        images.push({
          task_no: imageId,
          imageData: imageInfo.data,
          contentType: imageInfo.contentType
        });
        imgIndex++;
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

// Generate an Excel file from extracted tasks
export const generateExcelFile = async (tasks: Task[], docTitle: string): Promise<Blob> => {
  try {
    // Create workbook
    const wb = XLSX.utils.book_new();
    
    // Convert tasks to format for Excel
    const excelData = tasks.map(task => ({
      task_no: task.task_no,
      type: task.type,
      eta_sec: task.eta_sec,
      description: task.description,
      activity: task.activity,
      specification: task.specification,
      attachment: task.attachment
    }));
    
    // Create worksheet
    const ws = XLSX.utils.json_to_sheet(excelData);
    
    // Add worksheet to workbook
    XLSX.utils.book_append_sheet(wb, ws, "Tasks");
    
    // Generate Excel file
    const excelBlob = new Blob(
      [XLSX.write(wb, { bookType: 'xlsx', type: 'array' })], 
      { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }
    );
    
    return excelBlob;
  } catch (error) {
    console.error("Error generating Excel file:", error);
    throw new Error("Failed to generate Excel file");
  }
};

// Create a ZIP file containing task master document and extracted images
export const createDownloadPackage = async (
  excelBlob: Blob, 
  images: Array<{ task_no: string; imageData: Blob; contentType: string }>,
  docTitle: string
): Promise<Blob> => {
  // Create a new JSZip instance
  const zip = new JSZip();
  
  // Add the generated Excel file
  zip.file(`${docTitle} - Task Master.xlsx`, excelBlob);
  
  // Create images folder
  const imagesFolder = zip.folder("images");
  
  // Add each image with the appropriate name
  if (imagesFolder) {
    images.forEach(image => {
      const extension = image.contentType.split('/')[1] || 'png';
      imagesFolder.file(`${image.task_no}.${extension}`, image.imageData);
    });
  }
  
  // Generate the zip file and return it as a Blob
  return await zip.generateAsync({ type: "blob" });
};
