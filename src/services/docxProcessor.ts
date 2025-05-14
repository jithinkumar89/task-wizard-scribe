
import * as mammoth from 'mammoth';
import { Document, Packer, Paragraph, Table, TableRow, TableCell, TextRun } from 'docx';
import { saveAs } from 'file-saver';
import { Task } from '@/components/TaskPreview';
import * as XLSX from 'xlsx';
import JSZip from 'jszip';

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
  /^(\d+)\s*[\.\)]\s+/, // Simple number with dot or parenthesis: "1) ", "2. "
  /^[Ss][Ll]\s*\.?\s*[Nn][Oo]\s*\.?\s*(\d+)\s*[\.:]?\s*/, // SL NO 1: or SL. NO. 1.
  /^[Pp]rocedure\s+(\d+)/ // Format with "Procedure" prefix: Procedure 1, etc.
];

// Additional header terms to detect table headers more accurately
const headerTerms = [
  'sl no', 'sl.no', 'sl. no', 'serial no', 'serial number',
  'step no', 'task no', 'job details', 'description', 'activity',
  'operation', 'procedure', 'instruction', 'steps', 'sl-no', 
  'task description', 'work instruction', 'action', '#', 'no.',
  'tasks', 'item', 'procedures'
];

// Check if a line is likely a header row
const isHeaderRow = (line: string): boolean => {
  const lowerLine = line.toLowerCase();
  
  // If multiple header terms are found, it's likely a header
  let headerTermsFound = 0;
  for (const term of headerTerms) {
    if (lowerLine.includes(term)) {
      headerTermsFound++;
      if (headerTermsFound >= 1) {
        return true;
      }
    }
  }
  
  // Check for tab-separated or multiple-space-separated format that might indicate a table header
  const hasTabSeparation = /\t/.test(line);
  const hasMultipleSpaceSeparation = /\s{3,}/.test(line);
  const hasNumberLabels = /^\s*([0-9]+|[#])\s*\./.test(line) && /description|details|procedure|step/i.test(line);
  
  // If it has separation format and contains at least one header term
  return ((hasTabSeparation || hasMultipleSpaceSeparation) && headerTermsFound > 0) || hasNumberLabels;
};

// Process the uploaded Word document with improved handling for larger files
export const processDocument = async (file: File, assemblySequenceId: string = '1'): Promise<ExtractedContent> => {
  try {
    console.log("Processing document started with assembly sequence ID:", assemblySequenceId);
    
    // Extract HTML content from the docx file for text parsing
    const result = await mammoth.extractRawText({ 
      arrayBuffer: await file.arrayBuffer() 
    });
    
    // Also extract images using HTML conversion
    const imageResult = await mammoth.convertToHtml({
      arrayBuffer: await file.arrayBuffer(),
      // Set a high transformation limit to handle large documents
      transformDocument: mammoth.transforms.paragraph(paragraph => {
        return paragraph;
      })
    });

    console.log("Document text extracted successfully");
    
    // Process raw document to extract text lines
    let lines = result.value.split('\n').filter(line => line.trim() !== '');
    
    // Make sure we have some content
    if (lines.length === 0) {
      throw new Error("The document appears to be empty. Please check the file content.");
    }
    
    // Check if the first line is likely a document title or header
    let docTitle = '';
    let startLineIndex = 0;
    
    if (lines.length > 0) {
      // Try to find a good document title from the first few lines
      for (let i = 0; i < Math.min(5, lines.length); i++) {
        if (!isHeaderRow(lines[i]) && lines[i].length > 5 && lines[i].length < 100) {
          docTitle = lines[i].trim();
          startLineIndex = i + 1;
          break;
        }
      }
      
      // If we didn't find a good title, use the first line
      if (docTitle === '') {
        docTitle = lines[0].trim();
        startLineIndex = 1;
      }
    }
    
    console.log("Document title:", docTitle);
    
    // Extract images and their relationships
    const images = await extractImages(file, imageResult.value, assemblySequenceId);
    console.log(`Extracted ${images.length} images from the document`);
    
    // Try to detect if the document has a table structure
    const hasTableStructure = detectTableStructure(result.value);
    
    // Extract tasks based on document structure using multiple methods for better coverage
    let tasks: Task[] = [];
    let methodsToTry = [
      { name: "table", fn: () => extractTasksFromTable(result.value, docTitle, images, assemblySequenceId) },
      { name: "paragraph", fn: () => extractTasks(lines.slice(startLineIndex), docTitle, images, assemblySequenceId) },
      { name: "aggressive", fn: () => extractTasksAggressively(result.value, docTitle, images, assemblySequenceId) }
    ];
    
    // If we detected a table structure, prioritize table extraction
    if (hasTableStructure) {
      console.log("Detected table structure, prioritizing table-based extraction...");
      methodsToTry = [
        methodsToTry[0],  // table method
        methodsToTry[1],  // paragraph method
        methodsToTry[2]   // aggressive method
      ];
    } else {
      console.log("No table structure detected, prioritizing paragraph-based extraction...");
      methodsToTry = [
        methodsToTry[1],  // paragraph method
        methodsToTry[0],  // table method
        methodsToTry[2]   // aggressive method
      ];
    }
    
    // Try each extraction method until we get tasks
    for (const method of methodsToTry) {
      console.log(`Trying ${method.name}-based task extraction...`);
      tasks = method.fn();
      
      // If we found a good number of tasks, stop trying other methods
      if (tasks.length > 0) {
        console.log(`${method.name} extraction successful, found ${tasks.length} tasks`);
        break;
      }
    }
    
    console.log(`Extracted ${tasks.length} tasks from document`);
    
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

// Map images to tasks based on figure references in the text with improved pattern matching
const mapImagesToTasks = (
  tasks: Task[], 
  images: Array<{ task_no: string; imageData: Blob; contentType: string }>,
  documentText: string
): Task[] => {
  // Create a mapping of tasks to image references
  const taskImageMapping: Record<string, string[]> = {};
  
  // Extract figure references from document text with multiple patterns
  const figurePatterns = [
    /figure\s+(\d+)/gi,
    /fig\.?\s+(\d+)/gi,
    /photo\s+(\d+)/gi,
    /picture\s+(\d+)/gi,
    /image\s+(\d+)/gi,
    /illustration\s+(\d+)/gi
  ];
  
  const figureReferences: number[] = [];
  
  // Find all figure references in the document using multiple patterns
  figurePatterns.forEach(pattern => {
    let match;
    while ((match = pattern.exec(documentText)) !== null) {
      figureReferences.push(parseInt(match[1], 10));
    }
  });
  
  // Map each task to the images it references
  tasks.forEach(task => {
    // Extract the task content
    const taskContent = task.activity;
    const taskFigures: string[] = [];
    
    // Look for figure references in this task using multiple patterns
    figurePatterns.forEach(pattern => {
      // Reset the lastIndex property to start searching from the beginning
      pattern.lastIndex = 0;
      let match;
      while ((match = pattern.exec(taskContent)) !== null) {
        const figureNum = parseInt(match[1], 10);
        const imageId = formatImageId(figureNum, task.task_no?.split('.')[0] || '1');
        if (!taskFigures.includes(imageId)) {
          taskFigures.push(imageId);
        }
      }
    });
    
    // Store the mapping
    if (taskFigures.length > 0) {
      taskImageMapping[task.task_no || ''] = taskFigures;
    }
  });
  
  // If no specific figure references were found in tasks, distribute images evenly
  if (Object.keys(taskImageMapping).length === 0 && tasks.length > 0 && images.length > 0) {
    console.log("No specific figure references found, distributing images evenly among tasks");
    
    // Calculate roughly how many images per task
    const imagesPerTask = Math.ceil(images.length / tasks.length);
    
    tasks.forEach((task, index) => {
      const start = index * imagesPerTask;
      const end = Math.min(start + imagesPerTask, images.length);
      
      if (start < images.length) {
        const taskImages: string[] = [];
        for (let i = start; i < end; i++) {
          taskImages.push(images[i].task_no);
        }
        
        taskImageMapping[task.task_no || ''] = taskImages;
      }
    });
  }
  
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
  // More sophisticated heuristic to detect tables
  const tableIndicators = [
    /\t[^\t]+\t[^\t]+/g,  // Tab separated content
    /\s{2,}[^\s]+\s{2,}[^\s]+/g,  // Space separated (2+ spaces)
    /^\d+\.\s+[^\n]+\n\d+\.\s+/m  // Numbered list format (1. xxx\n2. xxx)
  ];
  
  // Count the number of lines that match table patterns
  let tableLineCount = 0;
  const lines = content.split('\n');
  
  for (const line of lines) {
    if (tableIndicators.some(pattern => pattern.test(line))) {
      tableLineCount++;
    }
  }
  
  // If more than 20% of lines look like table rows, consider it a table structure
  return tableLineCount > lines.length * 0.2;
};

// Extract tasks from text content with improved patterns
const extractTasks = (
  lines: string[], 
  docTitle: string, 
  images: Array<{ task_no: string; imageData: Blob; contentType: string }>,
  assemblySequenceId: string = '1'
): Task[] => {
  const tasks: Task[] = [];
  let currentTaskNum: number | null = null;
  let currentTask = '';
  let lastFoundTaskNum = 0;
  
  for (const line of lines) {
    const trimmedLine = line.trim();
    
    // Skip empty lines or header-like lines
    if (trimmedLine === '' || isHeaderRow(trimmedLine)) continue;
    
    // Check if line starts with any of the step number patterns
    let stepMatch = null;
    let matchedPattern = null;
    
    for (const pattern of stepNumberPatterns) {
      stepMatch = trimmedLine.match(pattern);
      if (stepMatch) {
        matchedPattern = pattern;
        break;
      }
    }
    
    if (stepMatch) {
      // If we were processing a previous task, save it
      if (currentTask && currentTaskNum !== null) {
        const formatted = formatTaskNumber(currentTaskNum.toString(), assemblySequenceId);
        
        tasks.push({
          task_no: formatted, 
          type: 'Operation',
          eta_sec: '',
          description: trimmedLine.substring(stepMatch[0].length).trim(),
          activity: currentTask.trim(),
          specification: '',
          attachment: '',
          hasImage: false
        });
        lastFoundTaskNum = currentTaskNum;
      }
      
      // Start a new task with the extracted task number
      currentTaskNum = parseInt(stepMatch[1], 10);
      
      // Extract content after the task number
      const taskContent = trimmedLine.substring(stepMatch[0].length).trim();
      currentTask = taskContent;
      
      // If there was a gap in task numbers, fill it with empty tasks to maintain sequence
      if (lastFoundTaskNum > 0 && currentTaskNum > lastFoundTaskNum + 1) {
        // Fill the gap with placeholder tasks
        for (let i = lastFoundTaskNum + 1; i < currentTaskNum; i++) {
          const formatted = formatTaskNumber(i.toString(), assemblySequenceId);
          tasks.push({
            task_no: formatted,
            type: 'Operation',
            eta_sec: '',
            description: '[Missing Task]',
            activity: '[This task was not found in the document]',
            specification: '',
            attachment: '',
            hasImage: false
          });
        }
      }
    } else {
      // Append to current task description
      if (currentTask) {
        currentTask += ' ' + trimmedLine;
      }
    }
  }
  
  // Add the last task if there is one
  if (currentTask && currentTaskNum !== null) {
    const formatted = formatTaskNumber(currentTaskNum.toString(), assemblySequenceId);
    
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

// More aggressive task extraction method as a fallback for complex document formats
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
  
  // First pass: look for paragraphs that appear to be tasks
  for (const paragraph of paragraphs) {
    // Skip very short paragraphs, likely headers, or header-like content
    if (paragraph.trim().length < 10 || 
        paragraph.trim() === docTitle || 
        isHeaderRow(paragraph.trim())) continue;
    
    // Check for number patterns at the beginning of paragraphs
    let taskNum: number | null = null;
    let restOfText = paragraph.trim();
    
    for (const pattern of stepNumberPatterns) {
      const match = paragraph.match(pattern);
      if (match) {
        taskNum = parseInt(match[1], 10);
        restOfText = paragraph.substring(match[0].length).trim();
        break;
      }
    }
    
    // If no explicit task number found, try to detect if it could be a task
    if (taskNum === null) {
      // Check for indicators that this might be a task instruction
      const mightBeTask = /^[A-Z][^\.]+\./.test(paragraph) || // Starts with capital letter and has a period
                           /^(Check|Ensure|Remove|Install|Attach|Connect|Verify|Place|Position)/.test(paragraph); // Starts with action verb
      
      if (mightBeTask && paragraph.length > 20) {
        taskNum = taskIndex++;
        restOfText = paragraph.trim();
      }
    } else {
      // If we found a task number, update our taskIndex to be at least this number
      taskIndex = Math.max(taskIndex, taskNum + 1);
    }
    
    // Add the task if we determined this paragraph is a task
    if (taskNum !== null) {
      const formatted = formatTaskNumber(taskNum.toString(), assemblySequenceId);
      
      tasks.push({
        task_no: formatted,
        type: 'Operation',
        eta_sec: '',
        description: restOfText,
        activity: restOfText,
        specification: '',
        attachment: '',
        hasImage: false
      });
    }
  }
  
  // If we still don't have tasks, make a more desperate attempt by treating each paragraph as a task
  if (tasks.length === 0) {
    console.log("No tasks found in first aggressive pass, treating paragraphs as sequential tasks");
    let index = 1;
    
    for (const paragraph of paragraphs) {
      // Skip very short paragraphs, duplicates of document title or obvious headers
      if (paragraph.trim().length < 15 || 
          paragraph.trim() === docTitle || 
          isHeaderRow(paragraph.trim())) continue;
      
      const formatted = formatTaskNumber(index.toString(), assemblySequenceId);
      tasks.push({
        task_no: formatted,
        type: 'Operation',
        eta_sec: '',
        description: paragraph.trim().split('\n')[0], // Use first line as description
        activity: paragraph.trim(),
        specification: '',
        attachment: '',
        hasImage: false
      });
      
      index++;
    }
  }
  
  return tasks;
};

// Extract tasks from table-structured content with improved parsing
const extractTasksFromTable = (
  content: string,
  docTitle: string,
  images: Array<{ task_no: string; imageData: Blob; contentType: string }>,
  assemblySequenceId: string = '1'
): Task[] => {
  const tasks: Task[] = [];
  const lines = content.split('\n').filter(line => line.trim().length > 0);
  let taskNum = 1;
  let skipLines = 0;
  
  // First pass to identify header row(s)
  const headerLines: number[] = [];
  for (let i = 0; i < Math.min(10, lines.length); i++) {
    if (isHeaderRow(lines[i])) {
      headerLines.push(i);
      console.log(`Identified header at line ${i}: ${lines[i]}`);
    }
  }
  
  // Skip header lines if any found
  skipLines = headerLines.length > 0 ? Math.max(...headerLines) + 1 : 0;
  
  // Process content lines (non-header)
  for (let i = skipLines; i < lines.length; i++) {
    const line = lines[i].trim();
    
    // Skip if this line also looks like a header
    if (isHeaderRow(line)) continue;
    
    // Try to find task number using various patterns
    let taskNumber: number | null = null;
    let activityContent = line;
    
    for (const pattern of stepNumberPatterns) {
      const match = line.match(pattern);
      if (match) {
        taskNumber = parseInt(match[1], 10);
        activityContent = line.substring(match[0].length).trim();
        break;
      }
    }
    
    // If no task number found, look for tab-separated or space-separated content
    if (taskNumber === null) {
      // Try to extract task number from tab or space separated content
      if (/\t/.test(line)) {
        // Tab-separated
        const parts = line.split('\t').filter(part => part.trim() !== '');
        if (parts.length >= 2) {
          const numMatch = parts[0].match(/\d+/);
          if (numMatch) {
            taskNumber = parseInt(numMatch[0], 10);
            activityContent = parts.slice(1).join(' ').trim();
          }
        }
      } else if (/\s{3,}/.test(line)) {
        // Space-separated (3+ spaces likely indicates columns)
        const parts = line.split(/\s{3,}/);
        if (parts.length >= 2) {
          const numMatch = parts[0].match(/\d+/);
          if (numMatch) {
            taskNumber = parseInt(numMatch[0], 10);
            activityContent = parts.slice(1).join(' ').trim();
          }
        }
      }
    }
    
    // If we found a task number, add the task
    if (taskNumber !== null) {
      const formatted = formatTaskNumber(taskNumber.toString(), assemblySequenceId);
      
      tasks.push({
        task_no: formatted,
        type: 'Operation',
        eta_sec: '',
        description: activityContent,
        activity: activityContent,
        specification: '',
        attachment: '',
        hasImage: false
      });
      
      // Update task number for next task if needed
      taskNum = Math.max(taskNum, taskNumber + 1);
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

// Extract images from the document with improved handling for large documents
const extractImages = async (
  file: File, 
  htmlContent: string, 
  assemblySequenceId: string = '1'
): Promise<Array<{ task_no: string; imageData: Blob; contentType: string }>> => {
  try {
    // Create a new JSZip instance
    const zip = new JSZip();
    const zipData = await file.arrayBuffer();
    const loadedZip = await zip.loadAsync(zipData);
    
    console.log("ZIP file loaded successfully");
    
    // Load the document.xml to identify image relationships
    const documentXml = await loadedZip.file('word/document.xml')?.async('text');
    const relationshipsXml = await loadedZip.file('word/_rels/document.xml.rels')?.async('text');
    
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
    const imageFiles: { [key: string]: { data: Blob, contentType: string } } = {};
    
    // First collect all images from word/media
    const zipFiles = loadedZip.files;
    for (const filePath in zipFiles) {
      const fileObj = zipFiles[filePath];
      
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
    
    // Extract figure references from HTML content with multiple patterns
    const figurePatterns = [
      /[Ff]igure\s+(\d+)/g,
      /[Ff]ig\.?\s+(\d+)/g,
      /[Ii]llustration\s+(\d+)/g,
      /[Pp]hoto\s+(\d+)/g,
      /[Pp]icture\s+(\d+)/g,
      /[Ii]mage\s+(\d+)/g
    ];
    
    const figureNumbers: Set<number> = new Set();
    
    // Find all figure references in the document
    figurePatterns.forEach(pattern => {
      let match;
      while ((match = pattern.exec(htmlContent)) !== null) {
        const figNum = parseInt(match[1], 10);
        figureNumbers.add(figNum);
      }
    });
    
    // Sort figure numbers
    const sortedFigureNumbers = Array.from(figureNumbers).sort((a, b) => a - b);
    
    // Map image files to images array with proper IDs
    const images: Array<{ task_no: string; imageData: Blob; contentType: string }> = [];
    const imageEntries = Object.entries(imageFiles);
    
    // If we have figure references, try to map them to images
    if (sortedFigureNumbers.length > 0 && imageEntries.length > 0) {
      // First handle explicit figure references
      sortedFigureNumbers.forEach((figNum, index) => {
        if (index < imageEntries.length) {
          const [, imageInfo] = imageEntries[index];
          const imageId = `${assemblySequenceId}-0-${figNum.toString().padStart(3, '0')}`;
          images.push({
            task_no: imageId,
            imageData: imageInfo.data,
            contentType: imageInfo.contentType
          });
        }
      });
      
      // Then handle any remaining images
      if (imageEntries.length > sortedFigureNumbers.length) {
        let nextFigNum = sortedFigureNumbers.length > 0 ? Math.max(...sortedFigureNumbers) + 1 : 1;
        
        for (let i = sortedFigureNumbers.length; i < imageEntries.length; i++) {
          const [, imageInfo] = imageEntries[i];
          const imageId = `${assemblySequenceId}-0-${nextFigNum.toString().padStart(3, '0')}`;
          images.push({
            task_no: imageId,
            imageData: imageInfo.data,
            contentType: imageInfo.contentType
          });
          nextFigNum++;
        }
      }
    } else {
      // If no figure references, assign sequential numbers to all images
      imageEntries.forEach(([, imageInfo], index) => {
        const imageId = `${assemblySequenceId}-0-${(index + 1).toString().padStart(3, '0')}`;
        images.push({
          task_no: imageId,
          imageData: imageInfo.data,
          contentType: imageInfo.contentType
        });
      });
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
  zip.file(`${docTitle}.xlsx`, excelBlob);
  
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
