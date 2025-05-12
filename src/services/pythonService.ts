
import { PythonShell } from 'python-shell';
import { Task } from '@/components/TaskPreview';
import * as fs from 'fs';
import * as path from 'path';
import * as os from 'os';

// This service is designed to be used in a Node.js environment (e.g., serverless function)
// It won't work directly in the browser

export const runPythonProcessor = async (
  fileBuffer: Buffer,
  fileName: string,
  assemblyId: string,
  assemblyName: string,
  figureStart: number,
  figureEnd: number,
  logoPath?: string
): Promise<{
  success: boolean;
  message: string;
  tasks?: Task[];
  zipBuffer?: Buffer;
  excelBuffer?: Buffer;
}> => {
  try {
    // Create temp directory
    const tempDir = fs.mkdtempSync(path.join(os.tmpdir(), 'sop-processor-'));
    const tempFilePath = path.join(tempDir, fileName);
    
    // Write file to temp location
    fs.writeFileSync(tempFilePath, fileBuffer);
    
    // Arguments for Python script
    const args = [
      tempFilePath,
      assemblyId,
      assemblyName,
      figureStart.toString(),
      figureEnd.toString()
    ];
    
    // Add logo path if provided
    if (logoPath) {
      args.push(logoPath);
    }
    
    // Run Python script
    const options = {
      mode: 'text' as 'text',
      pythonPath: 'python3',
      pythonOptions: ['-u'],
      scriptPath: path.join(__dirname, '../services'),
      args: args
    };
    
    console.log('Running Python processor with options:', JSON.stringify(options));
    
    const result = await PythonShell.run('pythonProcessor.py', options);
    
    // Parse result (assuming the Python script returns JSON)
    const jsonResult = JSON.parse(result[0]);
    
    console.log('Python processor result:', jsonResult);
    
    // Read the ZIP file if processing was successful
    let zipBuffer;
    let excelBuffer;
    
    if (jsonResult.success) {
      if (jsonResult.zip_path) {
        zipBuffer = fs.readFileSync(jsonResult.zip_path);
      }
      
      if (jsonResult.excel_path) {
        excelBuffer = fs.readFileSync(jsonResult.excel_path);
      }
    }
    
    // Clean up temp files
    fs.rmSync(tempDir, { recursive: true, force: true });
    
    return {
      success: jsonResult.success,
      message: jsonResult.message,
      tasks: jsonResult.tasks,
      zipBuffer,
      excelBuffer
    };
  } catch (error) {
    console.error('Error running Python processor:', error);
    return {
      success: false,
      message: `Error running Python processor: ${(error as Error).message}`
    };
  }
};
