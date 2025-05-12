
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
  figureEnd: number
): Promise<{
  success: boolean;
  message: string;
  tasks?: Task[];
  zipBuffer?: Buffer;
}> => {
  try {
    // Create temp directory
    const tempDir = fs.mkdtempSync(path.join(os.tmpdir(), 'sop-processor-'));
    const tempFilePath = path.join(tempDir, fileName);
    
    // Write file to temp location
    fs.writeFileSync(tempFilePath, fileBuffer);
    
    // Run Python script
    const options = {
      mode: 'text',
      pythonPath: 'python3', // Make sure Python is installed on the server
      pythonOptions: ['-u'], // Unbuffered output
      scriptPath: path.join(__dirname, '../services'),
      args: [
        tempFilePath,
        assemblyId,
        assemblyName,
        figureStart.toString(),
        figureEnd.toString()
      ]
    };
    
    const result = await PythonShell.run('pythonProcessor.py', options);
    
    // Parse result (assuming the Python script returns JSON)
    const jsonResult = JSON.parse(result[0]);
    
    // Read the ZIP file if processing was successful
    let zipBuffer;
    if (jsonResult.success && jsonResult.zip_path) {
      zipBuffer = fs.readFileSync(jsonResult.zip_path);
    }
    
    // Clean up temp files
    fs.rmSync(tempDir, { recursive: true, force: true });
    
    return {
      success: jsonResult.success,
      message: jsonResult.message,
      tasks: jsonResult.tasks,
      zipBuffer
    };
  } catch (error) {
    console.error('Error running Python processor:', error);
    return {
      success: false,
      message: `Error running Python processor: ${(error as Error).message}`
    };
  }
};
