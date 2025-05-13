
import React, { useState } from 'react';
import { Button } from '@/components/ui/button';
import { Card, CardContent } from '@/components/ui/card';
import { Download, HelpCircle, FileText } from 'lucide-react';
import { useToast } from '@/hooks/use-toast';
import { Input } from '@/components/ui/input';
import { Label } from '@/components/ui/label';
import FileUploader from '@/components/FileUploader';
import ProcessingIndicator from '@/components/ProcessingIndicator';
import TaskPreview, { Task } from '@/components/TaskPreview';
import { processDocument, generateExcelFile, createDownloadPackage } from '@/services/docxProcessor';
import { processPythonDocument } from '@/services/pythonBridge';
import { saveAs } from 'file-saver';

type ProcessingStatus = 'idle' | 'parsing' | 'extracting' | 'generating' | 'complete' | 'error';

const Index = () => {
  const [file, setFile] = useState<File | null>(null);
  const [status, setStatus] = useState<ProcessingStatus>('idle');
  const [progress, setProgress] = useState(0);
  const [tasks, setTasks] = useState<Task[]>([]);
  const [docTitle, setDocTitle] = useState('');
  const [errorMessage, setErrorMessage] = useState('');
  const [downloadPackage, setDownloadPackage] = useState<Blob | null>(null);
  const [excelFile, setExcelFile] = useState<Blob | null>(null);
  const [assemblySequenceId, setAssemblySequenceId] = useState<string>('1');
  const [assemblyName, setAssemblyName] = useState<string>('');
  const [figureStartRange, setFigureStartRange] = useState<string>('1');
  const [figureEndRange, setFigureEndRange] = useState<string>('10');
  const [useTableExtraction, setUseTableExtraction] = useState<boolean>(true);
  const { toast } = useToast();

  // Handle file upload
  const handleFileSelect = async (selectedFile: File) => {
    setFile(selectedFile);
    setStatus('parsing');
    setProgress(10);
    setTasks([]);
    setDocTitle('');
    setErrorMessage('');
    setDownloadPackage(null);
    setExcelFile(null);
    
    if (!assemblySequenceId || isNaN(Number(assemblySequenceId))) {
      toast({
        variant: "destructive",
        title: "Invalid Assembly Sequence ID",
        description: "Please enter a valid number for the Assembly Sequence ID"
      });
      setStatus('idle');
      return;
    }

    if (!assemblyName.trim()) {
      toast({
        variant: "destructive",
        title: "Missing Assembly Name",
        description: "Please enter a name for the assembly to use as the description"
      });
      setStatus('idle');
      return;
    }
    
    try {
      // Process the document
      setProgress(30);
      console.log('Processing document:', selectedFile.name, 'with Assembly Sequence ID:', assemblySequenceId);
      
      let extractedContent;
      
      if (useTableExtraction) {
        // Try using Python-based table extraction (this will throw an error in browser environment)
        try {
          const pythonResult = await processPythonDocument(
            selectedFile, 
            assemblySequenceId, 
            assemblyName,
            parseInt(figureStartRange, 10),
            parseInt(figureEndRange, 10)
          );
          
          extractedContent = {
            docTitle: pythonResult.docTitle,
            tasks: pythonResult.tasks,
            images: [] // Images are handled differently in Python implementation
          };
          
          // Set download package directly from Python result
          setDownloadPackage(pythonResult.zipPackage);
          setExcelFile(pythonResult.excelFile);
        } catch (error) {
          console.log("Python processing unavailable, falling back to JavaScript implementation:", error);
          // Fall back to JavaScript implementation
          extractedContent = await processDocument(selectedFile, assemblySequenceId);
          
          // Update all tasks to have the specified assembly name as description
          if (extractedContent.tasks && extractedContent.tasks.length > 0) {
            extractedContent.tasks = extractedContent.tasks.map(task => ({
              ...task,
              description: assemblyName
            }));
          }
        }
      } else {
        // Use JavaScript implementation directly
        extractedContent = await processDocument(selectedFile, assemblySequenceId);
        
        // Update all tasks to have the specified assembly name as description
        if (extractedContent.tasks && extractedContent.tasks.length > 0) {
          extractedContent.tasks = extractedContent.tasks.map(task => ({
            ...task,
            description: assemblyName
          }));
        }
      }
      
      // Update state with extracted content
      setStatus('extracting');
      setProgress(60);
      
      // Check if we have tasks
      if (!extractedContent.tasks || extractedContent.tasks.length === 0) {
        console.error('No tasks were extracted from the document');
        throw new Error('No tasks could be extracted from the document. Please check the file format or ensure it contains numbered steps or a table structure.');
      }
      
      setTasks(extractedContent.tasks);
      setDocTitle(extractedContent.docTitle || assemblyName || 'Unnamed Document');
      
      console.log(`Successfully extracted ${extractedContent.tasks.length} tasks`);
      
      // Generate the Excel file if not already done by Python implementation
      if (!excelFile) {
        setStatus('generating');
        setProgress(80);
        
        try {
          // Generate Excel file
          const excelBlob = await generateExcelFile(extractedContent.tasks, assemblyName || extractedContent.docTitle);
          setExcelFile(excelBlob);
          
          // Create downloadable package
          const zipBlob = await createDownloadPackage(
            excelBlob,
            extractedContent.images || [],
            assemblyName || extractedContent.docTitle || 'Unnamed Document'
          );
          setDownloadPackage(zipBlob);
        } catch (error) {
          console.error("Error generating download files:", error);
          throw new Error("Failed to create download files. Please try again or check file format.");
        }
      }
      
      // Complete
      setStatus('complete');
      setProgress(100);
      
      toast({
        title: "Processing complete",
        description: `${extractedContent.tasks.length} tasks extracted successfully.`
      });
    } catch (error) {
      console.error('Error processing document:', error);
      setStatus('error');
      setProgress(0);
      setErrorMessage((error as Error).message || 'Unknown error occurred');
      
      toast({
        variant: "destructive",
        title: "Processing failed",
        description: (error as Error).message || 'Unknown error occurred'
      });
    }
  };

  // Handle package download
  const handleDownload = () => {
    if (downloadPackage) {
      saveAs(downloadPackage, `${assemblyName || docTitle || 'Task_Master'}_extracted_data.zip`);
      toast({
        title: "Download started",
        description: "Your SOP package is being downloaded"
      });
    }
  };
  
  // Handle Excel-only download
  const handleExcelDownload = () => {
    if (excelFile) {
      saveAs(excelFile, `${assemblyName || docTitle || 'Task_Master'}.xlsx`);
      toast({
        title: "Excel download started",
        description: "Your task Excel file is being downloaded"
      });
    }
  };
  
  // Handle images-only download
  const handleImagesDownload = () => {
    if (downloadPackage) {
      // The createDownloadPackage already creates a zip with images folder
      saveAs(downloadPackage, `${assemblyName || docTitle || 'Task_Master'}_images.zip`);
      toast({
        title: "Images download started",
        description: "Your images package is being downloaded"
      });
    }
  };

  return (
    <div className="min-h-screen bg-gradient-to-br from-blue-50 to-indigo-50">
      {/* Header */}
      <header className="bg-gradient-to-r from-blue-600 to-indigo-700 text-white py-6 shadow-lg">
        <div className="container mx-auto px-4">
          <div className="flex justify-between items-center">
            <div className="flex items-center space-x-2">
              <FileText size={30} className="text-white/90" />
              <h1 className="text-2xl font-bold">SOP Task Master Generator</h1>
            </div>
            <img src="/lovable-uploads/1ac64f9b-f851-4336-8290-0ae34c0deb10.png" alt="BPL Medical Technologies Logo" className="h-10" />
          </div>
        </div>
      </header>
      
      {/* Main content */}
      <main className="container mx-auto px-4 py-10">
        <div className="max-w-4xl mx-auto">
          <Card className="mb-8 overflow-hidden shadow-lg border-0 bg-white/80 backdrop-blur-sm">
            <div className="bg-gradient-to-r from-indigo-100 to-blue-100 p-4 border-b">
              <h2 className="text-lg font-medium flex items-center text-gray-800">
                <HelpCircle size={20} className="mr-2 text-blue-600" />
                How it works
              </h2>
            </div>
            <CardContent className="p-6 space-y-3 text-sm">
              <p className="flex items-center">
                <span className="w-6 h-6 rounded-full bg-gradient-to-br from-blue-600 to-indigo-600 text-white flex items-center justify-center mr-2 text-xs font-bold">1</span>
                Enter the Assembly Sequence ID (e.g., 1, 2, 3...) and Assembly Name.
              </p>
              <p className="flex items-center">
                <span className="w-6 h-6 rounded-full bg-gradient-to-br from-blue-600 to-indigo-600 text-white flex items-center justify-center mr-2 text-xs font-bold">2</span>
                If your document contains figure references, enter the figure range (e.g., 1-10).
              </p>
              <p className="flex items-center">
                <span className="w-6 h-6 rounded-full bg-gradient-to-br from-blue-600 to-indigo-600 text-white flex items-center justify-center mr-2 text-xs font-bold">3</span>
                Upload your SOP Word document (.docx) containing a table with tasks.
              </p>
              <p className="flex items-center">
                <span className="w-6 h-6 rounded-full bg-gradient-to-br from-blue-600 to-indigo-600 text-white flex items-center justify-center mr-2 text-xs font-bold">4</span>
                The application will extract each task and assign task numbers in the format "{assemblySequenceId}.0.001".
              </p>
              <p className="flex items-center">
                <span className="w-6 h-6 rounded-full bg-gradient-to-br from-blue-600 to-indigo-600 text-white flex items-center justify-center mr-2 text-xs font-bold">5</span>
                Images will be extracted and renamed according to the figure references in the document.
              </p>
              <p className="flex items-center">
                <span className="w-6 h-6 rounded-full bg-gradient-to-br from-blue-600 to-indigo-600 text-white flex items-center justify-center mr-2 text-xs font-bold">6</span>
                Download the complete package, Excel file or just the images separately.
              </p>
            </CardContent>
          </Card>
          
          {/* Input form */}
          {status === 'idle' && (
            <div className="mb-6 bg-white/80 backdrop-blur-sm p-6 rounded-lg shadow-md border border-blue-100 space-y-5">
              <div>
                <Label htmlFor="sequenceId" className="text-sm font-medium text-gray-700">
                  Assembly Sequence ID
                </Label>
                <div className="flex mt-2 gap-2 items-center">
                  <Input
                    id="sequenceId"
                    type="number"
                    min="1"
                    value={assemblySequenceId}
                    onChange={(e) => setAssemblySequenceId(e.target.value)}
                    className="w-32 border-blue-200 focus:border-blue-400"
                    placeholder="e.g., 1"
                  />
                  <p className="text-sm text-gray-500">
                    This will be used to prefix task numbers (e.g., {assemblySequenceId}.0.001)
                  </p>
                </div>
              </div>
              
              <div>
                <Label htmlFor="assemblyName" className="text-sm font-medium text-gray-700">
                  Assembly Name
                </Label>
                <div className="flex mt-2 gap-2 items-center">
                  <Input
                    id="assemblyName"
                    type="text"
                    value={assemblyName}
                    onChange={(e) => setAssemblyName(e.target.value)}
                    className="w-full border-blue-200 focus:border-blue-400"
                    placeholder="e.g., Engine Assembly"
                  />
                </div>
                <p className="text-sm text-gray-500 mt-1">
                  This will be used as the description for all tasks and for output files naming
                </p>
              </div>
              
              <div>
                <Label className="text-sm font-medium text-gray-700">
                  Figure Reference Range
                </Label>
                <div className="flex mt-2 gap-2 items-center">
                  <Input
                    id="figureStartRange"
                    type="number"
                    min="1"
                    value={figureStartRange}
                    onChange={(e) => setFigureStartRange(e.target.value)}
                    className="w-24 border-blue-200 focus:border-blue-400"
                    placeholder="Start"
                  />
                  <span>to</span>
                  <Input
                    id="figureEndRange"
                    type="number"
                    min="1"
                    value={figureEndRange}
                    onChange={(e) => setFigureEndRange(e.target.value)}
                    className="w-24 border-blue-200 focus:border-blue-400"
                    placeholder="End"
                  />
                  <p className="text-sm text-gray-500">
                    If your document uses "Figure X" references
                  </p>
                </div>
              </div>
            </div>
          )}
          
          {/* File uploader */}
          {status === 'idle' && (
            <FileUploader onFileSelect={handleFileSelect} />
          )}
          
          {/* Processing indicator */}
          {['parsing', 'extracting', 'generating', 'complete', 'error'].includes(status) && (
            <div className="mb-8">
              <ProcessingIndicator 
                status={status as any}
                progress={progress} 
                error={errorMessage}
              />
            </div>
          )}
          
          {/* Results section */}
          {tasks.length > 0 && (
            <div className="space-y-6">
              <h2 className="text-xl font-semibold text-gray-800 flex items-center">
                <span className="w-2 h-8 bg-gradient-to-b from-blue-600 to-indigo-600 rounded mr-2"></span>
                Extracted Tasks Preview
              </h2>
              
              <TaskPreview tasks={tasks} documentTitle={assemblyName || docTitle} />
              
              {/* Download buttons */}
              {status === 'complete' && downloadPackage && (
                <div className="flex flex-wrap justify-center mt-8 gap-4">
                  <Button 
                    onClick={handleDownload}
                    className="bg-gradient-to-r from-blue-600 to-indigo-600 hover:from-blue-700 hover:to-indigo-700 px-6 py-5 shadow-lg transition-all duration-300 hover:shadow-xl"
                  >
                    <Download className="mr-2 h-4 w-4" />
                    Download Complete Package
                  </Button>
                  
                  <Button 
                    onClick={handleExcelDownload}
                    variant="outline"
                    className="px-6 py-5 border-blue-300 text-blue-700 hover:bg-blue-50 transition-colors duration-300"
                  >
                    <Download className="mr-2 h-4 w-4" />
                    Download Excel File
                  </Button>
                  
                  <Button 
                    onClick={handleImagesDownload}
                    variant="outline"
                    className="px-6 py-5 border-blue-300 text-blue-700 hover:bg-blue-50 transition-colors duration-300"
                  >
                    <Download className="mr-2 h-4 w-4" />
                    Download Images Only
                  </Button>
                </div>
              )}
              
              {/* Upload another file button */}
              {status === 'complete' && (
                <div className="flex justify-center mt-6">
                  <Button 
                    variant="outline"
                    onClick={() => setStatus('idle')}
                    className="border-blue-300 text-blue-700 hover:bg-blue-50 transition-colors duration-300"
                  >
                    Process Another Document
                  </Button>
                </div>
              )}
            </div>
          )}
        </div>
      </main>
      
      {/* Footer */}
      <footer className="bg-gradient-to-r from-gray-100 to-blue-100 py-6 mt-12">
        <div className="container mx-auto px-4 text-center text-sm text-gray-600 flex justify-center items-center space-x-2">
          <span>SOP Task Master Generator - BPL Medical Technologies</span>
          <img src="/lovable-uploads/1ac64f9b-f851-4336-8290-0ae34c0deb10.png" alt="BPL Logo" className="h-5" />
        </div>
      </footer>
    </div>
  );
};

export default Index;
