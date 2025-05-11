
import React, { useState } from 'react';
import { Button } from '@/components/ui/button';
import { Card, CardContent } from '@/components/ui/card';
import { Download, HelpCircle, FileText } from 'lucide-react';
import { useToast } from '@/components/ui/use-toast';
import FileUploader from '@/components/FileUploader';
import ProcessingIndicator from '@/components/ProcessingIndicator';
import TaskPreview, { Task } from '@/components/TaskPreview';
import { processDocument, generateTaskMasterDocument, createDownloadPackage } from '@/services/docxProcessor';
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
    
    try {
      // Process the document
      setProgress(30);
      const extractedContent = await processDocument(selectedFile);
      
      // Update state with extracted content
      setStatus('extracting');
      setProgress(60);
      setTasks(extractedContent.tasks);
      setDocTitle(extractedContent.docTitle);
      
      // Generate the Task Master document
      setStatus('generating');
      setProgress(80);
      const docBlob = generateTaskMasterDocument(
        extractedContent.docTitle,
        extractedContent.tasks
      );
      
      // Create downloadable package
      const zipBlob = await createDownloadPackage(
        docBlob,
        extractedContent.images,
        extractedContent.docTitle
      );
      setDownloadPackage(zipBlob);
      
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
      setErrorMessage((error as Error).message);
      
      toast({
        variant: "destructive",
        title: "Processing failed",
        description: (error as Error).message
      });
    }
  };

  // Handle package download
  const handleDownload = () => {
    if (downloadPackage) {
      saveAs(downloadPackage, `${docTitle} - SOP Package.zip`);
      toast({
        title: "Download started",
        description: "Your SOP package is being downloaded"
      });
    }
  };

  return (
    <div className="min-h-screen bg-gray-50">
      {/* Header */}
      <header className="bg-sop-blue text-white py-4">
        <div className="container mx-auto px-4">
          <div className="flex justify-between items-center">
            <div className="flex items-center space-x-2">
              <FileText size={24} />
              <h1 className="text-xl font-bold">SOP Task Master Generator</h1>
            </div>
          </div>
        </div>
      </header>
      
      {/* Main content */}
      <main className="container mx-auto px-4 py-8">
        <div className="max-w-4xl mx-auto">
          <Card className="mb-8 overflow-hidden">
            <div className="bg-sop-gray p-4 border-b">
              <h2 className="text-lg font-medium flex items-center">
                <HelpCircle size={18} className="mr-2 text-sop-blue" />
                How it works
              </h2>
            </div>
            <CardContent className="p-4 space-y-2 text-sm">
              <p>
                1. Upload your SOP Word document (.docx) containing step-by-step procedures.
              </p>
              <p>
                2. The application will extract each step, task details, and images from the document.
              </p>
              <p>
                3. A Task Master document will be generated with proper formatting.
              </p>
              <p>
                4. Images will be extracted and renamed according to the task numbers.
              </p>
              <p>
                5. Download the package containing the Task Master document and images.
              </p>
            </CardContent>
          </Card>
          
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
              <h2 className="text-xl font-semibold text-gray-800">
                Extracted Tasks Preview
              </h2>
              
              <TaskPreview tasks={tasks} documentTitle={docTitle} />
              
              {/* Download button */}
              {status === 'complete' && downloadPackage && (
                <div className="flex justify-center mt-6">
                  <Button 
                    onClick={handleDownload}
                    className="bg-sop-blue hover:bg-sop-lightBlue px-6 py-2"
                  >
                    <Download className="mr-2 h-4 w-4" />
                    Download SOP Package
                  </Button>
                </div>
              )}
              
              {/* Upload another file button */}
              {status === 'complete' && (
                <div className="flex justify-center mt-4">
                  <Button 
                    variant="outline"
                    onClick={() => setStatus('idle')}
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
      <footer className="bg-gray-100 py-4 mt-12">
        <div className="container mx-auto px-4 text-center text-sm text-gray-600">
          SOP Task Master Generator - Streamline your standard operating procedures
        </div>
      </footer>
    </div>
  );
};

export default Index;
