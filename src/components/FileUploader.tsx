
import React, { useState } from 'react';
import { Button } from '@/components/ui/button';
import { Upload, File, Check } from 'lucide-react';
import { useToast } from '@/components/ui/use-toast';

interface FileUploaderProps {
  onFileSelect: (file: File) => void;
  isProcessing?: boolean;
}

const FileUploader = ({ onFileSelect, isProcessing = false }: FileUploaderProps) => {
  const [isDragging, setIsDragging] = useState(false);
  const [selectedFile, setSelectedFile] = useState<File | null>(null);
  const { toast } = useToast();
  
  const handleDragOver = (e: React.DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    setIsDragging(true);
  };

  const handleDragLeave = () => {
    setIsDragging(false);
  };

  const handleDrop = (e: React.DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    setIsDragging(false);
    
    if (e.dataTransfer.files.length) {
      validateAndSetFile(e.dataTransfer.files[0]);
    }
  };

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files?.length) {
      validateAndSetFile(e.target.files[0]);
    }
  };

  const validateAndSetFile = (file: File) => {
    if (file.type !== 'application/vnd.openxmlformats-officedocument.wordprocessingml.document') {
      toast({
        variant: "destructive",
        title: "Invalid file format",
        description: "Please upload a Word document (.docx file)"
      });
      return;
    }
    
    setSelectedFile(file);
    toast({
      title: "File selected",
      description: `"${file.name}" has been selected. Click "Process Document" to continue.`
    });
  };

  const handleProcess = () => {
    if (selectedFile) {
      onFileSelect(selectedFile);
    }
  };

  return (
    <div className="space-y-4">
      <div 
        className={`border-2 border-dashed rounded-lg p-8 text-center ${
          isDragging ? 'bg-sop-gray border-sop-blue' : 'border-gray-300'
        } transition-colors duration-200`} 
        onDragOver={handleDragOver}
        onDragLeave={handleDragLeave}
        onDrop={handleDrop}
      >
        <div className="flex flex-col items-center justify-center space-y-4">
          <Upload 
            className="h-12 w-12 text-sop-blue" 
            strokeWidth={1.5}
          />
          <h3 className="text-lg font-medium">Upload SOP Word Document</h3>
          <p className="text-sm text-gray-500 max-w-sm">
            Drag and drop your file here, or click to browse
          </p>
          <input
            type="file"
            id="file-upload"
            className="hidden"
            accept=".docx"
            onChange={handleFileChange}
            disabled={isProcessing}
          />
          <Button 
            onClick={() => document.getElementById('file-upload')?.click()}
            className="bg-sop-blue hover:bg-sop-lightBlue text-white"
            disabled={isProcessing}
          >
            {isProcessing ? 'Processing...' : 'Select File'}
          </Button>
        </div>
      </div>
      
      {selectedFile && (
        <div className="border rounded-lg p-4 bg-gray-50">
          <div className="flex items-center justify-between">
            <div className="flex items-center space-x-3">
              <File className="h-8 w-8 text-sop-blue" />
              <div>
                <p className="font-medium">{selectedFile.name}</p>
                <p className="text-sm text-gray-500">
                  {(selectedFile.size / 1024).toFixed(1)} KB
                </p>
              </div>
            </div>
            <Button 
              onClick={handleProcess}
              className="bg-sop-blue hover:bg-sop-lightBlue text-white"
              disabled={isProcessing}
            >
              {isProcessing ? (
                <>Processing...</>
              ) : (
                <>
                  <Check className="mr-1" />
                  Process Document
                </>
              )}
            </Button>
          </div>
        </div>
      )}
    </div>
  );
};

export default FileUploader;
