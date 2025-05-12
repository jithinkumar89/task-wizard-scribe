
import React, { useState } from 'react';
import { Button } from '@/components/ui/button';
import { Upload } from 'lucide-react';
import { useToast } from '@/components/ui/use-toast';

interface FileUploaderProps {
  onFileSelect: (file: File) => void;
  isProcessing?: boolean;
}

const FileUploader = ({ onFileSelect, isProcessing = false }: FileUploaderProps) => {
  const [isDragging, setIsDragging] = useState(false);
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
      processFile(e.dataTransfer.files[0]);
    }
  };

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files?.length) {
      processFile(e.target.files[0]);
    }
  };

  const processFile = (file: File) => {
    if (file.type !== 'application/vnd.openxmlformats-officedocument.wordprocessingml.document') {
      toast({
        variant: "destructive",
        title: "Invalid file format",
        description: "Please upload a Word document (.docx file)"
      });
      return;
    }
    
    onFileSelect(file);
    toast({
      title: "File uploaded",
      description: `"${file.name}" has been uploaded successfully.`
    });
  };

  return (
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
  );
};

export default FileUploader;
