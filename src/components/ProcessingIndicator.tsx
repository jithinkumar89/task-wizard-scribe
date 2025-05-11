
import React from 'react';
import { Progress } from '@/components/ui/progress';
import { FileText, CheckCircle, AlertCircle } from 'lucide-react';

type ProcessingStatus = 'parsing' | 'extracting' | 'generating' | 'complete' | 'error';

interface ProcessingIndicatorProps {
  status: ProcessingStatus;
  progress: number;
  error?: string;
}

const ProcessingIndicator = ({ status, progress, error }: ProcessingIndicatorProps) => {
  const getStatusText = () => {
    switch (status) {
      case 'parsing':
        return 'Parsing document...';
      case 'extracting':
        return 'Extracting images...';
      case 'generating':
        return 'Generating Task Master document...';
      case 'complete':
        return 'Processing complete!';
      case 'error':
        return 'Error processing document';
      default:
        return 'Processing...';
    }
  };

  return (
    <div className="w-full p-6 bg-white rounded-lg shadow-sm">
      <div className="flex items-center mb-4">
        {status === 'complete' ? (
          <CheckCircle className="w-6 h-6 text-green-500 mr-2" />
        ) : status === 'error' ? (
          <AlertCircle className="w-6 h-6 text-red-500 mr-2" />
        ) : (
          <FileText className="w-6 h-6 text-sop-blue mr-2" />
        )}
        <h3 className="text-lg font-medium">{getStatusText()}</h3>
      </div>
      
      <Progress 
        value={progress} 
        className="h-2 mb-2"
        indicatorClassName={
          status === 'complete' ? 'bg-green-500' : 
          status === 'error' ? 'bg-red-500' :
          'bg-sop-blue'
        }
      />
      
      <div className="text-sm text-gray-500">
        {status === 'error' ? (
          <p className="text-red-500">{error || 'An error occurred during processing'}</p>
        ) : (
          <p>{Math.round(progress)}% complete</p>
        )}
      </div>
    </div>
  );
};

export default ProcessingIndicator;
