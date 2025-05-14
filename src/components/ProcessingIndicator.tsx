
import React from 'react';
import { Progress } from '@/components/ui/progress';
import { FileText, CheckCircle, AlertCircle } from 'lucide-react';

export type ProcessingStatus = 'parsing' | 'extracting' | 'generating' | 'complete' | 'error';

export interface ProcessingIndicatorProps {
  status: ProcessingStatus;
  progress: number;
  error?: string;
}

const ProcessingIndicator = ({ status = 'parsing', progress = 50, error }: ProcessingIndicatorProps) => {
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

  const getProgressColor = () => {
    if (status === 'complete') return 'bg-green-500';
    if (status === 'error') return 'bg-red-500';
    return 'bg-sop-blue';
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
      
      <div className="relative w-full h-2 mb-2 bg-gray-200 rounded-full overflow-hidden">
        <div 
          className={`absolute top-0 left-0 h-full ${getProgressColor()} transition-all duration-300`}
          style={{ width: `${progress}%` }}
        />
      </div>
      
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
