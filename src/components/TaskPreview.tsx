
import React from 'react';
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from '@/components/ui/table';
import { Card, CardContent, CardHeader, CardTitle } from '@/components/ui/card';
import { Badge } from '@/components/ui/badge';

export interface Task {
  task_no: string;
  type: string;
  eta_sec: string;
  description: string;
  activity: string;
  specification: string;
  attachment: string;
  hasImage?: boolean;  // Optional for backward compatibility
  
  // Legacy properties for compatibility with existing code
  taskNumber?: string;
  etaSec?: string;
}

interface TaskPreviewProps {
  tasks: Task[];
  documentTitle: string;
}

const TaskPreview = ({ tasks, documentTitle }: TaskPreviewProps) => {
  if (tasks.length === 0) {
    return null;
  }

  return (
    <Card className="w-full">
      <CardHeader className="bg-gradient-to-r from-blue-600 via-indigo-500 to-purple-600 text-white">
        <CardTitle className="text-white">{documentTitle || 'Task Master Preview'}</CardTitle>
      </CardHeader>
      <CardContent className="p-0 overflow-auto max-h-[500px]">
        <Table>
          <TableHeader className="bg-gray-50 sticky top-0">
            <TableRow>
              <TableHead className="w-20">Task No</TableHead>
              <TableHead className="w-24">Type</TableHead>
              <TableHead className="w-24">ETA (sec)</TableHead>
              <TableHead className="w-48">Description</TableHead>
              <TableHead className="w-48">Activity</TableHead>
              <TableHead className="w-48">Specification</TableHead>
              <TableHead className="w-32">Attachment</TableHead>
            </TableRow>
          </TableHeader>
          <TableBody>
            {tasks.map((task, index) => (
              <TableRow key={index}>
                <TableCell className="font-medium">{task.task_no}</TableCell>
                <TableCell>{task.type}</TableCell>
                <TableCell>{task.eta_sec}</TableCell>
                <TableCell>{task.description}</TableCell>
                <TableCell className="whitespace-pre-wrap max-w-48">{task.activity}</TableCell>
                <TableCell>{task.specification}</TableCell>
                <TableCell>
                  {task.attachment ? (
                    <div className="flex flex-wrap gap-1">
                      {task.attachment.split(',').map((id, idx) => (
                        <Badge key={idx} variant="outline" className="bg-green-50 text-green-700 border-green-200">
                          {id.trim()}
                        </Badge>
                      ))}
                    </div>
                  ) : (
                    '-'
                  )}
                </TableCell>
              </TableRow>
            ))}
          </TableBody>
        </Table>
      </CardContent>
    </Card>
  );
};

export default TaskPreview;
