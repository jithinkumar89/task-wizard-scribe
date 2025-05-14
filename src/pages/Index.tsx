
import { useState, useRef } from 'react';
import { Button } from '@/components/ui/button';
import { Input } from '@/components/ui/input';
import { Label } from '@/components/ui/label';
import { Card, CardContent, CardHeader, CardTitle } from '@/components/ui/card';
import FileUploader from '@/components/FileUploader';
import ProcessingIndicator from '@/components/ProcessingIndicator';
import TaskPreview from '@/components/TaskPreview';
import { processDocument, generateExcelFile, createDownloadPackage } from '@/services/docxProcessor';
import { toast } from '@/hooks/use-toast';
import Footer from '@/components/Footer';
import { saveAs } from 'file-saver';
import { Task } from '@/components/TaskPreview';
import { 
  Select,
  SelectContent,
  SelectItem,
  SelectTrigger,
  SelectValue,
} from "@/components/ui/select";

const Index = () => {
  const [assemblySequenceId, setAssemblySequenceId] = useState('1');
  const [assemblyName, setAssemblyName] = useState('');
  const [figureStartRange, setFigureStartRange] = useState(0);
  const [figureEndRange, setFigureEndRange] = useState(999);
  const [type, setType] = useState('');
  const [file, setFile] = useState<File | null>(null);
  const [isProcessing, setIsProcessing] = useState(false);
  const [tasks, setTasks] = useState<Task[]>([]);
  const [docTitle, setDocTitle] = useState('');
  const fileInputRef = useRef<HTMLInputElement>(null);

  const resetForm = () => {
    setAssemblySequenceId('1');
    setAssemblyName('');
    setFigureStartRange(0);
    setFigureEndRange(999);
    setType('');
    setFile(null);
    setTasks([]);
    setDocTitle('');
    if (fileInputRef.current) {
      fileInputRef.current.value = '';
    }
  };

  const handleFileChange = (file: File | null) => {
    setFile(file);
    // Reset tasks preview when a new file is selected
    setTasks([]);
    setDocTitle('');
  };

  const handleProcessDocument = async () => {
    if (!file) {
      toast({
        title: 'No File Selected',
        description: 'Please select a document to process.',
        variant: 'destructive',
      });
      return;
    }

    if (!assemblyName.trim()) {
      toast({
        title: 'Missing Assembly Name',
        description: 'Please enter an assembly name.',
        variant: 'destructive',
      });
      return;
    }

    try {
      setIsProcessing(true);
      const result = await processDocument(file, assemblySequenceId, type);
      setTasks(result.tasks);
      setDocTitle(result.docTitle || assemblyName);
      
      toast({
        title: 'Document Processed',
        description: `Extracted ${result.tasks.length} tasks and ${result.images.length} images.`,
      });
    } catch (error) {
      console.error('Error processing document:', error);
      toast({
        title: 'Processing Error',
        description: (error as Error).message,
        variant: 'destructive',
      });
    } finally {
      setIsProcessing(false);
    }
  };

  const handleDownload = async () => {
    if (!file || tasks.length === 0) {
      toast({
        title: 'Nothing to Download',
        description: 'Please process a document first.',
        variant: 'destructive',
      });
      return;
    }

    try {
      setIsProcessing(true);
      
      // Re-process the document to ensure we have the latest data
      const processedData = await processDocument(file, assemblySequenceId, type);
      
      // Generate Excel file with all data
      const excelBlob = await generateExcelFile(
        processedData.tasks, 
        processedData.docTitle || assemblyName,
        processedData.toolsData,
        processedData.imtData
      );
      
      // Create download package with Excel and images
      const zipBlob = await createDownloadPackage(
        excelBlob,
        processedData.images,
        processedData.docTitle || assemblyName
      );
      
      // Download the ZIP package
      saveAs(zipBlob, `${assemblyName}_processed.zip`);
      
      toast({
        title: 'Download Ready',
        description: 'Your processed document package has been downloaded.',
      });
    } catch (error) {
      console.error('Error generating download:', error);
      toast({
        title: 'Download Error',
        description: (error as Error).message,
        variant: 'destructive',
      });
    } finally {
      setIsProcessing(false);
    }
  };

  return (
    <div className="container mx-auto p-4 min-h-screen flex flex-col">
      <header className="text-center mb-8">
        <div className="inline-block bg-white p-3 rounded-xl shadow-md">
          <img 
            src="/lovable-uploads/1ac64f9b-f851-4336-8290-0ae34c0deb10.png" 
            alt="SOP Processor Logo" 
            className="h-16"
          />
        </div>
        <h1 className="text-2xl font-bold mt-4 bg-gradient-to-r from-blue-600 via-indigo-500 to-purple-600 bg-clip-text text-transparent">
          SOP Processor
        </h1>
      </header>

      <main className="flex-grow">
        <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
          <Card>
            <CardHeader className="bg-gradient-to-r from-blue-600 via-indigo-500 to-purple-600 text-white">
              <CardTitle className="text-white">Upload Document</CardTitle>
            </CardHeader>
            <CardContent className="space-y-4 pt-6">
              <div>
                <Label htmlFor="assemblySequenceId">Assembly Sequence ID</Label>
                <Input
                  id="assemblySequenceId"
                  placeholder="e.g., 1"
                  value={assemblySequenceId}
                  onChange={(e) => setAssemblySequenceId(e.target.value)}
                />
              </div>
              
              <div>
                <Label htmlFor="assemblyName">Assembly Name</Label>
                <Input
                  id="assemblyName"
                  placeholder="e.g., Unit Assembly"
                  value={assemblyName}
                  onChange={(e) => setAssemblyName(e.target.value)}
                />
              </div>
              
              <div>
                <Label htmlFor="type">Type</Label>
                <Select value={type} onValueChange={setType}>
                  <SelectTrigger>
                    <SelectValue placeholder="e.g., Operation, QC, Approval" />
                  </SelectTrigger>
                  <SelectContent>
                    <SelectItem value="Operation">Operation</SelectItem>
                    <SelectItem value="QC">QC</SelectItem>
                    <SelectItem value="Approval">Approval</SelectItem>
                  </SelectContent>
                </Select>
              </div>

              <div className="grid grid-cols-2 gap-4">
                <div>
                  <Label htmlFor="figureStartRange">Figure Start Range</Label>
                  <Input
                    id="figureStartRange"
                    type="number"
                    min="0"
                    placeholder="0"
                    value={figureStartRange}
                    onChange={(e) => setFigureStartRange(Number(e.target.value))}
                  />
                </div>
                <div>
                  <Label htmlFor="figureEndRange">Figure End Range</Label>
                  <Input
                    id="figureEndRange"
                    type="number"
                    min="0"
                    placeholder="999"
                    value={figureEndRange}
                    onChange={(e) => setFigureEndRange(Number(e.target.value))}
                  />
                </div>
              </div>

              <div>
                <Label>Document Upload (.docx only)</Label>
                <FileUploader 
                  onFileSelect={handleFileChange} 
                  accept=".docx" 
                  ref={fileInputRef}
                />
              </div>

              <div className="flex flex-col sm:flex-row gap-2 pt-4">
                <Button 
                  className="flex-1 bg-gradient-to-r from-blue-600 to-indigo-600 hover:from-blue-700 hover:to-indigo-700"
                  onClick={handleProcessDocument} 
                  disabled={isProcessing || !file}
                >
                  Process Document
                </Button>
                <Button 
                  className="flex-1" 
                  variant="outline"
                  onClick={handleDownload}
                  disabled={isProcessing || tasks.length === 0}
                >
                  Download Results
                </Button>
                <Button
                  className="flex-1"
                  variant="secondary"
                  onClick={resetForm}
                  disabled={isProcessing}
                >
                  Reset
                </Button>
              </div>
            </CardContent>
          </Card>

          <div className="space-y-4">
            {isProcessing ? (
              <ProcessingIndicator status="parsing" progress={50} />
            ) : tasks.length > 0 ? (
              <TaskPreview tasks={tasks} documentTitle={docTitle || assemblyName} />
            ) : (
              <Card className="h-full">
                <CardHeader className="bg-gradient-to-r from-gray-200 via-gray-100 to-gray-200 text-gray-600">
                  <CardTitle>Upload and Process a Document</CardTitle>
                </CardHeader>
                <CardContent className="flex items-center justify-center h-64 text-center text-gray-500">
                  <p>
                    Upload a Word document (.docx) and click "Process Document" to extract tasks and images.
                    <br /><br />
                    The results will appear here.
                  </p>
                </CardContent>
              </Card>
            )}
          </div>
        </div>
      </main>

      <Footer />
    </div>
  );
};

export default Index;
