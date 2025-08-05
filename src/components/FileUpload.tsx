import { useState, useRef } from 'react';
import { Button } from '@/components/ui/button';
import { Card, CardContent } from '@/components/ui/card';
import { Progress } from '@/components/ui/progress';
import { Badge } from '@/components/ui/badge';
import { useToast } from '@/hooks/use-toast';
import { Upload, FileText, Table, Presentation, CheckCircle, AlertCircle, Download } from 'lucide-react';
import { supabase } from '@/integrations/supabase/client';

interface FileUploadProps {
  onFileProcessed: (result: RepairResult) => void;
}

interface RepairResult {
  success: boolean;
  fileName: string;
  status: 'success' | 'partial' | 'failed';
  issues?: string[];
  repairedFile?: string; // base64 string from edge function
  repairedFileBlob?: Blob; // converted blob for download
  downloadUrl?: string;
  preview?: {
    content?: string;
    extractedSheets?: string[];
    extractedSlides?: number;
    recoveredFiles?: string[];
    fileSize?: number;
  };
  fileType?: 'DOCX' | 'XLSX' | 'PPTX' | 'ZIP' | 'PDF' | 'txt';
  recoveryStats?: {
    totalFiles: number;
    recoveredFiles: number;
    corruptedFiles: number;
    originalSize?: number;
    repairedSize?: number;
    corruptionLevel?: string;
    recoveredData?: number;
  };
}

const ACCEPTED_TYPES = {
  'application/vnd.openxmlformats-officedocument.wordprocessingml.document': 'DOCX',
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': 'XLSX',
  'application/vnd.openxmlformats-officedocument.presentationml.presentation': 'PPTX',
  'application/zip': 'ZIP',
  'application/x-zip-compressed': 'ZIP',
  'application/pdf': 'PDF'
};

export const FileUpload = ({ onFileProcessed }: FileUploadProps) => {
  const [isDragActive, setIsDragActive] = useState(false);
  const [isProcessing, setIsProcessing] = useState(false);
  const [progress, setProgress] = useState(0);
  const fileInputRef = useRef<HTMLInputElement>(null);
  const { toast } = useToast();

  const getFileIcon = (type: string) => {
    switch (type) {
      case 'DOCX': return <FileText className="w-8 h-8 text-primary" />;
      case 'XLSX': return <Table className="w-8 h-8 text-accent" />;
      case 'PPTX': return <Presentation className="w-8 h-8 text-warning" />;
      default: return <FileText className="w-8 h-8 text-muted-foreground" />;
    }
  };

  const repairZipStructure = async (arrayBuffer: ArrayBuffer): Promise<ArrayBuffer> => {
    // Basic ZIP repair - try to find ZIP header and reconstruct
    const uint8Array = new Uint8Array(arrayBuffer);
    
    // Look for ZIP file signature (PK header)
    const pkHeader = [0x50, 0x4B, 0x03, 0x04];
    let headerIndex = -1;
    
    for (let i = 0; i <= uint8Array.length - 4; i++) {
      if (uint8Array[i] === pkHeader[0] && 
          uint8Array[i + 1] === pkHeader[1] && 
          uint8Array[i + 2] === pkHeader[2] && 
          uint8Array[i + 3] === pkHeader[3]) {
        headerIndex = i;
        break;
      }
    }
    
    if (headerIndex > 0) {
      // Found header, try to trim any garbage before it
      return arrayBuffer.slice(headerIndex);
    }
    
    // If no valid header found, return original
    return arrayBuffer;
  };

  const repairDocumentXML = (content: string): string => {
    try {
      // Parse the XML to find content
      const parser = new DOMParser();
      const doc = parser.parseFromString(content, 'text/xml');
      
      // If parsing fails, try to truncate to last valid paragraph
      if (doc.querySelector('parsererror')) {
        // Find the last complete paragraph tag
        const lastParagraphMatch = content.lastIndexOf('</w:p>');
        if (lastParagraphMatch !== -1) {
          const truncatedContent = content.substring(0, lastParagraphMatch + 6);
          // Add closing tags if needed
          let repaired = truncatedContent;
          if (!repaired.includes('</w:body>')) {
            repaired += '</w:body>';
          }
          if (!repaired.includes('</w:document>')) {
            repaired += '</w:document>';
          }
          return repaired;
        }
      }
      
      return content;
    } catch {
      // Fallback: create minimal document.xml
      return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r>
        <w:t>Document recovered with minimal content.</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>`;
    }
  };

  const validateAndRepairFile = async (file: File): Promise<RepairResult> => {
    setProgress(10);
    
    try {
      setProgress(20);
      
      // Create FormData to send file to backend
      const formData = new FormData();
      formData.append('file', file);
      
      setProgress(40);
      
      // Call the Edge Function
      console.log('Calling edge function with file:', file.name);
      const { data, error } = await supabase.functions.invoke('repair-office-file', {
        body: formData,
      });

      setProgress(80);
      console.log('Edge function response:', { data, error });

      if (error) {
        console.error('Edge function error:', error);
        throw new Error(`Failed to send a request to the Edge Function`);
      }

      if (!data) {
        console.error('No data returned from edge function');
        throw new Error('No data returned from edge function');
      }

      setProgress(100);
      
      // Log the response structure
      console.log('Edge function returned data:', {
        success: data.success,
        fileName: data.fileName,
        repairedFileLength: data.repairedFile?.length || 0,
        hasRepairedFile: !!data.repairedFile
      });

      // The edge function returns base64 content in repairedFile
      if (data.repairedFile) {
        // Convert base64 to blob for download
        try {
          const binaryString = atob(data.repairedFile);
          const bytes = new Uint8Array(binaryString.length);
          for (let i = 0; i < binaryString.length; i++) {
            bytes[i] = binaryString.charCodeAt(i);
          }
          const blob = new Blob([bytes], { type: 'application/octet-stream' });
          data.repairedFileBlob = blob;
          console.log('Created blob with size:', blob.size);
        } catch (decodeError) {
          console.error('Error decoding base64:', decodeError);
          // Fallback: treat as plain text
          const blob = new Blob([data.repairedFile], { type: 'text/plain' });
          data.repairedFileBlob = blob;
        }
      }

      return data as RepairResult;

    } catch (error) {
      console.error('File processing error:', error);
      throw new Error(`Failed to send a request to the Edge Function`);
    }
  };

  const getEssentialFiles = (fileType: string): string[] => {
    switch (fileType) {
      case 'docx':
        return [
          '[Content_Types].xml',
          '_rels/.rels',
          'word/document.xml',
          'word/_rels/document.xml.rels'
        ];
      case 'xlsx':
        return [
          '[Content_Types].xml',
          '_rels/.rels',
          'xl/workbook.xml',
          'xl/_rels/workbook.xml.rels',
          'xl/worksheets/sheet1.xml'
        ];
      case 'pptx':
        return [
          '[Content_Types].xml',
          '_rels/.rels',
          'ppt/presentation.xml',
          'ppt/_rels/presentation.xml.rels'
        ];
      default:
        return [];
    }
  };

  const isValidXML = (content: string): boolean => {
    try {
      const parser = new DOMParser();
      const doc = parser.parseFromString(content, 'text/xml');
      const error = doc.querySelector('parsererror');
      return !error;
    } catch {
      return false;
    }
  };

  const repairXML = (content: string): string => {
    // Basic XML repair techniques
    let repaired = content;
    
    // Fix unclosed tags (basic approach)
    repaired = repaired.replace(/&(?!(?:amp|lt|gt|quot|apos);)/g, '&amp;');
    
    // Ensure XML declaration
    if (!repaired.startsWith('<?xml')) {
      repaired = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + repaired;
    }
    
    return repaired;
  };

  const generateMinimalContent = (filePath: string, fileType: string): string | null => {
    // Generate minimal valid content for missing essential files
    if (filePath === '[Content_Types].xml') {
      return getMinimalContentTypes(fileType);
    }
    if (filePath === '_rels/.rels') {
      return getMinimalRels(fileType);
    }
    // Add more minimal content generators as needed
    return null;
  };

  const getMinimalContentTypes = (fileType: string): string => {
    const base = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>`;
    
    switch (fileType) {
      case 'docx':
        return base + `
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`;
      case 'xlsx':
        return base + `
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
</Types>`;
      case 'pptx':
        return base + `
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
</Types>`;
      default:
        return base + '\n</Types>';
    }
  };

  const getMinimalRels = (fileType: string): string => {
    const base = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`;
    
    switch (fileType) {
      case 'docx':
        return base + `
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;
      case 'xlsx':
        return base + `
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`;
      case 'pptx':
        return base + `
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
</Relationships>`;
      default:
        return base + '\n</Relationships>';
    }
  };

  const handleFileSelect = async (files: FileList | null) => {
    if (!files || files.length === 0) return;
    
    const file = files[0];
    
    // Validate file type
    if (!Object.keys(ACCEPTED_TYPES).includes(file.type)) {
      toast({
        title: "Invalid file type",
        description: "Please upload a DOCX, XLSX, PPTX, ZIP, or PDF file.",
        variant: "destructive",
      });
      return;
    }

    setIsProcessing(true);
    setProgress(0);
    
    try {
      const result = await validateAndRepairFile(file);
      onFileProcessed(result);
      
      if (result.success) {
        toast({
          title: "File processed successfully",
          description: result.status === 'success' ? 
            "File appears to be healthy or has been repaired." :
            "File has been partially repaired. Check the details below.",
        });
      } else {
        toast({
          title: "Repair failed",
          description: "The file is too corrupted to repair.",
          variant: "destructive",
        });
      }
    } catch (error) {
      toast({
        title: "Processing error",
        description: "An unexpected error occurred while processing the file.",
        variant: "destructive",
      });
    } finally {
      setIsProcessing(false);
      setProgress(0);
    }
  };

  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragActive(false);
    handleFileSelect(e.dataTransfer.files);
  };

  const handleDragOver = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragActive(true);
  };

  const handleDragLeave = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragActive(false);
  };

  return (
    <Card className="w-full max-w-2xl mx-auto shadow-soft border-0 bg-gradient-card">
      <CardContent className="p-8">
        <div
          className={`relative border-2 border-dashed rounded-lg p-8 text-center transition-all duration-300 ${
            isDragActive 
              ? 'border-primary bg-primary/5 scale-105' 
              : 'border-border hover:border-primary/50 hover:bg-primary/2'
          }`}
          onDrop={handleDrop}
          onDragOver={handleDragOver}
          onDragLeave={handleDragLeave}
        >
          {isProcessing ? (
            <div className="space-y-4">
              <div className="animate-pulse-soft">
                <Upload className="w-12 h-12 text-primary mx-auto" />
              </div>
              <div className="space-y-2">
                <p className="text-sm font-medium">Processing file...</p>
                <Progress value={progress} className="w-full" />
                <p className="text-xs text-muted-foreground">{progress}% complete</p>
              </div>
            </div>
          ) : (
            <div className="space-y-6">
              <div className="flex justify-center space-x-4">
                {getFileIcon('docx')}
                {getFileIcon('xlsx')}
                {getFileIcon('pptx')}
              </div>
              
              <div className="space-y-2">
                <h3 className="text-lg font-semibold">Upload Corrupted File</h3>
                <p className="text-sm text-muted-foreground">
                  Drop your DOCX, XLSX, PPTX, ZIP, or PDF file here or click to browse
                </p>
              </div>
              
              <div className="flex flex-wrap justify-center gap-2">
                <Badge variant="outline">Microsoft Word</Badge>
                <Badge variant="outline">Microsoft Excel</Badge>
                <Badge variant="outline">Microsoft PowerPoint</Badge>
                <Badge variant="outline">ZIP Archives</Badge>
                <Badge variant="outline">PDF Documents</Badge>
              </div>
              
              <Button
                onClick={() => fileInputRef.current?.click()}
                className="bg-gradient-primary hover:shadow-medium transition-all duration-300"
              >
                <Upload className="w-4 h-4 mr-2" />
                Choose File
              </Button>
            </div>
          )}
          
          <input
            ref={fileInputRef}
            type="file"
            className="hidden"
            accept=".docx,.xlsx,.pptx,.zip,.pdf"
            onChange={(e) => handleFileSelect(e.target.files)}
          />
        </div>
      </CardContent>
    </Card>
  );
};