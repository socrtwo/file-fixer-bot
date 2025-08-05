import { useState, useRef } from 'react';
import { Button } from '@/components/ui/button';
import { Card, CardContent } from '@/components/ui/card';
import { Progress } from '@/components/ui/progress';
import { Badge } from '@/components/ui/badge';
import { useToast } from '@/hooks/use-toast';
import { Upload, FileText, Table, Presentation, CheckCircle, AlertCircle, Download } from 'lucide-react';
import JSZip from 'jszip';

interface FileUploadProps {
  onFileProcessed: (result: RepairResult) => void;
}

interface RepairResult {
  success: boolean;
  fileName: string;
  fileType: string;
  originalSize: number;
  repairedSize?: number;
  issues?: string[];
  repairedFile?: Blob;
  repairedFileV2?: Blob;
  status: 'success' | 'partial' | 'failed';
}

const ACCEPTED_TYPES = {
  'application/vnd.openxmlformats-officedocument.wordprocessingml.document': 'docx',
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': 'xlsx',
  'application/vnd.openxmlformats-officedocument.presentationml.presentation': 'pptx'
};

export const FileUpload = ({ onFileProcessed }: FileUploadProps) => {
  const [isDragActive, setIsDragActive] = useState(false);
  const [isProcessing, setIsProcessing] = useState(false);
  const [progress, setProgress] = useState(0);
  const fileInputRef = useRef<HTMLInputElement>(null);
  const { toast } = useToast();

  const getFileIcon = (type: string) => {
    switch (type) {
      case 'docx': return <FileText className="w-8 h-8 text-primary" />;
      case 'xlsx': return <Table className="w-8 h-8 text-accent" />;
      case 'pptx': return <Presentation className="w-8 h-8 text-warning" />;
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
    const fileType = ACCEPTED_TYPES[file.type as keyof typeof ACCEPTED_TYPES] || 'unknown';
    
    try {
      setProgress(15);
      
      // Read the file as array buffer
      const arrayBuffer = await file.arrayBuffer();
      setProgress(25);
      
      // Step 1: Attempt ZIP repair for all file types
      let repairedArrayBuffer = arrayBuffer;
      const issues: string[] = [];
      
      try {
        const zip = new JSZip();
        await zip.loadAsync(arrayBuffer);
        setProgress(35);
      } catch (error) {
        // ZIP is corrupted, attempt basic ZIP repair
        issues.push('ZIP structure was corrupted and repaired');
        repairedArrayBuffer = await repairZipStructure(arrayBuffer);
        setProgress(35);
      }
      
      // Try to load the (potentially repaired) ZIP
      const zip = new JSZip();
      let zipContent;
      
      try {
        zipContent = await zip.loadAsync(repairedArrayBuffer);
        setProgress(40);
      } catch (error) {
        // File is severely corrupted, try to extract what we can
        return {
          success: false,
          fileName: file.name,
          fileType,
          originalSize: file.size,
          issues: ['File is severely corrupted and cannot be opened as a ZIP archive even after repair attempts'],
          status: 'failed'
        };
      }

      let repairedFileV1: Blob | undefined;
      let repairedFileV2: Blob | undefined;

      // For Word documents, create Version 1 with document.xml repair
      if (fileType === 'docx') {
        setProgress(50);
        const v1Zip = new JSZip();
        
        // Copy all files and specifically repair document.xml
        for (const [path, zipFile] of Object.entries(zipContent.files)) {
          const file = zipFile as JSZip.JSZipObject;
          if (file.dir) {
            v1Zip.folder(path);
          } else if (path === 'word/document.xml') {
            try {
              const content = await file.async('string');
              const repairedContent = repairDocumentXML(content);
              v1Zip.file(path, repairedContent);
              if (content !== repairedContent) {
                issues.push('document.xml was truncated/repaired');
              }
            } catch (error) {
              // Create minimal document.xml if completely corrupted
              const minimalDoc = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r>
        <w:t>Document recovered. Original content was corrupted.</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>`;
              v1Zip.file(path, minimalDoc);
              issues.push('document.xml was completely replaced with minimal content');
            }
          } else {
            try {
              v1Zip.file(path, await file.async('arraybuffer'));
            } catch {
              // Skip corrupted files in V1
            }
          }
        }
        
        setProgress(60);
        
        // Generate Version 1
        const v1ArrayBuffer = await v1Zip.generateAsync({
          type: 'arraybuffer',
          compression: 'DEFLATE',
          compressionOptions: { level: 6 }
        });
        repairedFileV1 = new Blob([v1ArrayBuffer], { type: file.type });
      }

      setProgress(70);

      // Create Version 2 (original repair method)
      const repairedZip = new JSZip();
      const essentialFiles = getEssentialFiles(fileType);
      const missingFiles: string[] = [];
      
      // Validate and repair structure (original method)
      for (const [path, zipFile] of Object.entries(zipContent.files)) {
        try {
          const file = zipFile as JSZip.JSZipObject;
          if (file.dir) {
            repairedZip.folder(path);
          } else {
            const content = await file.async('string');
            
            // Basic XML validation for Office files
            if (path.endsWith('.xml') || path.endsWith('.rels')) {
              if (!isValidXML(content)) {
                issues.push(`Corrupted XML file: ${path}`);
                // Try to repair basic XML issues
                const repairedContent = repairXML(content);
                repairedZip.file(path, repairedContent);
              } else {
                repairedZip.file(path, content);
              }
            } else {
              repairedZip.file(path, await file.async('arraybuffer'));
            }
          }
        } catch (error) {
          issues.push(`Failed to process file: ${path}`);
        }
      }
      
      // Check for missing essential files
      essentialFiles.forEach(essentialFile => {
        if (!zipContent.files[essentialFile]) {
          missingFiles.push(essentialFile);
        }
      });
      
      if (missingFiles.length > 0) {
        issues.push(`Missing essential files: ${missingFiles.join(', ')}`);
        // Add minimal versions of missing files
        missingFiles.forEach(missingFile => {
          const minimalContent = generateMinimalContent(missingFile, fileType);
          if (minimalContent) {
            repairedZip.file(missingFile, minimalContent);
          }
        });
      }
      
      setProgress(90);
      
      // Generate Version 2
      const v2ArrayBuffer = await repairedZip.generateAsync({
        type: 'arraybuffer',
        compression: 'DEFLATE',
        compressionOptions: { level: 6 }
      });
      repairedFileV2 = new Blob([v2ArrayBuffer], { type: file.type });
      
      setProgress(100);
      
      const status = issues.length === 0 ? 'success' : 
                   missingFiles.length === 0 ? 'partial' : 'partial';
      
      return {
        success: true,
        fileName: file.name,
        fileType,
        originalSize: file.size,
        repairedSize: fileType === 'docx' ? repairedFileV1?.size : repairedFileV2?.size,
        issues: issues.length > 0 ? issues : undefined,
        repairedFile: fileType === 'docx' ? repairedFileV1 : repairedFileV2,
        repairedFileV2: fileType === 'docx' ? repairedFileV2 : undefined,
        status
      };
      
    } catch (error) {
      return {
        success: false,
        fileName: file.name,
        fileType,
        originalSize: file.size,
        issues: [`Repair failed: ${error instanceof Error ? error.message : 'Unknown error'}`],
        status: 'failed'
      };
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
        description: "Please upload a DOCX, XLSX, or PPTX file.",
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
                <h3 className="text-lg font-semibold">Upload Corrupted Office File</h3>
                <p className="text-sm text-muted-foreground">
                  Drop your DOCX, XLSX, or PPTX file here or click to browse
                </p>
              </div>
              
              <div className="flex flex-wrap justify-center gap-2">
                <Badge variant="outline">Microsoft Word</Badge>
                <Badge variant="outline">Microsoft Excel</Badge>
                <Badge variant="outline">Microsoft PowerPoint</Badge>
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
            accept=".docx,.xlsx,.pptx"
            onChange={(e) => handleFileSelect(e.target.files)}
          />
        </div>
      </CardContent>
    </Card>
  );
};