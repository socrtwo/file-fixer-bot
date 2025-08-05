import { Button } from "@/components/ui/button";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { Badge } from "@/components/ui/badge";
import { Alert, AlertDescription } from "@/components/ui/alert";
import { Separator } from "@/components/ui/separator";
import { 
  CheckCircle, 
  AlertTriangle, 
  XCircle, 
  Download, 
  RefreshCw,
  FileText,
  AlertCircle,
  FileCheck,
  Eye,
  Info
} from "lucide-react";

interface RepairResult {
  success: boolean;
  fileName: string;
  status: 'success' | 'partial' | 'failed';
  issues?: string[];
  repairedFile?: string; // base64 string from edge function
  repairedFileBlob?: Blob; // converted blob for download
  repairedFileV2?: Blob;
  repairedFileUrl?: string;
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

interface RepairResultsProps {
  result: RepairResult;
  onReset: () => void;
}

export const RepairResults = ({ result, onReset }: RepairResultsProps) => {
  const formatFileSize = (bytes: number): string => {
    if (bytes === 0) return '0 Bytes';
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
  };

  const downloadRepairedFile = () => {
    if (result.downloadUrl || result.repairedFileUrl) {
      window.open(result.downloadUrl || result.repairedFileUrl, '_blank');
    } else if (result.repairedFileBlob) {
      const url = URL.createObjectURL(result.repairedFileBlob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `repaired_${result.fileName}`;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
    } else if (result.repairedFileV2) {
      const url = URL.createObjectURL(result.repairedFileV2);
      const a = document.createElement('a');
      a.href = url;
      a.download = `repaired_${result.fileName}`;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
    } else if (result.repairedFile) {
      // Handle base64 string - convert to blob first
      try {
        const binaryString = atob(result.repairedFile);
        const bytes = new Uint8Array(binaryString.length);
        for (let i = 0; i < binaryString.length; i++) {
          bytes[i] = binaryString.charCodeAt(i);
        }
        const blob = new Blob([bytes], { type: 'application/octet-stream' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `repaired_${result.fileName}`;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
      } catch (error) {
        console.error('Error downloading file:', error);
        // Fallback: treat as text
        const blob = new Blob([result.repairedFile], { type: 'text/plain' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `repaired_${result.fileName}`;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
      }
    }
  };

  const getStatusIcon = (status: string) => {
    switch (status) {
      case 'success':
        return <CheckCircle className="h-12 w-12 text-green-500" />;
      case 'partial':
        return <AlertTriangle className="h-12 w-12 text-yellow-500" />;
      case 'failed':
        return <XCircle className="h-12 w-12 text-red-500" />;
      default:
        return <AlertCircle className="h-12 w-12 text-gray-500" />;
    }
  };

  const getStatusBadge = (status: string) => {
    switch (status) {
      case 'success':
        return <Badge className="bg-green-100 text-green-800">Success</Badge>;
      case 'partial':
        return <Badge className="bg-yellow-100 text-yellow-800">Partial</Badge>;
      case 'failed':
        return <Badge variant="destructive">Failed</Badge>;
      default:
        return <Badge variant="secondary">Unknown</Badge>;
    }
  };

  const getStatusMessage = (status: string) => {
    switch (status) {
      case 'success':
        return 'File repair completed successfully!';
      case 'partial':
        return 'File was partially repaired with some issues.';
      case 'failed':
        return 'File repair failed - too corrupted to recover.';
      default:
        return 'Repair status unknown.';
    }
  };

  return (
    <div className="w-full max-w-4xl mx-auto space-y-6">
      <Card className="shadow-soft border-0 bg-gradient-card">
        <CardHeader>
          <CardTitle className="flex items-center justify-between">
            <span>Repair Results</span>
            {getStatusBadge(result.status)}
          </CardTitle>
        </CardHeader>
        <CardContent className="space-y-6">
          <div className="space-y-4">
            <div className="text-center">
              {getStatusIcon(result.status)}
              <h3 className="text-xl font-semibold mt-4">{getStatusMessage(result.status)}</h3>
              <p className="text-muted-foreground mt-2">
                {result.fileName} - {formatFileSize(result.preview?.fileSize || result.repairedFileBlob?.size || result.repairedFileV2?.size || result.recoveryStats?.repairedSize || 0)}
              </p>
              {result.fileType && (
                <Badge variant="outline" className="mt-2">
                  {result.fileType} File
                </Badge>
              )}
            </div>

            <div className="flex gap-3 justify-center">
              <Button onClick={downloadRepairedFile} className="flex items-center gap-2">
                <Download className="h-4 w-4" />
                Download Repaired File
              </Button>
              <Button variant="outline" onClick={onReset} className="flex items-center gap-2">
                <RefreshCw className="h-4 w-4" />
                Try Another File
              </Button>
            </div>
          </div>

          {/* Recovery Statistics */}
          {result.recoveryStats && (
            <div className="bg-muted/50 p-4 rounded-lg">
              <h4 className="font-semibold mb-3 flex items-center gap-2">
                <FileCheck className="h-4 w-4" />
                Recovery Statistics
              </h4>
              <div className="grid grid-cols-3 gap-4 text-center">
                <div>
                  <div className="text-2xl font-bold text-primary">{result.recoveryStats.totalFiles}</div>
                  <div className="text-sm text-muted-foreground">Total Files</div>
                </div>
                <div>
                  <div className="text-2xl font-bold text-green-600">{result.recoveryStats.recoveredFiles}</div>
                  <div className="text-sm text-muted-foreground">Recovered</div>
                </div>
                <div>
                  <div className="text-2xl font-bold text-destructive">{result.recoveryStats.corruptedFiles}</div>
                  <div className="text-sm text-muted-foreground">Corrupted</div>
                </div>
              </div>
            </div>
          )}

          {/* Content Preview */}
          {result.preview && (
            <div className="bg-muted/50 p-4 rounded-lg">
              <h4 className="font-semibold mb-3 flex items-center gap-2">
                <Eye className="h-4 w-4" />
                Content Preview
              </h4>
              
              {result.fileType === 'DOCX' && result.preview.content && (
                <div className="bg-background p-3 rounded border text-sm font-mono">
                  {result.preview.content}
                </div>
              )}
              
              {result.fileType === 'XLSX' && result.preview.extractedSheets && (
                <div className="space-y-2">
                  <p className="text-sm text-muted-foreground">
                    Recovered {result.preview.extractedSheets.length} worksheets
                  </p>
                  <div className="flex flex-wrap gap-2">
                    {result.preview.extractedSheets.map((sheet, index) => (
                      <Badge key={index} variant="secondary">
                        {sheet}
                      </Badge>
                    ))}
                  </div>
                </div>
              )}
              
              {result.fileType === 'PPTX' && result.preview.extractedSlides !== undefined && (
                <div className="text-center">
                  <div className="text-2xl font-bold text-primary">{result.preview.extractedSlides}</div>
                  <div className="text-sm text-muted-foreground">Slides Recovered</div>
                </div>
              )}
            </div>
          )}

          {/* Issues */}
          {result.issues && result.issues.length > 0 && (
            <div className="bg-muted/50 p-4 rounded-lg">
              <h4 className="font-semibold mb-3 flex items-center gap-2">
                <AlertTriangle className="h-4 w-4" />
                Issues Detected & Repairs Made
              </h4>
              <div className="space-y-2">
                {result.issues.map((issue, index) => (
                  <div key={index} className="flex items-start gap-2 text-sm">
                    <AlertCircle className="h-4 w-4 text-yellow-500 mt-0.5 flex-shrink-0" />
                    <span>{issue}</span>
                  </div>
                ))}
              </div>
            </div>
          )}

          {/* Recommendations */}
          <div className="bg-muted/50 p-4 rounded-lg">
            <h4 className="font-semibold mb-3 flex items-center gap-2">
              <Info className="h-4 w-4" />
              Recommendations
            </h4>
            <div className="space-y-2 text-sm text-muted-foreground">
              {result.status === 'success' && (
                <p>• Your file has been successfully repaired and should work normally.</p>
              )}
              {result.status === 'partial' && (
                <>
                  <p>• Test the repaired file thoroughly before using it for important work.</p>
                  <p>• Some formatting or content may be lost due to the original corruption.</p>
                  <p>• Consider recovering from a recent backup if available.</p>
                </>
              )}
              {result.status === 'failed' && (
                <>
                  <p>• Try using Microsoft Office's built-in repair feature.</p>
                  <p>• Restore from a recent backup if available.</p>
                  <p>• The file may be too severely corrupted for automated repair.</p>
                </>
              )}
              <Separator className="my-2" />
              <p>• Always keep regular backups of important documents.</p>
              <p>• Consider using cloud storage with version history for critical files.</p>
            </div>
          </div>
        </CardContent>
      </Card>
    </div>
  );
};