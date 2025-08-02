import { Card, CardContent, CardHeader, CardTitle } from '@/components/ui/card';
import { Button } from '@/components/ui/button';
import { Badge } from '@/components/ui/badge';
import { Alert, AlertDescription } from '@/components/ui/alert';
import { Separator } from '@/components/ui/separator';
import { 
  CheckCircle, 
  AlertTriangle, 
  XCircle, 
  Download, 
  FileText, 
  RotateCcw,
  Info
} from 'lucide-react';

interface RepairResult {
  success: boolean;
  fileName: string;
  fileType: string;
  originalSize: number;
  repairedSize?: number;
  issues?: string[];
  repairedFile?: Blob;
  status: 'success' | 'partial' | 'failed';
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
    if (!result.repairedFile) return;
    
    const url = URL.createObjectURL(result.repairedFile);
    const a = document.createElement('a');
    a.href = url;
    a.download = `repaired_${result.fileName}`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };

  const getStatusIcon = () => {
    switch (result.status) {
      case 'success':
        return <CheckCircle className="w-6 h-6 text-success" />;
      case 'partial':
        return <AlertTriangle className="w-6 h-6 text-warning" />;
      case 'failed':
        return <XCircle className="w-6 h-6 text-destructive" />;
    }
  };

  const getStatusBadge = () => {
    switch (result.status) {
      case 'success':
        return <Badge className="bg-success text-success-foreground">Fully Repaired</Badge>;
      case 'partial':
        return <Badge className="bg-warning text-warning-foreground">Partially Repaired</Badge>;
      case 'failed':
        return <Badge variant="destructive">Repair Failed</Badge>;
    }
  };

  const getStatusMessage = () => {
    switch (result.status) {
      case 'success':
        return 'Your file has been successfully analyzed and appears to be in good condition.';
      case 'partial':
        return 'Your file has been partially repaired. Some issues were found and fixed, but the file may still have limitations.';
      case 'failed':
        return 'Unfortunately, your file is too corrupted to be repaired. The damage is too extensive.';
    }
  };

  return (
    <div className="w-full max-w-4xl mx-auto space-y-6">
      {/* Status Card */}
      <Card className="shadow-soft border-0 bg-gradient-card">
        <CardHeader className="pb-4">
          <div className="flex items-center justify-between">
            <CardTitle className="flex items-center gap-3">
              {getStatusIcon()}
              Repair Results
            </CardTitle>
            {getStatusBadge()}
          </div>
        </CardHeader>
        <CardContent className="space-y-4">
          <Alert className={`border-0 ${
            result.status === 'success' ? 'bg-success-light text-success-foreground' :
            result.status === 'partial' ? 'bg-warning-light text-warning-foreground' :
            'bg-destructive/10 text-destructive-foreground'
          }`}>
            <Info className="h-4 w-4" />
            <AlertDescription>
              {getStatusMessage()}
            </AlertDescription>
          </Alert>

          {/* File Information */}
          <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
            <div className="space-y-2">
              <h4 className="font-medium text-sm text-muted-foreground">File Details</h4>
              <div className="space-y-1">
                <p className="text-sm"><span className="font-medium">Name:</span> {result.fileName}</p>
                <p className="text-sm"><span className="font-medium">Type:</span> {result.fileType.toUpperCase()}</p>
                <p className="text-sm"><span className="font-medium">Original Size:</span> {formatFileSize(result.originalSize)}</p>
                {result.repairedSize && (
                  <p className="text-sm">
                    <span className="font-medium">Repaired Size:</span> {formatFileSize(result.repairedSize)}
                    {result.repairedSize !== result.originalSize && (
                      <span className={`ml-2 text-xs ${
                        result.repairedSize > result.originalSize ? 'text-warning' : 'text-success'
                      }`}>
                        ({result.repairedSize > result.originalSize ? '+' : ''}
                        {formatFileSize(result.repairedSize - result.originalSize)})
                      </span>
                    )}
                  </p>
                )}
              </div>
            </div>

            {/* Action Buttons */}
            <div className="space-y-2">
              <h4 className="font-medium text-sm text-muted-foreground">Actions</h4>
              <div className="flex flex-col gap-2">
                {result.success && result.repairedFile && (
                  <Button 
                    onClick={downloadRepairedFile}
                    className="bg-gradient-primary hover:shadow-medium transition-all duration-300"
                  >
                    <Download className="w-4 h-4 mr-2" />
                    Download Repaired File
                  </Button>
                )}
                <Button 
                  variant="outline" 
                  onClick={onReset}
                  className="hover:shadow-soft transition-all duration-300"
                >
                  <RotateCcw className="w-4 h-4 mr-2" />
                  Try Another File
                </Button>
              </div>
            </div>
          </div>
        </CardContent>
      </Card>

      {/* Issues Card */}
      {result.issues && result.issues.length > 0 && (
        <Card className="shadow-soft border-0 bg-gradient-card">
          <CardHeader>
            <CardTitle className="flex items-center gap-2">
              <FileText className="w-5 h-5" />
              Issues Detected & Repairs Made
            </CardTitle>
          </CardHeader>
          <CardContent>
            <div className="space-y-3">
              {result.issues.map((issue, index) => (
                <div key={index} className="flex items-start gap-3 p-3 rounded-lg bg-muted/50">
                  <AlertTriangle className="w-4 h-4 text-warning mt-0.5 flex-shrink-0" />
                  <p className="text-sm">{issue}</p>
                </div>
              ))}
            </div>
          </CardContent>
        </Card>
      )}

      {/* Recommendations Card */}
      <Card className="shadow-soft border-0 bg-gradient-card">
        <CardHeader>
          <CardTitle className="flex items-center gap-2">
            <Info className="w-5 h-5" />
            Recommendations
          </CardTitle>
        </CardHeader>
        <CardContent className="space-y-3">
          {result.status === 'success' && (
            <div className="p-3 rounded-lg bg-success-light">
              <p className="text-sm text-success-foreground">
                • Your file appears to be in good condition. You can use it normally.
              </p>
            </div>
          )}
          
          {result.status === 'partial' && (
            <div className="space-y-2">
              <div className="p-3 rounded-lg bg-warning-light">
                <p className="text-sm text-warning-foreground">
                  • Test the repaired file thoroughly before using it for important work.
                </p>
              </div>
              <div className="p-3 rounded-lg bg-warning-light">
                <p className="text-sm text-warning-foreground">
                  • Some formatting or content may be lost due to corruption.
                </p>
              </div>
              <div className="p-3 rounded-lg bg-warning-light">
                <p className="text-sm text-warning-foreground">
                  • Consider recovering from a recent backup if available.
                </p>
              </div>
            </div>
          )}
          
          {result.status === 'failed' && (
            <div className="space-y-2">
              <div className="p-3 rounded-lg bg-destructive/10">
                <p className="text-sm text-destructive-foreground">
                  • Try using Microsoft Office's built-in repair feature.
                </p>
              </div>
              <div className="p-3 rounded-lg bg-destructive/10">
                <p className="text-sm text-destructive-foreground">
                  • Restore from a recent backup if available.
                </p>
              </div>
              <div className="p-3 rounded-lg bg-destructive/10">
                <p className="text-sm text-destructive-foreground">
                  • Contact IT support if this is a critical business document.
                </p>
              </div>
            </div>
          )}
          
          <Separator />
          
          <div className="text-xs text-muted-foreground space-y-1">
            <p>• Always keep regular backups of important documents.</p>
            <p>• Scan your system for malware if file corruption happens frequently.</p>
            <p>• Consider using cloud storage with version history for critical files.</p>
          </div>
        </CardContent>
      </Card>
    </div>
  );
};