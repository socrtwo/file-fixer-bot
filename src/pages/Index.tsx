import { useState } from 'react';
import { FileUpload } from '@/components/FileUpload';
import { RepairResults } from '@/components/RepairResults';
import { Card, CardContent } from '@/components/ui/card';
import { Badge } from '@/components/ui/badge';
import { Shield, Zap, FileText, Layers, BarChart3, Presentation } from 'lucide-react';

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

const Index = () => {
  const [repairResult, setRepairResult] = useState<RepairResult | null>(null);

  const handleFileProcessed = (result: RepairResult) => {
    setRepairResult(result);
  };

  const handleReset = () => {
    setRepairResult(null);
  };

  return (
    <div className="min-h-screen bg-gradient-surface">
      {/* Header */}
      <div className="border-b bg-card/50 backdrop-blur-sm">
        <div className="container mx-auto px-4 py-6">
          <div className="text-center space-y-4">
            <div className="flex items-center justify-center gap-3">
              <div className="p-3 rounded-full bg-gradient-primary shadow-medium">
                <FileText className="w-8 h-8 text-primary-foreground" />
              </div>
              <h1 className="text-4xl font-bold bg-gradient-to-r from-primary to-primary-glow bg-clip-text text-transparent">
                Office File Repair Tool
              </h1>
            </div>
            <p className="text-lg text-muted-foreground max-w-2xl mx-auto">
              Advanced Office file repair with format-specific recovery, content preview, and detailed repair statistics
            </p>
            
            <div className="flex flex-wrap justify-center gap-3 mt-6">
              <div className="flex items-center gap-2 px-4 py-2 rounded-full bg-card border shadow-soft">
                <Shield className="w-4 h-4 text-primary" />
                <span className="text-sm font-medium">Secure Processing</span>
              </div>
              <div className="flex items-center gap-2 px-4 py-2 rounded-full bg-card border shadow-soft">
                <Zap className="w-4 h-4 text-accent" />
                <span className="text-sm font-medium">Smart Recovery</span>
              </div>
              <div className="flex items-center gap-2 px-4 py-2 rounded-full bg-card border shadow-soft">
                <Layers className="w-4 h-4 text-warning" />
                <span className="text-sm font-medium">Content Preview</span>
              </div>
            </div>
          </div>
        </div>
      </div>

      {/* Main Content */}
      <div className="container mx-auto px-4 py-12">
        {!repairResult ? (
          <div className="space-y-12">
            {/* Upload Section */}
            <FileUpload onFileProcessed={handleFileProcessed} />
            
            {/* Enhanced Features Section */}
            <div className="grid grid-cols-1 md:grid-cols-3 gap-6 max-w-4xl mx-auto">
              <Card className="shadow-soft border-0 bg-gradient-card hover:shadow-medium transition-all duration-300">
                <CardContent className="p-6 text-center space-y-4">
                  <div className="p-3 rounded-full bg-primary/10 w-fit mx-auto">
                    <FileText className="w-6 h-6 text-primary" />
                  </div>
                  <h3 className="font-semibold">DOCX Files</h3>
                  <p className="text-sm text-muted-foreground">
                    Advanced Word document repair with content extraction and text preview.
                  </p>
                  <div className="flex flex-wrap gap-1 justify-center">
                    <Badge variant="outline" className="text-xs">Text Recovery</Badge>
                    <Badge variant="outline" className="text-xs">XML Repair</Badge>
                  </div>
                </CardContent>
              </Card>
              
              <Card className="shadow-soft border-0 bg-gradient-card hover:shadow-medium transition-all duration-300">
                <CardContent className="p-6 text-center space-y-4">
                  <div className="p-3 rounded-full bg-accent/10 w-fit mx-auto">
                    <BarChart3 className="w-6 h-6 text-accent" />
                  </div>
                  <h3 className="font-semibold">XLSX Files</h3>
                  <p className="text-sm text-muted-foreground">
                    Excel spreadsheet recovery with worksheet-by-worksheet analysis.
                  </p>
                  <div className="flex flex-wrap gap-1 justify-center">
                    <Badge variant="outline" className="text-xs">Sheet Recovery</Badge>
                    <Badge variant="outline" className="text-xs">Data Validation</Badge>
                  </div>
                </CardContent>
              </Card>
              
              <Card className="shadow-soft border-0 bg-gradient-card hover:shadow-medium transition-all duration-300">
                <CardContent className="p-6 text-center space-y-4">
                  <div className="p-3 rounded-full bg-warning/10 w-fit mx-auto">
                    <Presentation className="w-6 h-6 text-warning" />
                  </div>
                  <h3 className="font-semibold">PPTX Files</h3>
                  <p className="text-sm text-muted-foreground">
                    PowerPoint presentation repair with slide counting and structure recovery.
                  </p>
                  <div className="flex flex-wrap gap-1 justify-center">
                    <Badge variant="outline" className="text-xs">Slide Recovery</Badge>
                    <Badge variant="outline" className="text-xs">Structure Repair</Badge>
                  </div>
                </CardContent>
              </Card>
            </div>

            {/* New Features Showcase */}
            <div className="max-w-4xl mx-auto">
              <h2 className="text-2xl font-bold text-center mb-8">Enhanced Recovery Features</h2>
              <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                <Card className="shadow-soft border-0 bg-gradient-card">
                  <CardContent className="p-6 space-y-4">
                    <div className="flex items-center gap-3">
                      <div className="p-2 rounded-full bg-primary/10">
                        <Layers className="w-5 h-5 text-primary" />
                      </div>
                      <h3 className="font-semibold">Content Preview</h3>
                    </div>
                    <p className="text-sm text-muted-foreground">
                      See what was recovered before downloading. Preview document text, worksheet names, or slide counts.
                    </p>
                  </CardContent>
                </Card>
                
                <Card className="shadow-soft border-0 bg-gradient-card">
                  <CardContent className="p-6 space-y-4">
                    <div className="flex items-center gap-3">
                      <div className="p-2 rounded-full bg-accent/10">
                        <BarChart3 className="w-5 h-5 text-accent" />
                      </div>
                      <h3 className="font-semibold">Recovery Statistics</h3>
                    </div>
                    <p className="text-sm text-muted-foreground">
                      Detailed breakdown of files recovered vs. corrupted with comprehensive repair reports.
                    </p>
                  </CardContent>
                </Card>
              </div>
            </div>
          </div>
        ) : (
          <RepairResults result={repairResult} onReset={handleReset} />
        )}
      </div>
      
      {/* Footer */}
      <footer className="border-t bg-card/30 backdrop-blur-sm">
        <div className="container mx-auto px-4 py-8">
          <div className="text-center space-y-4">
            <p className="text-sm text-muted-foreground">
              Enhanced Office file repair with format-specific algorithms and intelligent content recovery.
            </p>
            <div className="flex flex-wrap justify-center gap-2">
              <Badge variant="outline" className="text-xs">Format-Specific</Badge>
              <Badge variant="outline" className="text-xs">Content Preview</Badge>
              <Badge variant="outline" className="text-xs">Recovery Stats</Badge>
              <Badge variant="outline" className="text-xs">Privacy First</Badge>
            </div>
          </div>
        </div>
      </footer>
    </div>
  );
};

export default Index;