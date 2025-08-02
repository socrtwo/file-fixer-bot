import { useState } from 'react';
import { FileUpload } from '@/components/FileUpload';
import { RepairResults } from '@/components/RepairResults';
import { Card, CardContent } from '@/components/ui/card';
import { Badge } from '@/components/ui/badge';
import { Shield, Zap, FileText } from 'lucide-react';

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
              Repair corrupted Microsoft Word, Excel, and PowerPoint files with our advanced recovery technology
            </p>
            
            <div className="flex flex-wrap justify-center gap-3 mt-6">
              <div className="flex items-center gap-2 px-4 py-2 rounded-full bg-card border shadow-soft">
                <Shield className="w-4 h-4 text-primary" />
                <span className="text-sm font-medium">Secure Processing</span>
              </div>
              <div className="flex items-center gap-2 px-4 py-2 rounded-full bg-card border shadow-soft">
                <Zap className="w-4 h-4 text-accent" />
                <span className="text-sm font-medium">Fast Recovery</span>
              </div>
              <div className="flex items-center gap-2 px-4 py-2 rounded-full bg-card border shadow-soft">
                <FileText className="w-4 h-4 text-warning" />
                <span className="text-sm font-medium">All Office Formats</span>
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
            
            {/* Features Section */}
            <div className="grid grid-cols-1 md:grid-cols-3 gap-6 max-w-4xl mx-auto">
              <Card className="shadow-soft border-0 bg-gradient-card hover:shadow-medium transition-all duration-300">
                <CardContent className="p-6 text-center space-y-4">
                  <div className="p-3 rounded-full bg-primary/10 w-fit mx-auto">
                    <FileText className="w-6 h-6 text-primary" />
                  </div>
                  <h3 className="font-semibold">DOCX Files</h3>
                  <p className="text-sm text-muted-foreground">
                    Repair corrupted Microsoft Word documents and recover your important content.
                  </p>
                </CardContent>
              </Card>
              
              <Card className="shadow-soft border-0 bg-gradient-card hover:shadow-medium transition-all duration-300">
                <CardContent className="p-6 text-center space-y-4">
                  <div className="p-3 rounded-full bg-accent/10 w-fit mx-auto">
                    <FileText className="w-6 h-6 text-accent" />
                  </div>
                  <h3 className="font-semibold">XLSX Files</h3>
                  <p className="text-sm text-muted-foreground">
                    Restore damaged Excel spreadsheets and recover your valuable data.
                  </p>
                </CardContent>
              </Card>
              
              <Card className="shadow-soft border-0 bg-gradient-card hover:shadow-medium transition-all duration-300">
                <CardContent className="p-6 text-center space-y-4">
                  <div className="p-3 rounded-full bg-warning/10 w-fit mx-auto">
                    <FileText className="w-6 h-6 text-warning" />
                  </div>
                  <h3 className="font-semibold">PPTX Files</h3>
                  <p className="text-sm text-muted-foreground">
                    Fix broken PowerPoint presentations and save your slides.
                  </p>
                </CardContent>
              </Card>
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
              This tool processes files locally in your browser for maximum security and privacy.
            </p>
            <div className="flex flex-wrap justify-center gap-2">
              <Badge variant="outline" className="text-xs">Privacy First</Badge>
              <Badge variant="outline" className="text-xs">No File Upload</Badge>
              <Badge variant="outline" className="text-xs">Browser Based</Badge>
            </div>
          </div>
        </div>
      </footer>
    </div>
  );
};

export default Index;
