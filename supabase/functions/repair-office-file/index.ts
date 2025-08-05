import { serve } from "https://deno.land/std@0.168.0/http/server.ts";
import { createClient } from 'https://esm.sh/@supabase/supabase-js@2';

const corsHeaders = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Headers': 'authorization, x-client-info, apikey, content-type',
};

interface RepairResult {
  success: boolean;
  fileName: string;
  fileType: string;
  originalSize: number;
  repairedSize?: number;
  issues?: string[];
  repairedFileUrl?: string;
  status: 'success' | 'partial' | 'failed';
}

serve(async (req) => {
  // Handle CORS preflight requests
  if (req.method === 'OPTIONS') {
    return new Response(null, { headers: corsHeaders });
  }

  try {
    console.log('Starting file repair process');
    
    const formData = await req.formData();
    const file = formData.get('file') as File;
    
    if (!file) {
      return new Response(JSON.stringify({ error: 'No file provided' }), {
        status: 400,
        headers: { ...corsHeaders, 'Content-Type': 'application/json' },
      });
    }

    console.log(`Processing file: ${file.name}, size: ${file.size} bytes`);

    // Validate file type
    const allowedTypes = [
      'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      'application/vnd.openxmlformats-officedocument.presentationml.presentation'
    ];

    if (!allowedTypes.includes(file.type)) {
      return new Response(JSON.stringify({ 
        error: 'Unsupported file type. Only DOCX, XLSX, and PPTX files are supported.' 
      }), {
        status: 400,
        headers: { ...corsHeaders, 'Content-Type': 'application/json' },
      });
    }

    // Create temporary directory for extraction
    const tempDir = await Deno.makeTempDir({ prefix: 'office_repair_' });
    const inputFile = `${tempDir}/input_file`;
    const extractDir = `${tempDir}/extracted`;
    const outputFile = `${tempDir}/repaired_file`;

    try {
      // Write uploaded file to temporary location
      const fileBytes = new Uint8Array(await file.arrayBuffer());
      await Deno.writeFile(inputFile, fileBytes);

      // Create extraction directory
      await Deno.mkdir(extractDir, { recursive: true });

      console.log('Extracting file with p7zip...');
      
      // Extract using p7zip - more robust than standard unzip for corrupted files
      const extractProcess = new Deno.Command("7z", {
        args: ["x", "-y", `-o${extractDir}`, inputFile],
        stdout: "piped",
        stderr: "piped"
      });

      const extractResult = await extractProcess.output();
      
      if (!extractResult.success) {
        console.warn('Standard extraction failed, trying recovery mode...');
        
        // Try with recovery mode for heavily corrupted files
        const recoveryProcess = new Deno.Command("7z", {
          args: ["x", "-y", "-r", `-o${extractDir}`, inputFile],
          stdout: "piped",
          stderr: "piped"
        });
        
        const recoveryResult = await recoveryProcess.output();
        if (!recoveryResult.success) {
          throw new Error('Failed to extract file even with recovery mode');
        }
      }

      console.log('File extracted successfully');

      // Repair XML files in the extracted directory
      const issues = await repairXmlFiles(extractDir);
      
      console.log('Creating repaired archive...');
      
      // Create new archive with repaired files
      const compressProcess = new Deno.Command("7z", {
        args: ["a", "-tzip", outputFile, `${extractDir}/*`],
        stdout: "piped",
        stderr: "piped"
      });

      const compressResult = await compressProcess.output();
      
      if (!compressResult.success) {
        throw new Error('Failed to create repaired archive');
      }

      // Read repaired file
      const repairedFileBytes = await Deno.readFile(outputFile);
      
      // Upload to Supabase storage
      const supabase = createClient(
        Deno.env.get('SUPABASE_URL') ?? '',
        Deno.env.get('SUPABASE_SERVICE_ROLE_KEY') ?? ''
      );

      const fileName = `repaired_${Date.now()}_${file.name}`;
      const { data: uploadData, error: uploadError } = await supabase.storage
        .from('file-repairs')
        .upload(fileName, repairedFileBytes, {
          contentType: file.type,
          upsert: false
        });

      if (uploadError) {
        throw new Error(`Failed to upload repaired file: ${uploadError.message}`);
      }

      // Get signed URL for download
      const { data: signedUrlData } = await supabase.storage
        .from('file-repairs')
        .createSignedUrl(fileName, 3600); // 1 hour expiry

      const result: RepairResult = {
        success: true,
        fileName: file.name,
        fileType: getFileType(file.type),
        originalSize: file.size,
        repairedSize: repairedFileBytes.length,
        issues: issues.length > 0 ? issues : undefined,
        repairedFileUrl: signedUrlData?.signedUrl,
        status: issues.length > 0 ? 'partial' : 'success'
      };

      console.log('File repair completed successfully');

      return new Response(JSON.stringify(result), {
        headers: { ...corsHeaders, 'Content-Type': 'application/json' },
      });

    } finally {
      // Clean up temporary files
      try {
        await Deno.remove(tempDir, { recursive: true });
      } catch (error) {
        console.warn('Failed to clean up temporary directory:', error);
      }
    }

  } catch (error) {
    console.error('Error in repair-office-file function:', error);
    
    const result: RepairResult = {
      success: false,
      fileName: 'unknown',
      fileType: 'unknown',
      originalSize: 0,
      status: 'failed',
      issues: [error.message]
    };

    return new Response(JSON.stringify(result), {
      status: 500,
      headers: { ...corsHeaders, 'Content-Type': 'application/json' },
    });
  }
});

async function repairXmlFiles(extractDir: string): Promise<string[]> {
  const issues: string[] = [];
  
  try {
    // Walk through all files in the extracted directory
    for await (const entry of Deno.readDir(extractDir)) {
      if (entry.isFile && entry.name.endsWith('.xml')) {
        const filePath = `${extractDir}/${entry.name}`;
        await repairXmlFile(filePath, issues);
      } else if (entry.isDirectory) {
        // Recursively process subdirectories
        const subdirIssues = await repairXmlFiles(`${extractDir}/${entry.name}`);
        issues.push(...subdirIssues);
      }
    }
  } catch (error) {
    console.warn('Error walking directory:', error);
    issues.push(`Directory traversal error: ${error.message}`);
  }
  
  return issues;
}

async function repairXmlFile(filePath: string, issues: string[]): Promise<void> {
  try {
    let content = await Deno.readTextFile(filePath);
    const originalLength = content.length;
    
    // Basic XML repair strategies
    
    // 1. Remove null bytes and control characters (except allowed ones)
    content = content.replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g, '');
    
    // 2. Fix common XML issues
    
    // Fix unclosed tags by attempting to match opening and closing tags
    const openTags = content.match(/<[^\/][^>]*>/g) || [];
    const closeTags = content.match(/<\/[^>]+>/g) || [];
    
    if (openTags.length !== closeTags.length) {
      issues.push(`XML structure mismatch in ${filePath.split('/').pop()}`);
      
      // Attempt to close unclosed tags at the end
      const openTagNames = openTags.map(tag => {
        const match = tag.match(/<([^\s>]+)/);
        return match ? match[1] : null;
      }).filter(Boolean);
      
      const closeTagNames = closeTags.map(tag => {
        const match = tag.match(/<\/([^>]+)>/);
        return match ? match[1] : null;
      }).filter(Boolean);
      
      // Find unclosed tags
      const unclosedTags = openTagNames.filter(tagName => 
        !closeTagNames.includes(tagName) && !tag.endsWith('/>')
      );
      
      // Add closing tags
      for (const tagName of unclosedTags) {
        content += `</${tagName}>`;
      }
    }
    
    // 3. Fix malformed attributes
    content = content.replace(/(\w+)=([^"\s>]+)(?=\s|>)/g, '$1="$2"');
    
    // 4. Remove truncated/incomplete tags at the end
    const lastGt = content.lastIndexOf('>');
    const lastLt = content.lastIndexOf('<');
    
    if (lastLt > lastGt) {
      // Truncate incomplete tag at the end
      content = content.substring(0, lastLt);
      issues.push(`Truncated incomplete tag in ${filePath.split('/').pop()}`);
    }
    
    // 5. Ensure XML declaration exists and is well-formed
    if (!content.startsWith('<?xml')) {
      content = '<?xml version="1.0" encoding="UTF-8"?>\n' + content;
    }
    
    // 6. Basic entity reference fixes
    content = content.replace(/&(?![a-zA-Z0-9#]+;)/g, '&amp;');
    
    if (content.length !== originalLength) {
      issues.push(`Repaired XML content in ${filePath.split('/').pop()}`);
      await Deno.writeTextFile(filePath, content);
    }
    
  } catch (error) {
    console.warn(`Failed to repair XML file ${filePath}:`, error);
    issues.push(`Failed to repair ${filePath.split('/').pop()}: ${error.message}`);
  }
}

function getFileType(mimeType: string): string {
  switch (mimeType) {
    case 'application/vnd.openxmlformats-officedocument.wordprocessingml.document':
      return 'DOCX';
    case 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet':
      return 'XLSX';
    case 'application/vnd.openxmlformats-officedocument.presentationml.presentation':
      return 'PPTX';
    default:
      return 'Unknown';
  }
}