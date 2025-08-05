import "https://deno.land/x/xhr@0.1.0/mod.ts";
import { serve } from "https://deno.land/std@0.168.0/http/server.ts";
import { createClient } from 'https://esm.sh/@supabase/supabase-js@2';
import JSZip from "https://esm.sh/jszip@3.10.1";
import yauzl from "https://esm.sh/yauzl@2.10.0";
import * as XLSX from "https://esm.sh/xlsx@0.18.5";
import mammoth from "https://esm.sh/mammoth@1.10.0";

const corsHeaders = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Headers': 'authorization, x-client-info, apikey, content-type',
};

interface RepairResult {
  success: boolean;
  fileName: string;
  status: 'success' | 'partial' | 'failed';
  issues?: string[];
  downloadUrl?: string;
  preview?: {
    content?: string;
    extractedSheets?: string[];
    extractedSlides?: number;
    recoveredFiles?: string[];
    fileSize?: number;
  };
  fileType?: 'DOCX' | 'XLSX' | 'PPTX' | 'ZIP' | 'PDF';
  recoveryStats?: {
    totalFiles: number;
    recoveredFiles: number;
    corruptedFiles: number;
  };
}

serve(async (req) => {
  // Handle CORS preflight requests
  if (req.method === 'OPTIONS') {
    return new Response(null, { headers: corsHeaders });
  }

  try {
    console.log('Starting enhanced Office file repair process');
    
    const formData = await req.formData();
    const file = formData.get('file') as File;
    
    if (!file) {
      return new Response(JSON.stringify({ error: 'No file provided' }), {
        status: 400,
        headers: { ...corsHeaders, 'Content-Type': 'application/json' },
      });
    }

    // Validate file type
    const allowedTypes = [
      'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 
      'application/vnd.openxmlformats-officedocument.presentationml.presentation',
      'application/zip',
      'application/x-zip-compressed',
      'application/pdf'
    ];

    if (!allowedTypes.includes(file.type)) {
      return new Response(JSON.stringify({ 
        error: 'Unsupported file type. Only DOCX, XLSX, PPTX, ZIP, and PDF files are supported.' 
      }), {
        status: 400,
        headers: { ...corsHeaders, 'Content-Type': 'application/json' },
      });
    }

    const fileData = await file.arrayBuffer();

    // Process the file with format-specific repair
    console.log(`Processing file: ${file.name}, size: ${fileData.byteLength}`);
    
    const fileType = getFileType(file.type);
    let repairedFile: Uint8Array;
    let preview: any = {};
    let recoveryStats = { totalFiles: 0, recoveredFiles: 0, corruptedFiles: 0 };
    const issues: string[] = [];

    try {
      // Use format-specific repair based on file type
      const repairResult = await repairOfficeFile(fileData, fileType, file.name);
      repairedFile = repairResult.data;
      preview = repairResult.preview;
      recoveryStats = repairResult.stats;
      issues.push(...repairResult.issues);
      
      console.log(`${fileType} repair successful with ${recoveryStats.recoveredFiles}/${recoveryStats.totalFiles} files recovered`);
    } catch (error) {
      console.log('Format-specific repair failed, trying generic repair:', error.message);
      issues.push(`${fileType}-specific repair failed: ${error.message}`);
      
      try {
        // Fallback to generic repair
        repairedFile = await advancedZipRepair(fileData);
        issues.push('Repaired using generic ZIP recovery');
      } catch (fallbackError) {
        console.log('All repair methods failed:', fallbackError.message);
        
        return new Response(
          JSON.stringify({
            success: false,
            fileName: file.name,
            status: 'failed',
            fileType,
            issues: ['Unable to repair file: corrupted beyond recovery']
          } as RepairResult),
          { 
            status: 400,
            headers: { ...corsHeaders, 'Content-Type': 'application/json' } 
          }
        );
      }
    }

    // Upload to Supabase storage
    const supabase = createClient(
      Deno.env.get('SUPABASE_URL') ?? '',
      Deno.env.get('SUPABASE_SERVICE_ROLE_KEY') ?? ''
    );

    const fileName = `repaired_${Date.now()}_${file.name}`;
    const { data: uploadData, error: uploadError } = await supabase.storage
      .from('file-repairs')
      .upload(fileName, repairedFile, {
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

    const signedUrl = signedUrlData?.signedUrl;

    return new Response(
      JSON.stringify({
        success: true,
        fileName: file.name,
        status: issues.length > 0 ? 'partial' : 'success',
        issues: issues.length > 0 ? issues : undefined,
        downloadUrl: signedUrl,
        fileType,
        preview,
        recoveryStats
      } as RepairResult),
      { headers: { ...corsHeaders, 'Content-Type': 'application/json' } }
    );

  } catch (error) {
    console.error('Error in repair-office-file function:', error);
    
    return new Response(
      JSON.stringify({
        success: false,
        fileName: 'unknown',
        status: 'failed',
        issues: [error.message]
      } as RepairResult),
      { 
        status: 500,
        headers: { ...corsHeaders, 'Content-Type': 'application/json' } 
      }
    );
  }
});

function getFileType(mimeType: string): 'DOCX' | 'XLSX' | 'PPTX' | 'ZIP' | 'PDF' {
  if (mimeType.includes('wordprocessingml')) return 'DOCX';
  if (mimeType.includes('spreadsheetml')) return 'XLSX';
  if (mimeType.includes('presentationml')) return 'PPTX';
  if (mimeType.includes('zip')) return 'ZIP';
  if (mimeType.includes('pdf')) return 'PDF';
  return 'DOCX'; // Default fallback
}

// Format-specific repair function
async function repairOfficeFile(
  fileData: ArrayBuffer, 
  fileType: 'DOCX' | 'XLSX' | 'PPTX' | 'ZIP' | 'PDF', 
  fileName: string
): Promise<{
  data: Uint8Array;
  preview: any;
  stats: { totalFiles: number; recoveredFiles: number; corruptedFiles: number };
  issues: string[];
}> {
  const issues: string[] = [];

  switch (fileType) {
    case 'DOCX':
      return await repairDocx(fileData, issues);
    case 'XLSX':
      return await repairXlsx(fileData, issues);
    case 'PPTX':
      return await repairPptx(fileData, issues);
    case 'ZIP':
      return await repairZip(fileData, issues);
    case 'PDF':
      return await repairPdf(fileData, issues);
    default:
      throw new Error(`Unsupported file type: ${fileType}`);
  }
}

// DOCX-specific repair
async function repairDocx(
  fileData: ArrayBuffer, 
  issues: string[]
): Promise<{
  data: Uint8Array;
  preview: any;
  stats: { totalFiles: number; recoveredFiles: number; corruptedFiles: number };
  issues: string[];
}> {
  const stats = { totalFiles: 0, recoveredFiles: 0, corruptedFiles: 0 };
  
  try {
    // Load with error tolerance
    const zip = await JSZip.loadAsync(fileData, { checkCRC32: false });
    const newZip = new JSZip();
    
    // Essential DOCX files
    const essentialFiles = [
      'word/document.xml',
      'word/styles.xml',
      '_rels/.rels',
      '[Content_Types].xml'
    ];
    
    let documentContent = '';
    
    for (const [path, file] of Object.entries(zip.files)) {
      if (!file.dir) {
        stats.totalFiles++;
        try {
          const content = await file.async('arraybuffer');
          newZip.file(path, content);
          stats.recoveredFiles++;
          
          // Extract document content for preview
          if (path === 'word/document.xml') {
            const xmlContent = await file.async('string');
            const repairedXml = repairDocumentXml(xmlContent);
            documentContent = extractTextFromDocumentXml(repairedXml);
            
            // If we repaired the XML, update it in the zip
            if (repairedXml !== xmlContent) {
              newZip.file(path, repairedXml);
              issues.push('Repaired corrupted document.xml with XML tag fixes');
            }
          }
        } catch (e) {
          stats.corruptedFiles++;
          issues.push(`Skipped corrupted file: ${path}`);
        }
      }
    }
    
    // Ensure essential files exist
    for (const essential of essentialFiles) {
      if (!newZip.file(essential)) {
        const minimalContent = generateMinimalDocxContent(essential);
        newZip.file(essential, minimalContent);
        issues.push(`Regenerated missing file: ${essential}`);
      }
    }
    
    const repairedData = await newZip.generateAsync({ type: 'uint8array' });
    
    const preview = {
      content: documentContent.slice(0, 500) + (documentContent.length > 500 ? '...' : ''),
      recoveredFiles: Object.keys(newZip.files).filter(f => !newZip.files[f].dir),
      fileSize: repairedData.length
    };
    
    return { data: repairedData, preview, stats, issues };
  } catch (error) {
    throw new Error(`DOCX repair failed: ${error.message}`);
  }
}

// XLSX-specific repair
async function repairXlsx(
  fileData: ArrayBuffer, 
  issues: string[]
): Promise<{
  data: Uint8Array;
  preview: any;
  stats: { totalFiles: number; recoveredFiles: number; corruptedFiles: number };
  issues: string[];
}> {
  const stats = { totalFiles: 0, recoveredFiles: 0, corruptedFiles: 0 };
  
  try {
    // Use SheetJS with error tolerance
    const workbook = XLSX.read(fileData, { 
      cellStyles: true, 
      sheetStubs: true,
      bookDeps: true,
      bookFiles: true,
      bookProps: true,
      bookSheets: true,
      bookVBA: true
    });
    
    const extractedSheets: string[] = [];
    let totalCells = 0;
    
    // Process each worksheet
    workbook.SheetNames.forEach(sheetName => {
      const worksheet = workbook.Sheets[sheetName];
      if (worksheet) {
        extractedSheets.push(sheetName);
        const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1:A1');
        totalCells += (range.e.r - range.s.r + 1) * (range.e.c - range.s.c + 1);
        stats.recoveredFiles++;
      } else {
        stats.corruptedFiles++;
        issues.push(`Worksheet "${sheetName}" is corrupted`);
      }
      stats.totalFiles++;
    });
    
    // Rebuild the workbook
    const repairedBuffer = XLSX.write(workbook, { 
      type: 'array',
      bookType: 'xlsx',
      compression: true 
    });
    
    const preview = {
      extractedSheets,
      totalCells,
      recoveredFiles: extractedSheets,
      fileSize: repairedBuffer.length
    };
    
    return { 
      data: new Uint8Array(repairedBuffer), 
      preview, 
      stats, 
      issues 
    };
  } catch (error) {
    throw new Error(`XLSX repair failed: ${error.message}`);
  }
}

// PPTX-specific repair
async function repairPptx(
  fileData: ArrayBuffer, 
  issues: string[]
): Promise<{
  data: Uint8Array;
  preview: any;
  stats: { totalFiles: number; recoveredFiles: number; corruptedFiles: number };
  issues: string[];
}> {
  const stats = { totalFiles: 0, recoveredFiles: 0, corruptedFiles: 0 };
  
  try {
    const zip = await JSZip.loadAsync(fileData, { checkCRC32: false });
    const newZip = new JSZip();
    
    let slideCount = 0;
    const slideFiles: string[] = [];
    
    for (const [path, file] of Object.entries(zip.files)) {
      if (!file.dir) {
        stats.totalFiles++;
        try {
          const content = await file.async('arraybuffer');
          newZip.file(path, content);
          stats.recoveredFiles++;
          
          // Count slides
          if (path.match(/ppt\/slides\/slide\d+\.xml/)) {
            slideCount++;
            slideFiles.push(path);
          }
        } catch (e) {
          stats.corruptedFiles++;
          issues.push(`Skipped corrupted file: ${path}`);
        }
      }
    }
    
    // Ensure essential PPTX structure
    const essentialFiles = [
      'ppt/presentation.xml',
      '_rels/.rels',
      '[Content_Types].xml'
    ];
    
    for (const essential of essentialFiles) {
      if (!newZip.file(essential)) {
        const minimalContent = generateMinimalPptxContent(essential);
        newZip.file(essential, minimalContent);
        issues.push(`Regenerated missing file: ${essential}`);
      }
    }
    
    const repairedData = await newZip.generateAsync({ type: 'uint8array' });
    
    const preview = {
      extractedSlides: slideCount,
      slideFiles,
      recoveredFiles: Object.keys(newZip.files).filter(f => !newZip.files[f].dir),
      fileSize: repairedData.length
    };
    
    return { data: repairedData, preview, stats, issues };
  } catch (error) {
    throw new Error(`PPTX repair failed: ${error.message}`);
  }
}

// Helper functions
function extractTextFromDocumentXml(xmlContent: string): string {
  // Simple text extraction from Word document XML
  const textMatches = xmlContent.match(/<w:t[^>]*>(.*?)<\/w:t>/g);
  if (!textMatches) return '';
  
  return textMatches
    .map(match => match.replace(/<[^>]*>/g, ''))
    .join(' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function generateMinimalDocxContent(filePath: string): string {
  switch (filePath) {
    case 'word/document.xml':
      return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t>Document recovered</w:t></w:r></w:p>
  </w:body>
</w:document>`;
    case '[Content_Types].xml':
      return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`;
    default:
      return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><root/>';
  }
}

function generateMinimalPptxContent(filePath: string): string {
  switch (filePath) {
    case 'ppt/presentation.xml':
      return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:sldMasterIdLst/>
  <p:sldIdLst/>
  <p:sldSz cx="9144000" cy="6858000"/>
</p:presentation>`;
    case '[Content_Types].xml':
      return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
</Types>`;
    default:
      return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><root/>';
  }
}

async function advancedZipRepair(arrayBuffer: ArrayBuffer): Promise<Uint8Array> {
  console.log('Using zip -FF command for ZIP repair...');
  
  try {
    // Create temporary files
    const tempDir = await Deno.makeTempDir();
    const inputPath = `${tempDir}/damaged.zip`;
    const outputPath = `${tempDir}/fixed.zip`;
    
    // Write damaged file
    await Deno.writeFile(inputPath, new Uint8Array(arrayBuffer));
    
    // Run zip -FF command
    const command = new Deno.Command("zip", {
      args: ["-FF", inputPath, "--out", outputPath],
      cwd: tempDir,
    });
    
    const { code, stderr } = await command.output();
    
    if (code === 0) {
      // Read repaired file
      const repairedData = await Deno.readFile(outputPath);
      
      // Cleanup
      await Deno.remove(tempDir, { recursive: true });
      
      console.log('ZIP repair successful using zip -FF command');
      return repairedData;
    } else {
      const errorText = new TextDecoder().decode(stderr);
      console.log('zip -FF failed:', errorText);
      
      // Fallback to JavaScript-based repair
      return await fallbackZipRepair(arrayBuffer);
    }
  } catch (error) {
    console.log('zip -FF command not available, using fallback:', error.message);
    return await fallbackZipRepair(arrayBuffer);
  }
}

async function fallbackZipRepair(arrayBuffer: ArrayBuffer): Promise<Uint8Array> {
  console.log('Using advanced binary ZIP repair...');
  console.log('Input data size:', arrayBuffer.byteLength, 'bytes');
  
  const uint8Array = new Uint8Array(arrayBuffer);
  const extractedFiles: Record<string, Uint8Array> = {};
  
  // Step 1: Scan for ZIP local file header signatures (0x04034b50)
  const localFileHeaders: number[] = [];
  console.log('Scanning for ZIP signatures...');
  
  for (let i = 0; i < uint8Array.length - 4; i++) {
    if (uint8Array[i] === 0x50 && uint8Array[i+1] === 0x4b && 
        uint8Array[i+2] === 0x03 && uint8Array[i+3] === 0x04) {
      localFileHeaders.push(i);
      console.log('Found ZIP signature at offset', i);
    }
  }
  
  console.log(`Found ${localFileHeaders.length} potential file headers`);
  
  // Step 2: Extract files and rebuild using JSZip for proper decompression
  const newZip = new JSZip();
  let recoveredCount = 0;
  
  for (const headerOffset of localFileHeaders) {
    try {
      const fileInfo = await parseLocalFileHeader(uint8Array, headerOffset);
      if (fileInfo && fileInfo.filename && fileInfo.data && !fileInfo.filename.endsWith('/')) {
        // For ZIP repair, add the decompressed data to avoid JSZip auto-decompression issues
        newZip.file(fileInfo.filename, fileInfo.data, { compression: 'STORE' });
        recoveredCount++;
        console.log(`Recovered file: ${fileInfo.filename} (${fileInfo.data.length} bytes)`);
      }
    } catch (error) {
      console.log(`Failed to parse header at offset ${headerOffset}:`, error.message);
    }
  }
  
  if (recoveredCount === 0) {
    throw new Error('No files could be recovered from corrupt ZIP');
  }
  
  console.log(`Successfully recovered ${recoveredCount} files`);
  
  // Generate new ZIP with proper compression handling
  const repairedZip = await newZip.generateAsync({ 
    type: 'uint8array',
    compression: 'DEFLATE',
    compressionOptions: { level: 6 }
  });
  
  return repairedZip;
}

// Parse ZIP local file header manually (mimics zip -FF behavior)
async function parseLocalFileHeader(data: Uint8Array, offset: number): Promise<{ filename: string; data: Uint8Array } | null> {
  try {
    if (offset + 30 > data.length) return null;
    
    const filenameLength = data[offset + 26] | (data[offset + 27] << 8);
    const extraFieldLength = data[offset + 28] | (data[offset + 29] << 8);
    const compressedSize = (data[offset + 18] | (data[offset + 19] << 8) | 
                           (data[offset + 20] << 16) | (data[offset + 21] << 24)) >>> 0;
    const uncompressedSize = (data[offset + 22] | (data[offset + 23] << 8) | 
                             (data[offset + 24] << 16) | (data[offset + 25] << 24)) >>> 0;
    const compressionMethod = data[offset + 8] | (data[offset + 9] << 8);
    
    const filenameStart = offset + 30;
    const dataStart = filenameStart + filenameLength + extraFieldLength;
    
    // Extract filename
    if (filenameStart + filenameLength > data.length) return null;
    const filename = new TextDecoder().decode(data.slice(filenameStart, filenameStart + filenameLength));
    
    // Skip directories
    if (filename.endsWith('/')) return null;
    
    // Extract file data
    let compressedData: Uint8Array;
    if (compressedSize > 0 && dataStart + compressedSize <= data.length) {
      compressedData = data.slice(dataStart, dataStart + compressedSize);
    } else {
      // Size corrupted, find next header or use remaining data
      let endOffset = data.length;
      for (let i = dataStart + 1; i < data.length - 4; i++) {
        if (data[i] === 0x50 && data[i+1] === 0x4b && 
            data[i+2] === 0x03 && data[i+3] === 0x04) {
          endOffset = i;
          break;
        }
      }
      compressedData = data.slice(dataStart, endOffset);
    }
    
    // Debug info about the file
    console.log(`Processing ${filename}: compression=${compressionMethod}, compressed=${compressedData.length} bytes`);
    console.log(`First 32 bytes of compressed data:`, Array.from(compressedData.slice(0, 32)).map(b => b.toString(16).padStart(2, '0')).join(' '));
    
    // Decompress data if needed
    let fileData: Uint8Array;
    if (compressionMethod === 8) { // DEFLATE compression
      try {
        // ZIP uses raw deflate without zlib headers
        // Try different decompression approaches
        let decompressed = false;
        
        // First try: deflate-raw (raw deflate without zlib wrapper)
        try {
          const stream1 = new ReadableStream({
            start(controller) {
              controller.enqueue(compressedData);
              controller.close();
            }
          });
          const decompressedStream1 = stream1.pipeThrough(new DecompressionStream('deflate-raw'));
          const response1 = new Response(decompressedStream1);
          const decompressedBuffer1 = await response1.arrayBuffer();
          fileData = new Uint8Array(decompressedBuffer1);
          console.log(`Successfully decompressed ${filename} with deflate-raw: ${compressedData.length} -> ${fileData.length} bytes`);
          decompressed = true;
        } catch (rawError) {
          console.log(`deflate-raw failed for ${filename}:`, rawError.message);
        }
        
        // Second try: standard deflate (with zlib wrapper)
        if (!decompressed) {
          try {
            const stream2 = new ReadableStream({
              start(controller) {
                controller.enqueue(compressedData);
                controller.close();
              }
            });
            const decompressedStream2 = stream2.pipeThrough(new DecompressionStream('deflate'));
            const response2 = new Response(decompressedStream2);
            const decompressedBuffer2 = await response2.arrayBuffer();
            fileData = new Uint8Array(decompressedBuffer2);
            console.log(`Successfully decompressed ${filename} with deflate: ${compressedData.length} -> ${fileData.length} bytes`);
            decompressed = true;
          } catch (deflateError) {
            console.log(`deflate failed for ${filename}:`, deflateError.message);
          }
        }
        
        // Third try: For corrupted deflate streams, try to extract readable content
        if (!decompressed && filename.endsWith('.xml')) {
          console.log(`Attempting manual XML recovery for ${filename}`);
          
          // Try alternative decompression methods first
          let extractedText = '';
          
          // Try gzip decompression (sometimes ZIP files have different compression)
          try {
            const gzipStream = new ReadableStream({
              start(controller) {
                controller.enqueue(compressedData);
                controller.close();
              }
            });
            const decompressedGzip = gzipStream.pipeThrough(new DecompressionStream('gzip'));
            const responseGzip = new Response(decompressedGzip);
            const decompressedBuffer = await responseGzip.arrayBuffer();
            const xmlText = new TextDecoder('utf-8', { fatal: false }).decode(decompressedBuffer);
            
            // Check if this looks like valid XML
            if (xmlText.includes('<?xml') || xmlText.includes('<w:')) {
              fileData = new Uint8Array(decompressedBuffer);
              console.log(`Successfully decompressed ${filename} with gzip: ${compressedData.length} -> ${fileData.length} bytes`);
              decompressed = true;
            }
          } catch (gzipError) {
            console.log(`Gzip failed for ${filename}:`, gzipError.message);
          }
          
          // If inflate didn't work, try to find readable text in a smarter way
          if (!decompressed) {
            try {
              // Convert to string and look for actual readable sentences
              const textDecoder = new TextDecoder('utf-8', { fatal: false });
              const rawText = textDecoder.decode(compressedData);
              
              // Look for sequences of actual words (3+ letters, with spaces)
              const wordPattern = /\b[a-zA-Z]{3,}(?:\s+[a-zA-Z]{2,})*\b/g;
              const sentences = rawText.match(wordPattern);
              
              if (sentences && sentences.length > 0) {
                // Join sentences that seem to be actual text
                extractedText = sentences
                  .filter(sentence => {
                    // Filter out sequences that are mostly special characters
                    const letterCount = (sentence.match(/[a-zA-Z]/g) || []).length;
                    const totalLength = sentence.length;
                    return letterCount / totalLength > 0.7; // At least 70% letters
                  })
                  .join('. ')
                  .trim();
                  
                // Clean up the text
                if (extractedText.length > 10) {
                  extractedText = extractedText
                    .replace(/\s+/g, ' ') // Multiple spaces to single space
                    .replace(/[^\w\s.,!?;:()\-'"]/g, '') // Remove weird characters
                    .trim();
                }
              }
              
              // If no good text found, look for any XML content patterns
              if (!extractedText) {
                const xmlPattern = /<w:t[^>]*>([^<]+)<\/w:t>/g;
                let match;
                const textParts = [];
                
                while ((match = xmlPattern.exec(rawText)) !== null) {
                  const text = match[1].trim();
                  if (text.length > 2 && /[a-zA-Z]/.test(text)) {
                    textParts.push(text);
                  }
                }
                
                if (textParts.length > 0) {
                  extractedText = textParts.join(' ');
                }
              }
            } catch (textError) {
              console.log(`Text extraction failed for ${filename}:`, textError.message);
            }
          }
          
          // Create appropriate XML content based on filename
          let xmlContent = '';
          if (filename === 'word/document.xml') {
            const documentText = extractedText || 'Document content could not be recovered due to compression corruption.';
            xmlContent = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p>
      <w:r>
        <w:t>${documentText}</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>`;
          } else {
            // Generic XML for other files
            xmlContent = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<root>
  <!-- File ${filename} recovered with minimal content -->
</root>`;
          }
          
          fileData = new TextEncoder().encode(xmlContent);
          console.log(`Created replacement XML content for ${filename}: ${fileData.length} bytes`);
          if (extractedText) {
            console.log(`Extracted text: ${extractedText.substring(0, 100)}...`);
          }
          decompressed = true;
        }
        
        if (!decompressed) {
          console.log(`All decompression methods failed for ${filename}, using compressed data`);
          fileData = compressedData;
        }
      } catch (error) {
        console.log(`Failed to decompress ${filename}, using compressed data:`, error.message);
        fileData = compressedData;
      }
    } else if (compressionMethod === 0) {
      // No compression
      console.log(`${filename} is not compressed`);
      fileData = compressedData;
    } else {
      // Unknown compression method
      console.log(`${filename} uses unknown compression method ${compressionMethod}`);
      fileData = compressedData;
    }
    
    return { filename, data: fileData };
  } catch (error) {
    return null;
  }
}


// ZIP-specific repair
async function repairZip(
  fileData: ArrayBuffer, 
  issues: string[]
): Promise<{
  data: Uint8Array;
  preview: any;
  stats: { totalFiles: number; recoveredFiles: number; corruptedFiles: number };
  issues: string[];
}> {
  const stats = { totalFiles: 0, recoveredFiles: 0, corruptedFiles: 0 };
  
  try {
    // Use yauzl for robust ZIP extraction
    const repairedData = await advancedZipRepair(fileData);
    
    // Count recovered files by re-reading the repaired ZIP
    const zip = await JSZip.loadAsync(repairedData);
    const recoveredFiles: string[] = [];
    
    for (const [path, file] of Object.entries(zip.files)) {
      if (!file.dir) {
        stats.totalFiles++;
        stats.recoveredFiles++;
        recoveredFiles.push(path);
      }
    }
    
    const preview = {
      recoveredFiles,
      fileSize: repairedData.length,
      content: `ZIP archive with ${recoveredFiles.length} recovered files`
    };
    
    return { data: repairedData, preview, stats, issues };
  } catch (error) {
    throw new Error(`ZIP repair failed: ${error.message}`);
  }
}

// PDF-specific repair
async function repairPdf(
  fileData: ArrayBuffer, 
  issues: string[]
): Promise<{
  data: Uint8Array;
  preview: any;
  stats: { totalFiles: number; recoveredFiles: number; corruptedFiles: number };
  issues: string[];
}> {
  const stats = { totalFiles: 1, recoveredFiles: 0, corruptedFiles: 0 };
  
  try {
    const uint8Array = new Uint8Array(fileData);
    
    // Basic PDF repair - find PDF header and trailer
    const pdfHeader = '%PDF-';
    const pdfTrailer = '%%EOF';
    
    let headerIndex = -1;
    let trailerIndex = -1;
    
    // Find PDF header
    for (let i = 0; i <= uint8Array.length - 5; i++) {
      const chunk = new TextDecoder().decode(uint8Array.slice(i, i + 5));
      if (chunk === pdfHeader) {
        headerIndex = i;
        break;
      }
    }
    
    // Find PDF trailer (search from end)
    for (let i = uint8Array.length - 5; i >= 0; i--) {
      const chunk = new TextDecoder().decode(uint8Array.slice(i, i + 5));
      if (chunk === pdfTrailer) {
        trailerIndex = i + 5;
        break;
      }
    }
    
    let repairedData = uint8Array;
    
    if (headerIndex > 0) {
      // Remove garbage before PDF header
      repairedData = uint8Array.slice(headerIndex);
      issues.push('Removed garbage data before PDF header');
    }
    
    if (trailerIndex > 0 && trailerIndex < repairedData.length) {
      // Truncate after EOF marker
      repairedData = repairedData.slice(0, trailerIndex - headerIndex);
      issues.push('Truncated data after PDF EOF marker');
    }
    
    if (headerIndex >= 0) {
      stats.recoveredFiles = 1;
    } else {
      stats.corruptedFiles = 1;
      issues.push('PDF header not found - file may be severely corrupted');
    }
    
    const preview = {
      fileSize: repairedData.length,
      content: `PDF document (${(repairedData.length / 1024).toFixed(1)} KB)`
    };
    
    return { data: repairedData, preview, stats, issues };
  } catch (error) {
    throw new Error(`PDF repair failed: ${error.message}`);
  }
}

// XML repair function for document.xml
function repairDocumentXml(xmlContent: string): string {
  try {
    // Parse the XML to find content
    const parser = new DOMParser();
    const doc = parser.parseFromString(xmlContent, 'text/xml');
    
    // If parsing fails, try to truncate to last valid paragraph
    if (doc.querySelector('parsererror')) {
      // Find the last complete paragraph tag
      const lastParagraphMatch = xmlContent.lastIndexOf('</w:p>');
      if (lastParagraphMatch !== -1) {
        const truncatedContent = xmlContent.substring(0, lastParagraphMatch + 6);
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
    
    return xmlContent;
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
}

async function rebuildZipFromExtractedFiles(extractedFiles: { [key: string]: Uint8Array }): Promise<ArrayBuffer> {
  const zip = new JSZip();
  
  for (const [fileName, fileData] of Object.entries(extractedFiles)) {
    zip.file(fileName, fileData);
  }
  
  return await zip.generateAsync({
    type: 'arraybuffer',
    compression: 'DEFLATE',
    compressionOptions: { level: 6 }
  });
}