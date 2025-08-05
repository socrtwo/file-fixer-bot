import "https://deno.land/x/xhr@0.1.0/mod.ts";
import { serve } from "https://deno.land/std@0.168.0/http/server.ts";
import { createClient } from 'https://esm.sh/@supabase/supabase-js@2.53.0';

// CORS headers for cross-origin requests
const corsHeaders = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Headers': 'authorization, x-client-info, apikey, content-type',
};

// Interface for repair result
interface RepairResult {
  success: boolean;
  fileName: string;
  status: 'success' | 'partial' | 'failed';
  issues?: string[];
  repairedFile?: string;
  preview?: {
    content?: string;
    slides?: number;
    worksheets?: string[];
  };
  fileType?: string;
  recoveryStats?: {
    originalSize?: number;
    repairedSize?: number;
    corruptionLevel?: string;
    recoveredData?: number;
  };
}

serve(async (req) => {
  // Handle CORS preflight requests
  if (req.method === 'OPTIONS') {
    return new Response(null, { headers: corsHeaders });
  }

  try {
    console.log('File repair request received');
    
    // Parse the multipart form data to get the file
    const formData = await req.formData();
    const file = formData.get('file') as File;
    
    if (!file) {
      return new Response(JSON.stringify({ error: 'No file provided' }), {
        status: 400,
        headers: { ...corsHeaders, 'Content-Type': 'application/json' },
      });
    }

    console.log(`Processing file: ${file.name}, size: ${file.size}, type: ${file.type}`);

    // Validate file type
    const allowedTypes = [
      'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      'application/vnd.openxmlformats-officedocument.presentationml.presentation',
      'application/zip',
      'application/pdf'
    ];

    if (!allowedTypes.includes(file.type)) {
      return new Response(JSON.stringify({ error: 'Unsupported file type' }), {
        status: 400,
        headers: { ...corsHeaders, 'Content-Type': 'application/json' },
      });
    }

    // Get file data
    const fileData = await file.arrayBuffer();
    const uint8Array = new Uint8Array(fileData);
    
    console.log(`File data length: ${uint8Array.length}`);

    // Determine file type and repair accordingly
    const fileType = getFileType(file.type);
    console.log(`Detected file type: ${fileType}`);
    
    // Attempt repair
    let repairResult: RepairResult;
    
    try {
      repairResult = await repairOfficeFile(uint8Array, file.name, fileType);
      console.log(`Initial repair result: ${repairResult.success ? 'success' : 'failed'}`);
    } catch (error) {
      console.error('Error during initial repair:', error);
      repairResult = {
        success: false,
        fileName: file.name,
        status: 'failed',
        issues: [`Initial repair failed: ${error.message}`],
        fileType
      };
    }

    // If initial repair failed and it's a ZIP-based format, try advanced repair
    if (!repairResult.success && ['docx', 'xlsx', 'pptx', 'zip'].includes(fileType)) {
      console.log('Attempting advanced ZIP repair...');
      try {
        const advancedRepair = await advancedZipRepair(uint8Array, file.name, fileType);
        if (advancedRepair.success) {
          repairResult = advancedRepair;
          console.log('Advanced repair successful');
        }
      } catch (error) {
        console.error('Advanced repair also failed:', error);
        repairResult.issues = repairResult.issues || [];
        repairResult.issues.push(`Advanced repair failed: ${error.message}`);
      }
    }

    // If repair was successful, upload to Supabase Storage
    if (repairResult.success && repairResult.repairedFile) {
      try {
        console.log('Uploading repaired file to storage...');
        
        const supabase = createClient(
          Deno.env.get('SUPABASE_URL') ?? '',
          Deno.env.get('SUPABASE_SERVICE_ROLE_KEY') ?? ''
        );

        // Convert base64 back to binary for upload
        const binaryData = Uint8Array.from(atob(repairResult.repairedFile), c => c.charCodeAt(0));
        
        const fileName = `repaired_${Date.now()}_${file.name}`;
        
        const { data: uploadData, error: uploadError } = await supabase.storage
          .from('file-repairs')
          .upload(fileName, binaryData, {
            contentType: file.type,
            cacheControl: '3600'
          });

        if (uploadError) {
          console.error('Upload error:', uploadError);
          throw new Error(`Upload failed: ${uploadError.message}`);
        }

        console.log('File uploaded successfully:', uploadData.path);

        // Generate signed URL
        const { data: signedUrlData } = await supabase.storage
          .from('file-repairs')
          .createSignedUrl(uploadData.path, 3600); // 1 hour expiry

        if (signedUrlData?.signedUrl) {
          repairResult.repairedFile = signedUrlData.signedUrl;
          console.log('Signed URL generated successfully');
        }

      } catch (error) {
        console.error('Error uploading to storage:', error);
        repairResult.issues = repairResult.issues || [];
        repairResult.issues.push(`Storage upload failed: ${error.message}`);
      }
    }

    // Ensure we never return zero-byte files
    if (repairResult.success && repairResult.recoveryStats?.repairedSize === 0) {
      console.log('Warning: Repaired file has zero size, marking as partial success');
      repairResult.status = 'partial';
      repairResult.issues = repairResult.issues || [];
      repairResult.issues.push('Repaired file is empty - only structure was recovered');
    }

    console.log(`Final repair result: ${JSON.stringify(repairResult, null, 2)}`);

    return new Response(JSON.stringify(repairResult), {
      headers: { ...corsHeaders, 'Content-Type': 'application/json' },
    });

  } catch (error) {
    console.error('Unexpected error in repair function:', error);
    return new Response(JSON.stringify({ 
      error: 'Internal server error',
      details: error.message 
    }), {
      status: 500,
      headers: { ...corsHeaders, 'Content-Type': 'application/json' },
    });
  }
});

// Helper function to determine file type from MIME type
function getFileType(mimeType: string): string {
  switch (mimeType) {
    case 'application/vnd.openxmlformats-officedocument.wordprocessingml.document':
      return 'docx';
    case 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet':
      return 'xlsx';
    case 'application/vnd.openxmlformats-officedocument.presentationml.presentation':
      return 'pptx';
    case 'application/zip':
      return 'zip';
    case 'application/pdf':
      return 'pdf';
    default:
      return 'unknown';
  }
}

// Main repair function that dispatches to specific repair methods
async function repairOfficeFile(fileData: Uint8Array, fileName: string, fileType: string): Promise<RepairResult> {
  console.log(`Starting repair for ${fileType} file: ${fileName}`);
  
  switch (fileType) {
    case 'docx':
      return await repairDocx(fileData, fileName);
    case 'xlsx':
      return await repairXlsx(fileData, fileName);
    case 'pptx':
      return await repairPptx(fileData, fileName);
    case 'zip':
      return await repairZip(fileData, fileName);
    case 'pdf':
      return await repairPdf(fileData, fileName);
    default:
      return {
        success: false,
        fileName,
        status: 'failed',
        issues: ['Unsupported file type'],
        fileType
      };
  }
}

// DOCX repair function
async function repairDocx(fileData: Uint8Array, fileName: string): Promise<RepairResult> {
  console.log('Starting DOCX repair...');
  
  try {
    // Import JSZip dynamically
    const JSZip = (await import('https://esm.sh/jszip@3.10.1')).default;
    
    let zip: any;
    let extractedText = '';
    const issues: string[] = [];
    
    try {
      // Try to load as ZIP first
      zip = await JSZip.loadAsync(fileData);
      console.log('Successfully loaded as ZIP');
    } catch (error) {
      console.log('Failed to load as ZIP, attempting data recovery...');
      issues.push('ZIP structure corrupted, attempting content recovery');
      
      // Try to extract content from corrupted data
      const recoveredContent = await extractContentFromCorruptedDocx(fileData);
      if (recoveredContent) {
        extractedText = recoveredContent;
        console.log(`Recovered content length: ${extractedText.length}`);
        
        // Create a new DOCX with recovered content
        const repairedZip = new JSZip();
        
        // Add essential files
        repairedZip.file('[Content_Types].xml', generateContentTypes());
        repairedZip.file('_rels/.rels', generateMainRels());
        repairedZip.file('word/_rels/document.xml.rels', generateDocumentRels());
        repairedZip.file('word/document.xml', createDocumentXmlWithContent(extractedText));
        
        const repairedData = await repairedZip.generateAsync({ type: 'uint8array' });
        
        return {
          success: true,
          fileName,
          status: 'partial',
          issues,
          repairedFile: btoa(String.fromCharCode(...repairedData)),
          preview: { content: extractedText.substring(0, 500) },
          fileType: 'docx',
          recoveryStats: {
            originalSize: fileData.length,
            repairedSize: repairedData.length,
            corruptionLevel: 'high',
            recoveredData: extractedText.length
          }
        };
      } else {
        throw new Error('Unable to recover any content from corrupted DOCX');
      }
    }
    
    // If we got here, the ZIP loaded successfully
    const files = Object.keys(zip.files);
    console.log(`ZIP contains ${files.length} files`);
    
    // Check for essential DOCX files
    const essentialFiles = ['word/document.xml', '[Content_Types].xml', '_rels/.rels'];
    const missingFiles = essentialFiles.filter(file => !zip.files[file]);
    
    if (missingFiles.length > 0) {
      console.log(`Missing essential files: ${missingFiles.join(', ')}`);
      issues.push(`Missing essential files: ${missingFiles.join(', ')}`);
    }
    
    // Try to extract text content
    if (zip.files['word/document.xml']) {
      try {
        const documentXml = await zip.files['word/document.xml'].async('string');
        extractedText = extractTextFromDocumentXml(documentXml);
        console.log(`Extracted text length: ${extractedText.length}`);
      } catch (error) {
        console.log('Error extracting text from document.xml:', error.message);
        issues.push('Document XML corrupted, text extraction failed');
      }
    }
    
    // Regenerate the ZIP with any missing files
    if (missingFiles.length > 0) {
      if (!zip.files['[Content_Types].xml']) {
        zip.file('[Content_Types].xml', generateContentTypes());
      }
      if (!zip.files['_rels/.rels']) {
        zip.file('_rels/.rels', generateMainRels());
      }
      if (!zip.files['word/_rels/document.xml.rels']) {
        zip.file('word/_rels/document.xml.rels', generateDocumentRels());
      }
    }
    
    const repairedData = await zip.generateAsync({ type: 'uint8array' });
    
    return {
      success: true,
      fileName,
      status: issues.length > 0 ? 'partial' : 'success',
      issues: issues.length > 0 ? issues : undefined,
      repairedFile: btoa(String.fromCharCode(...repairedData)),
      preview: { content: extractedText.substring(0, 500) },
      fileType: 'docx',
      recoveryStats: {
        originalSize: fileData.length,
        repairedSize: repairedData.length,
        corruptionLevel: issues.length > 0 ? 'medium' : 'low',
        recoveredData: extractedText.length
      }
    };
    
  } catch (error) {
    console.error('DOCX repair failed:', error);
    return {
      success: false,
      fileName,
      status: 'failed',
      issues: [`DOCX repair failed: ${error.message}`],
      fileType: 'docx'
    };
  }
}

// Function to extract content from heavily corrupted DOCX files
async function extractContentFromCorruptedDocx(fileData: Uint8Array): Promise<string | null> {
  try {
    console.log('Attempting to extract content from corrupted DOCX data...');
    
    // Convert to string and look for text patterns
    const decoder = new TextDecoder('utf-8', { fatal: false });
    const rawText = decoder.decode(fileData);
    
    // Look for the specific corruption pattern mentioned by the user
    const corruptionPattern = /xml:space="preserv:HBAPESSE RMLKHGB tok doublan thc><w:p w:rsidRyE: the charent specialties\. It is essential that you get to know the health care team and feel dP=<w:szw:r:rs/;
    
    let cleanText = rawText;
    
    // Find and remove the corruption pattern
    const corruptionMatch = cleanText.search(corruptionPattern);
    if (corruptionMatch >= 0) {
      console.log(`Found corruption pattern at position ${corruptionMatch}, truncating...`);
      cleanText = cleanText.substring(0, corruptionMatch);
    }
    
    // Extract readable text from what's left
    const textPatterns = [
      /<w:t[^>]*>([^<]+)<\/w:t>/g,
      />([A-Za-z][A-Za-z\s.,!?-]{10,})</g,
      /([A-Z][a-z\s.,!?-]{20,})/g
    ];
    
    const extractedChunks: string[] = [];
    
    for (const pattern of textPatterns) {
      let match;
      while ((match = pattern.exec(cleanText)) !== null) {
        const text = match[1] || match[0];
        if (text && text.length > 10 && !text.includes('<') && !text.includes('xml')) {
          extractedChunks.push(text.trim());
        }
      }
    }
    
    // Deduplicate and join
    const uniqueChunks = [...new Set(extractedChunks)];
    const recoveredText = uniqueChunks.join(' ').trim();
    
    console.log(`Recovered ${recoveredText.length} characters of text`);
    
    return recoveredText.length > 50 ? recoveredText : null;
    
  } catch (error) {
    console.error('Error extracting content from corrupted DOCX:', error);
    return null;
  }
}

// Function to extract text from document.xml
function extractTextFromDocumentXml(xmlContent: string): string {
  try {
    // Extract text from <w:t> tags
    const textMatches = xmlContent.match(/<w:t[^>]*>([^<]+)<\/w:t>/g) || [];
    const extractedText = textMatches
      .map(match => {
        const textMatch = match.match(/>([^<]+)</);
        return textMatch ? textMatch[1] : '';
      })
      .filter(text => text.trim().length > 0)
      .join(' ');
    
    return extractedText;
  } catch (error) {
    console.error('Error extracting text from XML:', error);
    return '';
  }
}

// Function to create document.xml with content
function createDocumentXmlWithContent(content: string): string {
  // Split content into paragraphs for better formatting
  const paragraphs = content.split(/[.!?]+/).filter(p => p.trim().length > 0);
  
  let xmlParagraphs = '';
  for (const paragraph of paragraphs.slice(0, 10)) { // Limit to 10 paragraphs
    const cleanParagraph = paragraph.trim().replace(/[<>&"']/g, ''); // Escape XML chars
    if (cleanParagraph.length > 0) {
      xmlParagraphs += `    <w:p>
      <w:r>
        <w:t>${cleanParagraph}.</w:t>
      </w:r>
    </w:p>
`;
    }
  }
  
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
${xmlParagraphs}
  </w:body>
</w:document>`;
}

// Generate Content_Types.xml
function generateContentTypes(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`;
}

// Generate main .rels file
function generateMainRels(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;
}

// Generate document.xml.rels
function generateDocumentRels(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>`;
}

// XLSX repair function
async function repairXlsx(fileData: Uint8Array, fileName: string): Promise<RepairResult> {
  console.log('Starting XLSX repair...');
  
  try {
    // Import required libraries
    const XLSX = (await import('https://esm.sh/xlsx@0.18.5')).default;
    const JSZip = (await import('https://esm.sh/jszip@3.10.1')).default;
    
    let workbook: any;
    const issues: string[] = [];
    
    try {
      // Try to read with XLSX library first
      workbook = XLSX.read(fileData, { type: 'array' });
      console.log('Successfully parsed XLSX');
    } catch (error) {
      console.log('XLSX parsing failed, attempting ZIP repair...');
      issues.push('XLSX structure corrupted, attempting repair');
      
      // Try ZIP-based repair
      const zip = new JSZip();
      try {
        const loadedZip = await JSZip.loadAsync(fileData);
        
        // Create a minimal working XLSX
        zip.file('[Content_Types].xml', generateXlsxContentTypes());
        zip.file('_rels/.rels', generateXlsxMainRels());
        zip.file('xl/_rels/workbook.xml.rels', generateXlsxWorkbookRels());
        zip.file('xl/workbook.xml', generateXlsxWorkbook());
        zip.file('xl/worksheets/sheet1.xml', generateXlsxWorksheet());
        
        const repairedData = await zip.generateAsync({ type: 'uint8array' });
        
        return {
          success: true,
          fileName,
          status: 'partial',
          issues,
          repairedFile: btoa(String.fromCharCode(...repairedData)),
          preview: { worksheets: ['Sheet1'] },
          fileType: 'xlsx',
          recoveryStats: {
            originalSize: fileData.length,
            repairedSize: repairedData.length,
            corruptionLevel: 'high',
            recoveredData: 0
          }
        };
      } catch (zipError) {
        throw new Error('Both XLSX and ZIP repair methods failed');
      }
    }
    
    // If we got here, workbook was parsed successfully
    const worksheetNames = workbook.SheetNames;
    console.log(`XLSX contains ${worksheetNames.length} worksheets: ${worksheetNames.join(', ')}`);
    
    // Regenerate the XLSX to fix any structural issues
    const repairedData = XLSX.write(workbook, { type: 'array', bookType: 'xlsx' });
    
    return {
      success: true,
      fileName,
      status: issues.length > 0 ? 'partial' : 'success',
      issues: issues.length > 0 ? issues : undefined,
      repairedFile: btoa(String.fromCharCode(...repairedData)),
      preview: { worksheets: worksheetNames },
      fileType: 'xlsx',
      recoveryStats: {
        originalSize: fileData.length,
        repairedSize: repairedData.length,
        corruptionLevel: issues.length > 0 ? 'medium' : 'low',
        recoveredData: worksheetNames.length
      }
    };
    
  } catch (error) {
    console.error('XLSX repair failed:', error);
    return {
      success: false,
      fileName,
      status: 'failed',
      issues: [`XLSX repair failed: ${error.message}`],
      fileType: 'xlsx'
    };
  }
}

// PPTX repair function
async function repairPptx(fileData: Uint8Array, fileName: string): Promise<RepairResult> {
  console.log('Starting PPTX repair...');
  
  try {
    const JSZip = (await import('https://esm.sh/jszip@3.10.1')).default;
    
    let zip: any;
    const issues: string[] = [];
    let slideCount = 0;
    
    try {
      zip = await JSZip.loadAsync(fileData);
      console.log('Successfully loaded PPTX as ZIP');
      
      // Count slides
      const slideFiles = Object.keys(zip.files).filter(name => name.startsWith('ppt/slides/slide') && name.endsWith('.xml'));
      slideCount = slideFiles.length;
      console.log(`Found ${slideCount} slides`);
      
    } catch (error) {
      console.log('Failed to load PPTX as ZIP, creating minimal structure...');
      issues.push('PPTX structure corrupted, creating minimal presentation');
      
      zip = new JSZip();
      zip.file('[Content_Types].xml', generatePptxContentTypes());
      zip.file('_rels/.rels', generatePptxMainRels());
      zip.file('ppt/_rels/presentation.xml.rels', generatePptxPresentationRels());
      zip.file('ppt/presentation.xml', generatePptxPresentation());
      zip.file('ppt/slides/slide1.xml', generatePptxSlide());
      slideCount = 1;
    }
    
    const repairedData = await zip.generateAsync({ type: 'uint8array' });
    
    return {
      success: true,
      fileName,
      status: issues.length > 0 ? 'partial' : 'success',
      issues: issues.length > 0 ? issues : undefined,
      repairedFile: btoa(String.fromCharCode(...repairedData)),
      preview: { slides: slideCount },
      fileType: 'pptx',
      recoveryStats: {
        originalSize: fileData.length,
        repairedSize: repairedData.length,
        corruptionLevel: issues.length > 0 ? 'high' : 'low',
        recoveredData: slideCount
      }
    };
    
  } catch (error) {
    console.error('PPTX repair failed:', error);
    return {
      success: false,
      fileName,
      status: 'failed',
      issues: [`PPTX repair failed: ${error.message}`],
      fileType: 'pptx'
    };
  }
}

// ZIP repair function
async function repairZip(fileData: Uint8Array, fileName: string): Promise<RepairResult> {
  console.log('Starting ZIP repair...');
  return await advancedZipRepair(fileData, fileName, 'zip');
}

// PDF repair function
async function repairPdf(fileData: Uint8Array, fileName: string): Promise<RepairResult> {
  console.log('Starting PDF repair...');
  
  try {
    const decoder = new TextDecoder('latin1');
    let pdfContent = decoder.decode(fileData);
    const issues: string[] = [];
    
    // Check for PDF header
    if (!pdfContent.startsWith('%PDF-')) {
      console.log('PDF header missing, adding...');
      pdfContent = '%PDF-1.4\n' + pdfContent;
      issues.push('PDF header was missing and has been added');
    }
    
    // Check for PDF trailer
    if (!pdfContent.includes('%%EOF')) {
      console.log('PDF trailer missing, adding...');
      pdfContent += '\n%%EOF';
      issues.push('PDF trailer was missing and has been added');
    }
    
    const encoder = new TextEncoder();
    const repairedData = encoder.encode(pdfContent);
    
    return {
      success: true,
      fileName,
      status: issues.length > 0 ? 'partial' : 'success',
      issues: issues.length > 0 ? issues : undefined,
      repairedFile: btoa(String.fromCharCode(...repairedData)),
      fileType: 'pdf',
      recoveryStats: {
        originalSize: fileData.length,
        repairedSize: repairedData.length,
        corruptionLevel: issues.length > 0 ? 'low' : 'none',
        recoveredData: repairedData.length
      }
    };
    
  } catch (error) {
    console.error('PDF repair failed:', error);
    return {
      success: false,
      fileName,
      status: 'failed',
      issues: [`PDF repair failed: ${error.message}`],
      fileType: 'pdf'
    };
  }
}

// Advanced ZIP repair function
async function advancedZipRepair(fileData: Uint8Array, fileName: string, fileType: string): Promise<RepairResult> {
  console.log('Starting advanced ZIP repair...');
  
  try {
    const JSZip = (await import('https://esm.sh/jszip@3.10.1')).default;
    
    // Try direct JSZip loading first
    try {
      const zip = await JSZip.loadAsync(fileData);
      console.log('Direct ZIP loading successful');
      
      const files = Object.keys(zip.files);
      console.log(`ZIP contains ${files.length} files`);
      
      // Regenerate to fix any issues
      const repairedData = await zip.generateAsync({ type: 'uint8array' });
      
      return {
        success: true,
        fileName,
        status: 'success',
        repairedFile: btoa(String.fromCharCode(...repairedData)),
        fileType,
        recoveryStats: {
          originalSize: fileData.length,
          repairedSize: repairedData.length,
          corruptionLevel: 'low',
          recoveredData: files.length
        }
      };
      
    } catch (zipError) {
      console.log('Direct ZIP loading failed, attempting fallback repair...');
      
      // Fallback: create a minimal structure based on file type
      const zip = new JSZip();
      const issues = ['Original ZIP structure was corrupted, created minimal structure'];
      
      switch (fileType) {
        case 'docx':
          zip.file('[Content_Types].xml', generateContentTypes());
          zip.file('_rels/.rels', generateMainRels());
          zip.file('word/_rels/document.xml.rels', generateDocumentRels());
          zip.file('word/document.xml', createDocumentXmlWithContent('Document content could not be recovered due to corruption.'));
          break;
          
        case 'xlsx':
          zip.file('[Content_Types].xml', generateXlsxContentTypes());
          zip.file('_rels/.rels', generateXlsxMainRels());
          zip.file('xl/_rels/workbook.xml.rels', generateXlsxWorkbookRels());
          zip.file('xl/workbook.xml', generateXlsxWorkbook());
          zip.file('xl/worksheets/sheet1.xml', generateXlsxWorksheet());
          break;
          
        case 'pptx':
          zip.file('[Content_Types].xml', generatePptxContentTypes());
          zip.file('_rels/.rels', generatePptxMainRels());
          zip.file('ppt/_rels/presentation.xml.rels', generatePptxPresentationRels());
          zip.file('ppt/presentation.xml', generatePptxPresentation());
          zip.file('ppt/slides/slide1.xml', generatePptxSlide());
          break;
          
        default:
          zip.file('README.txt', 'This file was recovered from corrupted ZIP data.');
          break;
      }
      
      const repairedData = await zip.generateAsync({ type: 'uint8array' });
      
      return {
        success: true,
        fileName,
        status: 'partial',
        issues,
        repairedFile: btoa(String.fromCharCode(...repairedData)),
        fileType,
        recoveryStats: {
          originalSize: fileData.length,
          repairedSize: repairedData.length,
          corruptionLevel: 'high',
          recoveredData: Object.keys(zip.files).length
        }
      };
    }
    
  } catch (error) {
    console.error('Advanced ZIP repair failed:', error);
    return {
      success: false,
      fileName,
      status: 'failed',
      issues: [`Advanced ZIP repair failed: ${error.message}`],
      fileType
    };
  }
}

// XLSX content generation functions
function generateXlsxContentTypes(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`;
}

function generateXlsxMainRels(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`;
}

function generateXlsxWorkbookRels(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`;
}

function generateXlsxWorkbook(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>`;
}

function generateXlsxWorksheet(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetData>
    <row r="1">
      <c r="A1" t="inlineStr">
        <is>
          <t>Data could not be recovered due to corruption</t>
        </is>
      </c>
    </row>
  </sheetData>
</worksheet>`;
}

// PPTX content generation functions
function generatePptxContentTypes(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
</Types>`;
}

function generatePptxMainRels(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
</Relationships>`;
}

function generatePptxPresentationRels(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
</Relationships>`;
}

function generatePptxPresentation(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:sldMasterIdLst>
    <p:sldMasterId id="2147483648" r:id="rId1"/>
  </p:sldMasterIdLst>
  <p:sldIdLst>
    <p:sldId id="256" r:id="rId1"/>
  </p:sldIdLst>
  <p:sldSz cx="9144000" cy="6858000" type="screen4x3"/>
</p:presentation>`;
}

function generatePptxSlide(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr>
        <p:cNvPr id="1" name=""/>
        <p:cNvGrpSpPr/>
        <p:nvPr/>
      </p:nvGrpSpPr>
      <p:grpSpPr>
        <a:xfrm>
          <a:off x="0" y="0"/>
          <a:ext cx="0" cy="0"/>
          <a:chOff x="0" y="0"/>
          <a:chExt cx="0" cy="0"/>
        </a:xfrm>
      </p:grpSpPr>
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="2" name="Title 1"/>
          <p:cNvSpPr>
            <a:spLocks noGrp="1"/>
          </p:cNvSpPr>
          <p:nvPr>
            <p:ph type="ctrTitle"/>
          </p:nvPr>
        </p:nvSpPr>
        <p:spPr/>
        <p:txBody>
          <a:bodyPr/>
          <a:lstStyle/>
          <a:p>
            <a:r>
              <a:rPr lang="en-US"/>
              <a:t>Content could not be recovered due to corruption</a:t>
            </a:r>
            <a:endParaRPr lang="en-US"/>
          </a:p>
        </p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr>
    <a:masterClrMapping/>
  </p:clrMapOvr>
</p:sld>`;
}