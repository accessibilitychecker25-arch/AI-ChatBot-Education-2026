const Busboy = require('busboy');
const JSZip = require('jszip');

// Helper function to extract text from paragraph XML - moved to top for availability
function extractTextFromParagraph(paragraphXml) {
  const textMatches = paragraphXml.match(/<w:t[^>]*>(.*?)<\/w:t>/g);
  if (!textMatches) return '';
  
  return textMatches
    .map(t => t.replace(/<w:t[^>]*>|<\/w:t>/g, ''))
    .join('')
    .trim();
}

module.exports = async (req, res) => {
  // CORS: safe allowlist — echo back the requesting Origin when allowed.
  const ALLOWED_ORIGINS = [
    'https://accessibilitychecker25-arch.github.io',
    'https://kmoreland126.github.io',
    'http://localhost:3000',
    'http://localhost:4200'
  ];
  const origin = req.headers.origin;
  if (origin && ALLOWED_ORIGINS.includes(origin)) {
    res.setHeader('Access-Control-Allow-Origin', origin);
  }
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  // Expose Content-Disposition for downloads and Content-Type for clients
  res.setHeader('Access-Control-Expose-Headers', 'Content-Disposition, Content-Type');

  if (req.method === 'OPTIONS') {
    res.status(200).end();
    return;
  }

  if (req.method !== 'POST') {
    res.status(405).json({ error: 'Method not allowed' });
    return;
  }

  try {
    const busboy = Busboy({ headers: req.headers });
    let fileData = null;
    let filename = null;

    busboy.on('file', (fieldname, file, info) => {
      filename = info.filename;
      const chunks = [];
      
      file.on('data', (chunk) => {
        chunks.push(chunk);
      });
      
      file.on('end', () => {
        fileData = Buffer.concat(chunks);
      });
    });

    busboy.on('finish', async () => {
      if (!fileData || !filename) {
        res.status(400).json({ error: 'No file uploaded' });
        return;
      }

      if (!filename.toLowerCase().endsWith('.docx')) {
        res.status(400).json({ error: 'Please upload a .docx file' });
        return;
      }

      try {
        const report = await analyzeDocx(fileData, filename);
        res.status(200).json({
          fileName: filename,
          suggestedFileName: filename,
          report: report
        });
      } catch (error) {
        res.status(500).json({ error: error.message });
      }
    });

    req.pipe(busboy);

  } catch (error) {
    res.status(500).json({ error: error.message });
  }
};
module.exports.analyzeDocx = analyzeDocx;
async function analyzeDocx(fileData, filename) {
  const report = {
    fileName: filename,
    suggestedFileName: filename,
    summary: { fixed: 0, flagged: 0 },
    details: {
      // Requirement 1: Lists are formatted correctly
      hyphenatedParagraphsNeedingLists: [],
      formattedListsCount: 0,
      
      // Requirement 2: Images have alt text (max 250 chars)
      imagesMissingAltText: [],
      imagesWithAltTextOver250Chars: [],
      imagesWithValidAltText: 0,
    }
  };

  try {
    const zip = await JSZip.loadAsync(fileData);
    
    // Read core documents needed for the two requirements
    const documentXml = await zip.file('word/document.xml')?.async('string');
    const relsXml = await zip.file('word/_rels/document.xml.rels')?.async('string');
    
    // ===== REQUIREMENT 1: Check for lists formatted correctly =====
    if (documentXml) {
      const listIssues = analyzeListFormatting(documentXml);
      if (listIssues.hyphenatedParagraphs.length > 0) {
        report.details.hyphenatedParagraphsNeedingLists = listIssues.hyphenatedParagraphs;
        report.summary.flagged += listIssues.hyphenatedParagraphs.length;
      }
      report.details.formattedListsCount = listIssues.properlyFormattedLists;
    }
    
    // ===== REQUIREMENT 2: Check for images with alt text =====
    if (relsXml && documentXml) {
      const imageAnalysis = analyzeImageAltText(documentXml, relsXml);
      
      if (imageAnalysis.missingAltText.length > 0) {
        report.details.imagesMissingAltText = imageAnalysis.missingAltText;
        report.summary.flagged += imageAnalysis.missingAltText.length;
      }
      
      if (imageAnalysis.altTextOver250Chars.length > 0) {
        report.details.imagesWithAltTextOver250Chars = imageAnalysis.altTextOver250Chars;
        report.summary.flagged += imageAnalysis.altTextOver250Chars.length;
      }
      
      report.details.imagesWithValidAltText = imageAnalysis.validAltTextCount;
    }
    
    return report;
    
  } catch (error) {
    console.error('[analyzeDocx] Error analyzing document:', error);
    return {
      fileName: filename,
      error: error.message,
      summary: { fixed: 0, flagged: 0 },
      details: {}
    };
  }
}

// ===== HELPER FUNCTIONS =====

/**
 * Analyze list formatting in the document
 * Detects hyphenated paragraphs that should be formatted as lists
 */
function analyzeListFormatting(documentXml) {
  const results = {
    hyphenatedParagraphs: [],
    properlyFormattedLists: 0
  };

  if (!documentXml) return results;

  // Extract all paragraphs
  const paragraphMatches = documentXml.match(/<w:p[^>]*>([\s\S]*?)<\/w:p>/g) || [];
  
  paragraphMatches.forEach((paragraph, index) => {
    // Extract text content from paragraph
    const textMatches = paragraph.match(/<w:t[^>]*>(.*?)<\/w:t>/g) || [];
    const text = textMatches
      .map(t => t.replace(/<w:t[^>]*>|<\/w:t>/g, ''))
      .join('')
      .trim();

    // Check if paragraph starts with hyphen/dash (indicates list formatting issue)
    if (text && /^[-–—]\s+/.test(text)) {
      results.hyphenatedParagraphs.push({
        index: index + 1,
        text: text.substring(0, 100), // First 100 chars
        message: 'This paragraph appears to be a list item but is formatted as a regular paragraph'
      });
    }
    
    // Count properly formatted lists (pPr contains pStyle with list references)
    if (paragraph.includes('pStyle w:val="ListParagraph"') || paragraph.includes('numPr')) {
      results.properlyFormattedLists++;
    }
  });

  return results;
}

/**
 * Analyze image alt text requirements
 * Checks for missing alt text and validates length
 */
function analyzeImageAltText(documentXml, relsXml) {
  const results = {
    missingAltText: [],
    altTextOver250Chars: [],
    validAltTextCount: 0
  };

  if (!documentXml || !relsXml) return results;

  // Find all images/drawings
  const drawingMatches = documentXml.match(/<wp:inline[^>]*>[\s\S]*?<\/wp:inline>|<wp:anchor[^>]*>[\s\S]*?<\/wp:anchor>/g) || [];
  
  drawingMatches.forEach((drawing, index) => {
    // Extract relationship ID to find the image file
    const rIdMatch = drawing.match(/r:embed="(rId\d+)"/);
    if (!rIdMatch) return;

    const rId = rIdMatch[1];

    // Extract alternate text (docProperties)
    const altTextMatch = drawing.match(/<wp:docPr[^>]*descr="([^"]*)"/) || drawing.match(/<wp:cNvPicPr[^>]*>[\s\S]*?<a:picLocks[^>]*descr="([^"]*)"/);
    const altText = altTextMatch ? altTextMatch[1] : null;

    // Also check for extent/alt description in other formats
    const titleMatch = drawing.match(/<wp:docPr[^>]*name="([^"]*)"[^>]*title="([^"]*)"/) || drawing.match(/<wp:docPr[^>]*title="([^"]*)"[^>]*name="([^"]*)"/);

    // Check if this image has proper alt text
    if (!altText || altText.trim() === '') {
      results.missingAltText.push({
        index: index + 1,
        rId: rId,
        message: 'Image is missing alt text description'
      });
    } else if (altText.length > 250) {
      results.altTextOver250Chars.push({
        index: index + 1,
        rId: rId,
        altText: altText.substring(0, 100) + '...',
        length: altText.length,
        message: `Alt text is ${altText.length} characters (max 250)`
      });
    } else {
      // Valid alt text
      results.validAltTextCount++;
    }
  });

  return results;
}

