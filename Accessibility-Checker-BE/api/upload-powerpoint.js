const Busboy = require('busboy');
const JSZip = require('jszip');

module.exports = async (req, res) => {
  // CORS configuration
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

      // Validate PowerPoint file types
      const validExtensions = ['.pptx', '.ppt', '.pps', '.potx'];
      const isValid = validExtensions.some(ext => filename.toLowerCase().endsWith(ext));
      
      if (!isValid) {
        res.status(400).json({ 
          error: 'Please upload a PowerPoint file (.pptx, .ppt, .pps, or .potx)' 
        });
        return;
      }

      try {
        const report = await analyzePowerPoint(fileData, filename);
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

// Main PowerPoint analysis function
async function analyzePowerPoint(fileData, filename) {
  const report = {
    fileName: filename,
    suggestedFileName: filename,
    summary: { fixed: 0, flagged: 0 },
    details: {
      listFormattingIssues: [],
      imagesMissingOrBadAlt: [],
    }
  };

  try {
    const zip = await JSZip.loadAsync(fileData);
    
    // Get list of slides
    const slides = [];
    zip.forEach((relativePath, file) => {
      if (relativePath.match(/^ppt\/slides\/slide\d+\.xml$/)) {
        slides.push(relativePath);
      }
    });
    
    // Sort slides by number
    slides.sort((a, b) => {
      const numA = parseInt(a.match(/slide(\d+)\.xml$/)?.[1] || '0');
      const numB = parseInt(b.match(/slide(\d+)\.xml$/)?.[1] || '0');
      return numA - numB;
    });
    
    console.log(`[analyzePowerPoint] Found ${slides.length} slides`);
    
    // Analyze each slide
    for (let i = 0; i < slides.length; i++) {
      const slidePath = slides[i];
      const slideNumber = i + 1;
      const slideXml = await zip.file(slidePath)?.async('string');
      const slideRelsPath = slidePath.replace('ppt/slides/', 'ppt/slides/_rels/').replace('.xml', '.xml.rels');
      const slideRels = await zip.file(slideRelsPath)?.async('string');
      
      if (slideXml) {
        // Check for list formatting issues (hyphenated paragraphs)
        const listIssues = checkListFormatting(slideXml, slideNumber);
        if (listIssues.length > 0) {
          report.details.listFormattingIssues.push(...listIssues);
          report.summary.flagged += listIssues.length;
        }
        
        // Check images for alt text
        const imageIssues = await analyzeSlideImages(slideXml, slideRels, slideNumber);
        if (imageIssues.length > 0) {
          report.details.imagesMissingOrBadAlt.push(...imageIssues);
          report.summary.flagged += imageIssues.length;
        }
      }
    }
    
    console.log(`[analyzePowerPoint] Analysis complete. Fixed: ${report.summary.fixed}, Flagged: ${report.summary.flagged}`);
    return report;
    
  } catch (error) {
    console.error('[analyzePowerPoint] Error:', error);
    throw new Error(`Failed to analyze PowerPoint: ${error.message}`);
  }
}

// Check for list formatting issues (hyphenated paragraphs that should be lists)
function checkListFormatting(slideXml, slideNumber) {
  const issues = [];
  
  // Find all text elements in the slide
  const textMatches = slideXml.matchAll(/<a:t[^>]*>(.*?)<\/a:t>/g);
  
  for (const match of textMatches) {
    const text = match[1];
    
    // Check for hyphenated paragraphs that look like lists
    // Pattern: line starting with "-", "•", "–", "—" followed by text
    if (/^[\s]*[-–—•]\s+.+/.test(text)) {
      issues.push({
        slideNumber: slideNumber,
        location: `Slide ${slideNumber}`,
        issue: `Possible improperly formatted list: "${text.substring(0, 50)}..."`,
        type: 'listFormatting'
      });
    }
  }
  
  return issues;
}

// Analyze images in a slide
async function analyzeSlideImages(slideXml, slideRels, slideNumber) {
  const issues = [];
  
  // Find all picture elements
  const picMatches = slideXml.matchAll(/<p:pic[\s\S]*?<\/p:pic>/g);
  
  for (const picMatch of picMatches) {
    const picXml = picMatch[0];
    
    // Check for alt text (descr attribute in <p:cNvPr>)
    const nvPicPr = picXml.match(/<p:nvPicPr>([\s\S]*?)<\/p:nvPicPr>/);
    if (nvPicPr) {
      const cNvPr = nvPicPr[1].match(/<p:cNvPr[^>]*>/);
      if (cNvPr) {
        const descrMatch = cNvPr[0].match(/descr="([^"]*)"/);
        const altText = descrMatch ? descrMatch[1] : '';
        
        if (!altText || altText.trim().length === 0) {
          issues.push({
            slideNumber: slideNumber,
            location: `Slide ${slideNumber}`,
            issue: 'Image missing alt text',
            type: 'image'
          });
        } else if (altText.length > 250) {
          issues.push({
            slideNumber: slideNumber,
            location: `Slide ${slideNumber}`,
            issue: `Image alt text is too long (${altText.length} characters, max 250)`,
            type: 'image'
          });
        }
      }
    }
  }
  
  return issues;
}

module.exports.analyzePowerPoint = analyzePowerPoint;
