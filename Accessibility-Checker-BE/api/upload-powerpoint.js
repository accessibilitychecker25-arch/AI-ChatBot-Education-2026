const Busboy = require('busboy');

let analyzePowerPoint;
try {
  const pptxAnalyzer = require('../lib/pptx-analyzer');
  analyzePowerPoint = pptxAnalyzer.analyzePowerPoint;
} catch (err) {
  console.error('Failed to load pptx-analyzer:', err);
}

// Helper function to send JSON with proper headers
function sendJson(res, status, data) {
  res.setHeader('Content-Type', 'application/json');
  res.status(status).end(JSON.stringify(data));
}

module.exports = async (req, res) => {
  // Set CORS headers IMMEDIATELY for all requests
  // This is crucial in Vercel serverless environment
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, PUT, DELETE, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type, Authorization');
  res.setHeader('Access-Control-Expose-Headers', 'Content-Disposition, Content-Type');
  res.setHeader('Access-Control-Max-Age', '86400');

  // Handle preflight requests
  if (req.method === 'OPTIONS') {
    return res.status(200).end();
  }

  if (req.method !== 'POST') {
    sendJson(res, 405, { error: 'Method not allowed' });
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
        sendJson(res, 400, { error: 'No file uploaded' });
        return;
      }

      // Validate PowerPoint file types
      const validExtensions = ['.pptx', '.ppt', '.pps', '.potx'];
      const isValid = validExtensions.some(ext => filename.toLowerCase().endsWith(ext));
      
      if (!isValid) {
        sendJson(res, 400, { error: 'Please upload a PowerPoint file (.pptx, .ppt, .pps, or .potx)' });
        return;
      }

      try {
        if (!analyzePowerPoint) {
          throw new Error('PowerPoint analyzer not available');
        }
        const report = await analyzePowerPoint(fileData, filename);
        sendJson(res, 200, {
          fileName: filename,
          suggestedFileName: filename,
          report: report
        });
      } catch (error) {
        console.error('PowerPoint analysis error:', error);
        sendJson(res, 500, { error: error.message });
      }
    });

    req.pipe(busboy);

  } catch (error) {
    console.error('Upload error:', error);
    sendJson(res, 500, { error: error.message });
  }
};
