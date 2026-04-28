const Busboy = require('busboy');
const { analyzePowerPoint } = require('../lib/pptx-analyzer');

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
