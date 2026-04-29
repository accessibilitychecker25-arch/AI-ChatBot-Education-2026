const { applyCorsHeaders, handleCorsPreflight } = require('../lib/cors-middleware');

module.exports = async (req, res) => {
  if (handleCorsPreflight(req, res, { allowedMethods: 'GET, POST, PUT, DELETE, OPTIONS' })) {
    return;
  }
  applyCorsHeaders(req, res, { allowedMethods: 'GET, POST, PUT, DELETE, OPTIONS' });

  res.setHeader('Content-Type', 'application/json');
  
  if (req.method === 'OPTIONS') {
    return res.status(200).end();
  }
  
  return res.status(200).end(JSON.stringify({ ok: true }));
};
