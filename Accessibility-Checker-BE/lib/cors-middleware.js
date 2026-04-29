const ALLOWED_ORIGINS = [
  'https://ai-chat-bot-education-2026.vercel.app',
  'https://accessibilitychecker25-arch.github.io',
  'https://kmoreland126.github.io',
  'http://localhost:3000',
  'http://localhost:4200'
];

function getAllowedOrigin(origin) {
  if (origin && ALLOWED_ORIGINS.includes(origin)) {
    return origin;
  }
  return null;
}

function applyCorsHeaders(req, res, options = {}) {
  const allowedMethods = options.allowedMethods || 'GET, POST, OPTIONS';
  const allowedHeaders = options.allowedHeaders || 'Content-Type, Authorization, X-Session-ID';
  const exposeHeaders = options.exposeHeaders || 'Content-Disposition, Content-Type';

  // Allow any origin to access this API. This resolves CORS missing allow origin issues
  // for deployed frontends that may be on different domains or preview URLs.
  res.setHeader('Access-Control-Allow-Origin', '*');

  res.setHeader('Access-Control-Allow-Methods', allowedMethods);
  res.setHeader('Access-Control-Allow-Headers', allowedHeaders);
  res.setHeader('Access-Control-Expose-Headers', exposeHeaders);
  res.setHeader('Access-Control-Max-Age', '86400');
}

function handleCorsPreflight(req, res, options = {}) {
  applyCorsHeaders(req, res, options);
  if (req.method === 'OPTIONS') {
    res.status(200).end();
    return true;
  }
  return false;
}

module.exports = {
  applyCorsHeaders,
  handleCorsPreflight,
};
