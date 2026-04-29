const sessionManager = require('../lib/session-manager');
const { applyCorsHeaders, handleCorsPreflight } = require('../lib/cors-middleware');

module.exports = async (req, res) => {
  if (handleCorsPreflight(req, res, { allowedMethods: 'POST, GET, OPTIONS', allowedHeaders: 'Content-Type, Authorization, X-Session-ID' })) {
    return;
  }
  applyCorsHeaders(req, res, { allowedMethods: 'POST, GET, OPTIONS', allowedHeaders: 'Content-Type, Authorization, X-Session-ID' });

  try {
    const sessionId = req.headers['x-session-id'] || req.query.sessionId || req.body?.sessionId;

    switch (req.method) {
      case 'POST':
        // Heartbeat - keep session alive
        if (sessionId && sessionManager.heartbeat(sessionId)) {
          res.json({ 
            success: true, 
            sessionId: sessionId,
            message: 'Session refreshed' 
          });
        } else {
          // Create new session if doesn't exist
          const newSession = sessionManager.getOrCreateSession(null);
          res.json({ 
            success: true, 
            sessionId: newSession.sessionId,
            message: 'New session created' 
          });
        }
        break;

      case 'GET':
        if (req.query.action === 'stats') {
          // Get session statistics (for debugging)
          const stats = sessionManager.getSessionStats();
          res.json(stats);
        } else if (sessionId) {
          // Get session info
          const session = sessionManager.getOrCreateSession(sessionId);
          res.json({
            sessionId: session.sessionId,
            createdAt: session.createdAt,
            lastActivity: session.lastActivity,
            files: sessionManager.getSessionFiles(sessionId),
            batches: sessionManager.getSessionBatches(sessionId),
            expiresIn: '1 hour from last activity'
          });
        } else {
          res.status(400).json({ error: 'sessionId required' });
        }
        break;

      default:
        res.status(405).json({ error: 'Method not allowed' });
    }
  } catch (error) {
    console.error('Session API error:', error);
    res.status(500).json({ error: 'Internal server error' });
  }
};