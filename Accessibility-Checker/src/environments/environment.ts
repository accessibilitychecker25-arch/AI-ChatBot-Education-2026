export const environment = {
  production: false,
  // 1. Point to your local Python backend
  apiUrl: 'http://localhost:5000', 
  // 2. Point to the specific route we created in server.py
  uploadEndpoint: '/upload',
  downloadEndpoint: '/download', 
};