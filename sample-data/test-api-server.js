/**
 * Lightweight local REST API server for testing the REST API data source.
 *
 * Run with: node sample-data/test-api-server.js
 * Then in the web part, use URL: http://localhost:3001/data
 * and Data Path: value
 *
 * NOTE: This is for local Workbench testing only.
 * You may need to disable CORS restrictions in your browser or
 * run the SPFx workbench without SSL for this to work.
 */
const http = require('http');
const fs = require('fs');
const path = require('path');

const PORT = 3001;

const data = JSON.parse(
  fs.readFileSync(path.join(__dirname, 'api-response-example.json'), 'utf8')
);

const server = http.createServer((req, res) => {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type, Accept');

  if (req.method === 'OPTIONS') {
    res.writeHead(204);
    res.end();
    return;
  }

  if (req.url === '/data' || req.url === '/') {
    res.writeHead(200, { 'Content-Type': 'application/json' });
    res.end(JSON.stringify(data));
  } else {
    res.writeHead(404, { 'Content-Type': 'application/json' });
    res.end(JSON.stringify({ error: 'Not found. Use /data' }));
  }
});

server.listen(PORT, () => {
  console.log(`Test API server running at http://localhost:${PORT}`);
  console.log(`Data endpoint: http://localhost:${PORT}/data`);
  console.log('Use Data Path: value');
  console.log('\nPress Ctrl+C to stop.');
});
