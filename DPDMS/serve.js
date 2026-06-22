const http = require('http');
const fs = require('fs');
const path = require('path');
const { exec } = require('child_process');

const PORT = 8080;

const server = http.createServer((req, res) => {
  const urlObj = new URL(req.url, `http://${req.headers.host || 'localhost'}`);
  const urlPath = urlObj.pathname;
  
  if (urlPath.startsWith('/api/')) {
    const importer = require('../work/import_excel_node.js');

    // 1. Raw binary file upload
    if (urlPath === '/api/upload' && req.method === 'POST') {
      let chunks = [];
      req.on('data', chunk => chunks.push(chunk));
      req.on('end', async () => {
        try {
          const buffer = Buffer.concat(chunks);
          const tempDir = path.join(__dirname, '..', 'scratch');
          if (!fs.existsSync(tempDir)) {
            fs.mkdirSync(tempDir, { recursive: true });
          }
          const tempPath = path.join(tempDir, `upload_${Date.now()}.xlsx`);
          fs.writeFileSync(tempPath, buffer);

          console.log(`Uploaded file saved to ${tempPath}. Invoking merge and sync...`);
          const result = await importer.handleMergeAndSync(tempPath);

          try {
            fs.unlinkSync(tempPath);
          } catch (e) {}

          res.writeHead(200, { 'Content-Type': 'application/json' });
          res.end(JSON.stringify(result));
        } catch (error) {
          console.error('API upload error:', error);
          res.writeHead(500, { 'Content-Type': 'application/json' });
          res.end(JSON.stringify({ ok: false, error: error.message }));
        }
      });
      return;
    }

    // 2. JSON Body and Query Parser for other API routes
    let bodyData = '';
    req.on('data', chunk => bodyData += chunk);
    req.on('end', async () => {
      try {
        let payload = {};
        if (bodyData && req.headers['content-type'] && req.headers['content-type'].includes('application/json')) {
          payload = JSON.parse(bodyData);
        }

        // Parse query params
        const query = {};
        for (const [key, value] of urlObj.searchParams.entries()) {
          query[key] = value;
        }

        let result = null;

        if (urlPath === '/api/transfer-request') {
          if (req.method === 'POST') {
            result = await importer.saveTransferRequest(payload);
          } else if (req.method === 'DELETE') {
            const kgid = query.kgid;
            result = await importer.deleteTransferRequest(kgid);
          }
        } else if (urlPath === '/api/employee') {
          if (req.method === 'POST') {
            result = await importer.saveEmployee(payload);
          } else if (req.method === 'DELETE') {
            const kgid = query.kgid;
            result = await importer.deleteEmployee(kgid);
          }
        } else if (urlPath === '/api/check-duplicate' && req.method === 'GET') {
          result = await importer.checkDuplicate(query);
        } else if (urlPath === '/api/data' && req.method === 'GET') {
          const outputsDir = path.join(__dirname, '..', 'outputs');
          const masterPath = path.join(outputsDir, 'District_Employee_Transfer_Database_IMPORTED.xlsx');
          const fallbackTemplatePath = path.join(outputsDir, 'District_Employee_Transfer_Database_Template_FIXED.xlsx');
          const activeMasterPath = fs.existsSync(masterPath) ? masterPath : fallbackTemplatePath;

          result = await importer.buildStandaloneJSON(activeMasterPath);
        }

        if (result !== null) {
          res.writeHead(200, { 'Content-Type': 'application/json' });
          res.end(JSON.stringify(result));
        } else {
          res.writeHead(404, { 'Content-Type': 'application/json' });
          res.end(JSON.stringify({ error: `Method ${req.method} for ${urlPath} not supported` }));
        }
      } catch (error) {
        console.error(`API Error on ${req.method} ${urlPath}:`, error);
        res.writeHead(500, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ error: error.message }));
      }
    });
    return;
  }

  let filePath = path.join(__dirname, req.url === '/' ? 'index.html' : req.url.split('?')[0]);
  
  fs.stat(filePath, (err, stats) => {
    if (err || !stats.isFile()) {
      filePath = path.join(__dirname, 'index.html'); // fallback to index.html for SPA router
    }
    
    fs.readFile(filePath, (err, content) => {
      if (err) {
        res.writeHead(500);
        res.end('Error loading file');
      } else {
        let contentType = 'text/html';
        if (filePath.endsWith('.js')) contentType = 'text/javascript';
        else if (filePath.endsWith('.css')) contentType = 'text/css';
        else if (filePath.endsWith('.json')) contentType = 'application/json';
        else if (filePath.endsWith('.png')) contentType = 'image/png';
        
        res.writeHead(200, { 'Content-Type': contentType });
        res.end(content);
      }
    });
  });
});

server.listen(PORT, () => {
  const url = `http://localhost:${PORT}`;
  console.log(`Server running at ${url}`);
  exec(`start ${url}`);
});
