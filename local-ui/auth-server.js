const http = require('http');
const https = require('https');
const { execSync } = require('child_process');
const fs = require('fs');
const path = require('path');
const { URL } = require('url');

const DATAVERSE_URL = String(process.env.DATAVERSE_URL || '').trim().replace(/\/$/, '');
const PORT = Number(process.env.PORT || 3001);
const RESOURCE = DATAVERSE_URL;
const STATIC_ROOT = __dirname;
const DATAVERSE_API_PREFIX = '/api/data/v9.2/';

let token = null;
let expiresAt = null;

const isConfigured = () => Boolean(DATAVERSE_URL);

const getClientConfig = () => ({
  dataverseUrl: DATAVERSE_URL || 'Not configured',
  authMode: 'Azure CLI',
  isConfigured: isConfigured(),
  dataverseApiPrefix: DATAVERSE_API_PREFIX
});

const getContentType = (filePath) => {
  const extension = path.extname(filePath).toLowerCase();
  if (extension === '.html') return 'text/html; charset=utf-8';
  if (extension === '.css') return 'text/css; charset=utf-8';
  if (extension === '.js') return 'application/javascript; charset=utf-8';
  if (extension === '.json') return 'application/json; charset=utf-8';
  return 'application/octet-stream';
};

const getToken = () => {
  if (!isConfigured()) {
    throw new Error('Set DATAVERSE_URL before starting the local UI server.');
  }

  if (token && Date.now() < expiresAt - 300000) {
    return token;
  }

  const result = JSON.parse(execSync(
    `az account get-access-token --resource "${RESOURCE}" --query "{accessToken: accessToken, expiresOn: expiresOn}" -o json`
  ).toString());

  token = result.accessToken;
  expiresAt = new Date(result.expiresOn).getTime();
  return token;
};

const readRequestBody = (req) => new Promise((resolve, reject) => {
  const chunks = [];
  req.on('data', (chunk) => chunks.push(chunk));
  req.on('end', () => resolve(Buffer.concat(chunks)));
  req.on('error', reject);
});

const proxyDataverseRequest = async (req, res) => {
  try {
    if (!isConfigured()) {
      res.writeHead(400, { 'Content-Type': 'text/plain; charset=utf-8' });
      res.end('Set DATAVERSE_URL before proxying Dataverse requests.');
      return;
    }

    const body = await readRequestBody(req);
    const upstreamUrl = new URL(req.url, DATAVERSE_URL);
    const headers = {
      Authorization: `Bearer ${getToken()}`,
      Accept: req.headers.accept || 'application/json',
      'OData-Version': req.headers['odata-version'] || '4.0',
      'OData-MaxVersion': req.headers['odata-maxversion'] || '4.0'
    };

    if (req.headers['content-type']) {
      headers['Content-Type'] = req.headers['content-type'];
    }

    if (body.length > 0) {
      headers['Content-Length'] = body.length;
    }

    const upstreamRequest = https.request(upstreamUrl, {
      method: req.method,
      headers
    }, (upstreamResponse) => {
      const responseChunks = [];
      upstreamResponse.on('data', (chunk) => responseChunks.push(chunk));
      upstreamResponse.on('end', () => {
        const responseBody = Buffer.concat(responseChunks);
        const responseHeaders = {};

        if (upstreamResponse.headers['content-type']) {
          responseHeaders['Content-Type'] = upstreamResponse.headers['content-type'];
        }

        res.writeHead(upstreamResponse.statusCode || 500, responseHeaders);
        res.end(responseBody);
      });
    });

    upstreamRequest.on('error', (error) => {
      res.writeHead(502, { 'Content-Type': 'text/plain; charset=utf-8' });
      res.end(error.message);
    });

    if (body.length > 0) {
      upstreamRequest.write(body);
    }

    upstreamRequest.end();
  } catch (error) {
    res.writeHead(500, { 'Content-Type': 'text/plain; charset=utf-8' });
    res.end(error.message);
  }
};

const serveFile = (requestPath, res) => {
  const safePath = requestPath === '/' ? '/index.html' : requestPath;
  const resolvedPath = path.normalize(path.join(STATIC_ROOT, safePath));

  if (!resolvedPath.startsWith(STATIC_ROOT)) {
    res.writeHead(403);
    res.end('Forbidden');
    return;
  }

  fs.readFile(resolvedPath, (error, buffer) => {
    if (error) {
      res.writeHead(error.code === 'ENOENT' ? 404 : 500, { 'Content-Type': 'text/plain; charset=utf-8' });
      res.end(error.code === 'ENOENT' ? 'Not found' : 'Server error');
      return;
    }

    res.writeHead(200, { 'Content-Type': getContentType(resolvedPath) });
    res.end(buffer);
  });
};

http.createServer((req, res) => {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET,POST,OPTIONS');

  if (req.method === 'OPTIONS') {
    res.writeHead(204);
    res.end();
    return;
  }

  if (req.url === '/token') {
    try {
      res.writeHead(200, { 'Content-Type': 'text/plain; charset=utf-8' });
      res.end(getToken());
    } catch (error) {
      res.writeHead(500, { 'Content-Type': 'text/plain; charset=utf-8' });
      res.end('Error: Run "az login --allow-no-subscriptions"');
    }
    return;
  }

  if (req.url === '/config') {
    res.writeHead(200, { 'Content-Type': 'application/json; charset=utf-8' });
    res.end(JSON.stringify(getClientConfig()));
    return;
  }

  if (req.url && req.url.startsWith(DATAVERSE_API_PREFIX)) {
    proxyDataverseRequest(req, res);
    return;
  }

  serveFile(req.url || '/', res);
}).listen(PORT, () => {
  console.log(`Local UI: http://localhost:${PORT}`);
  if (!isConfigured()) {
    console.log('Set DATAVERSE_URL to your Dataverse environment URL before using the proxy endpoints.');
  }
});