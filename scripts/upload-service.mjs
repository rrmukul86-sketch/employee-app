import http from 'node:http';
import https from 'node:https';
import { URL } from 'node:url';
import 'dotenv/config';

const PORT = process.env.PORT || 3001;

/**
 * Perform a request using the native https module to ensure headers are preserved exactly
 */
function httpsRequest(method, url, headers, body) {
  return new Promise((resolve, reject) => {
    const parsedUrl = new URL(url);
    const options = {
      hostname: parsedUrl.hostname,
      port: 443,
      path: parsedUrl.pathname + parsedUrl.search,
      method: method,
      headers: headers
    };

    const req = https.request(options, (res) => {
      let responseData = '';
      res.on('data', (chunk) => { responseData += chunk; });
      res.on('end', () => {
        resolve({ 
          ok: res.statusCode >= 200 && res.statusCode < 300, 
          status: res.statusCode, 
          text: () => Promise.resolve(responseData) 
        });
      });
    });

    req.on('error', (err) => {
      console.error(`[Backend-Service] HTTPS Error:`, err);
      reject(err);
    });
    if (body) req.write(body);
    req.end();
  });
}

/**
 * Gets an Access Token from Azure AD using Client Credentials flow
 */
async function getAccessToken() {
  const tenantId = process.env.AZURE_TENANT_ID;
  const clientId = process.env.AZURE_CLIENT_ID;
  const clientSecret = process.env.AZURE_CLIENT_SECRET;
  const resource = process.env.DATAVERSE_URL.replace(/\/$/, "");

  const url = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;
  const body = new URLSearchParams({
    client_id: clientId,
    client_secret: clientSecret,
    grant_type: 'client_credentials',
    scope: `${resource}/.default`
  });

  const response = await fetch(url, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: body.toString()
  });

  const data = await response.json();
  if (!response.ok) {
    console.error(`[Backend-Service] Azure Auth Failed:`, data);
    throw new Error(`Azure Auth Failed: ${data.error_description || data.error}`);
  }

  return data.access_token;
}

function asciiFilename(filename) {
  return filename.replace(/[^\x00-\x7F]/g, "_");
}

async function uploadToDataverse(token, recordId, fieldName, fileName, contentType, buffer) {
  const resource = process.env.DATAVERSE_URL.replace(/\/$/, "");
  const baseUrl = `${resource}/api/data/v9.1/cr8b3_gwia_employee_exceptionses(${recordId})/${fieldName}`;

  const baseHeaders = {
    "Authorization": `Bearer ${token}`,
    "Content-Type": "application/octet-stream",
    "x-ms-file-name": asciiFilename(fileName),
    "OData-Version": "4.0",
    "OData-MaxVersion": "4.0",
    "Prefer": "return=minimal",
    "Content-Length": buffer.length
  };

  const attempts = [
    { method: 'PATCH', url: baseUrl, label: 'Standard PATCH' },
    { method: 'PUT', url: `${baseUrl}/$value`, label: 'Legacy /$value PUT' }
  ];

  for (const attempt of attempts) {
    try {
      console.log(`    > Trying ${attempt.label}...`);
      const response = await httpsRequest(attempt.method, attempt.url, baseHeaders, buffer);
      if (response.ok) return true;
    } catch (err) {}
  }
  throw new Error("All upload methods were rejected.");
}

const server = http.createServer(async (req, res) => {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type, x-record-id, x-file-name, x-payload');

  if (req.method === 'OPTIONS') {
    res.writeHead(204);
    res.end();
    return;
  }

  if (req.method === 'POST' && req.url === '/create-and-upload') {
    try {
      const fileName = req.headers['x-file-name'] || 'attachment';
      const payloadStr = req.headers['x-payload'];
      const chunks = [];
      for await (const chunk of req) chunks.push(chunk);
      const buffer = Buffer.concat(chunks);
      const payload = JSON.parse(decodeURIComponent(payloadStr));

      const token = await getAccessToken();
      const resource = process.env.DATAVERSE_URL.replace(/\/$/, "");
      
      const createResponse = await fetch(`${resource}/api/data/v9.1/cr8b3_gwia_employee_exceptionses`, {
        method: "POST",
        headers: { "Authorization": `Bearer ${token}`, "Content-Type": "application/json", "Prefer": "return=representation" },
        body: JSON.stringify(payload)
      });
      if (!createResponse.ok) throw new Error("Record creation failed.");
      
      const createdRecord = await createResponse.json();
      const recordId = createdRecord.cr8b3_gwia_employee_exceptionsid;
      
      if (buffer.length > 0) {
        await uploadToDataverse(token, recordId, 'cr8b3_gw_attachments', fileName, 'application/octet-stream', buffer);
      }

      res.writeHead(200, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ success: true, id: recordId, data: createdRecord }));
    } catch (err) {
      res.writeHead(500, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ error: err.message }));
    }
  } 
  
  else if (req.method === 'GET' && (req.url.startsWith('/download') || req.url.startsWith('/display'))) {
    try {
      const mode = req.url.startsWith('/download') ? 'DOWNLOAD' : 'DISPLAY';
      const urlParams = new URL(req.url, `http://${req.headers.host}`).searchParams;
      const recordId = urlParams.get('recordId');
      const fileName = urlParams.get('fileName') || 'attachment.file';
      
      const token = await getAccessToken();
      const resource = process.env.DATAVERSE_URL.replace(/\/$/, "");
      const fieldName = 'cr8b3_gw_attachments';
      const endpoints = [
        `api/data/v9.1/cr8b3_gwia_employee_exceptionses(${recordId})/${fieldName}/$value`,
        `api/data/v9.1/cr8b3_gwia_employee_exceptions(${recordId})/${fieldName}/$value`
      ];

      for (const endpoint of endpoints) {
        const response = await fetch(`${resource}/${endpoint}`, {
          headers: { "Authorization": `Bearer ${token}`, "OData-Version": "4.0", "OData-MaxVersion": "4.0" }
        });
        if (response.ok) {
          const buffer = await response.arrayBuffer();
          let contentType = response.headers.get('Content-Type') || 'application/octet-stream';
          
          if (contentType === 'application/octet-stream') {
             const ext = fileName.split('.').pop().toLowerCase();
             const map = { 'png': 'image/png', 'jpg': 'image/jpeg', 'jpeg': 'image/jpeg', 'pdf': 'application/pdf' };
             contentType = map[ext] || contentType;
          }

          const headers = { 'Content-Type': contentType, 'Content-Length': buffer.byteLength };
          if (mode === 'DOWNLOAD') headers['Content-Disposition'] = `attachment; filename="${asciiFilename(fileName)}"`;
          res.writeHead(200, headers);
          res.end(Buffer.from(buffer));
          return;
        }
      }
      throw new Error("File not found.");
    } catch (err) {
      res.writeHead(500, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ error: err.message }));
    }
  }
  
  else {
    res.writeHead(404);
    res.end();
  }
});

server.listen(PORT, () => {
  console.log(`\n🚀 [UNIFIED GATEWAY] Dataverse Service running at http://localhost:${PORT}`);
});
