import http from 'node:http';
import https from 'node:https';
import { readFile } from 'node:fs/promises';
import { URL } from 'node:url';
import 'dotenv/config';
import mime from 'mime-types';

const PORT = 3001;

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
      console.error(`[Backend-Service] HTTPS ${method} Error:`, err);
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
  const resource = process.env.DATAVERSE_URL;

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

/**
 * Sanitizes filename to ASCII only as per reference
 */
function asciiFilename(filename) {
  return filename.replace(/[^\x00-\x7F]/g, "_");
}

/**
 * Uploads a file to Dataverse via Resilient Multi-Method Logic
 */
async function uploadToDataverse(token, recordId, fieldName, fileName, contentType, buffer) {
  const resource = process.env.DATAVERSE_URL;
  const baseUrl = `${resource}api/data/v9.1/cr8b3_gwia_employee_exceptionses(${recordId})/${fieldName}`;

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
    { method: 'PUT', url: baseUrl, label: 'Base PUT' },
    { method: 'PUT', url: `${baseUrl}/$value`, label: 'Legacy /$value PUT' }
  ];

  for (const attempt of attempts) {
    try {
      process.stdout.write(`    > Trying ${attempt.label}... `);
      const response = await httpsRequest(attempt.method, attempt.url, baseHeaders, buffer);
      
      if (response.ok) {
        console.log(`Success ✅ (${response.status})`);
        return true;
      }
      console.log(`Failed (${response.status})`);
    } catch (err) {
      console.log(`Error ❌ (${err.message})`);
    }
  }

  throw new Error("All upload methods (PATCH, PUT, PUT/$value) were rejected by Dataverse.");
}

const server = http.createServer(async (req, res) => {
  // CORS Headers
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type, x-record-id, x-file-name, x-content-type, x-payload');

  if (req.method === 'OPTIONS') {
    res.writeHead(204);
    res.end();
    return;
  }

  if (req.method === 'POST' && req.url === '/create-and-upload') {
    try {
      const fileName = req.headers['x-file-name'];
      const contentType = "application/octet-stream"; // Header initial
      const payloadStr = req.headers['x-payload'];

      console.log(`\n[STEP 0] Incoming Unified Transaction Request: ${fileName}`);
      
      if (!payloadStr) {
        throw new Error('Step 0 Failed: Missing x-payload metadata header.');
      }

      // Collect binary
      process.stdout.write(`[STEP 1] Collecting binary stream... `);
      const chunks = [];
      for await (const chunk of req) { chunks.push(chunk); }
      const buffer = Buffer.concat(chunks);
      console.log(`Success ✅ (${buffer.length} bytes)`);

      const payload = JSON.parse(decodeURIComponent(payloadStr));

      // 1. Authenticate
      process.stdout.write(`[STEP 2] Authenticating with Azure AD... `);
      const token = await getAccessToken();
      console.log(`Success ✅`);

      // 2. CREATE THE RECORD
      process.stdout.write(`[STEP 3] Creating Dataverse Record... `);
      const createResponse = await fetch(`${process.env.DATAVERSE_URL}api/data/v9.1/cr8b3_gwia_employee_exceptionses`, {
        method: "POST",
        headers: {
          "Authorization": `Bearer ${token}`,
          "Content-Type": "application/json",
          "Prefer": "return=representation"
        },
        body: JSON.stringify(payload)
      });

      if (!createResponse.ok) {
        const err = await createResponse.text();
        console.log(`FAILED ❌`);
        throw new Error(`Step 3 Failed: ${createResponse.status} ${err}`);
      }

      const createdRecord = await createResponse.json();
      const recordId = createdRecord.cr8b3_gwia_employee_exceptionsid;
      console.log(`Success ✅ (ID: ${recordId})`);

      // 3. UPLOAD THE ATTACHMENT
      if (buffer.length > 0) {
        process.stdout.write(`[STEP 4] Storing Binary Attachment... \n`);
        await uploadToDataverse(token, recordId, 'cr8b3_gw_attachments', fileName, contentType, buffer);
        console.log(`    🎉 SUCCESS: Attachment persisted!`);
      } else {
        console.log(`[STEP 4] Skipped (No attachment found)`);
      }

      console.log(`\n🎉 TRANSACTION COMPLETE: Record and attachment persisted successfully!\n`);

      res.writeHead(200, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ 
        success: true, 
        message: 'Transaction finished successfully',
        id: recordId,
        data: createdRecord
      }));
      
    } catch (err) {
      console.error(`\n❌ TRANSACTION FAILED:`, err.message);
      res.writeHead(500, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ error: err.message }));
    }
  } else {
    res.writeHead(404);
    res.end();
  }
});

server.listen(PORT, () => {
  console.log(`\n🚀 [V3-RESILIENT] Dataverse Backend Upload Service running at http://localhost:${PORT}`);
  console.log(`Using Azure App ID: ${process.env.AZURE_CLIENT_ID}\n`);
});
