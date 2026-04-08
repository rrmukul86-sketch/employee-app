import http from 'node:http';
import https from 'node:https';
import { URL } from 'node:url';
import 'dotenv/config';

const PORT = 3002;

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
    console.error(`[Fetch-Service] Azure Auth Failed:`, data);
    throw new Error(`Azure Auth Failed: ${data.error_description || data.error}`);
  }

  return data.access_token;
}

function asciiFilename(filename) {
  return filename.replace(/[^\x00-\x7F]/g, "_");
}

const server = http.createServer(async (req, res) => {
  // CORS Headers
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') {
    res.writeHead(204);
    res.end();
    return;
  }

  const urlObj = new URL(req.url, `http://${req.headers.host}`);
  const pathname = urlObj.pathname;

  if (req.method === 'GET' && pathname === '/proxy') {
    try {
      const path = urlObj.searchParams.get('path');
      console.log(`\n[PROXY] OData Query: ${path}`);
      
      if (!path) throw new Error('Missing path parameter.');

      const token = await getAccessToken();
      const resource = process.env.DATAVERSE_URL.replace(/\/$/, "");
      
      const response = await fetch(`${resource}/${path}`, {
        headers: { 
          "Authorization": `Bearer ${token}`, 
          "OData-Version": "4.0", 
          "OData-MaxVersion": "4.0", 
          "Content-Type": "application/json" 
        }
      });

      const data = await response.json();
      res.writeHead(response.status, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify(data));
      console.log(`    > Relayed ✅`);
    } catch (err) {
      console.error(`\n❌ [PROXY ERROR]:`, err.message);
      res.writeHead(500, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ error: err.message }));
    }
  } 
  
  else if (req.method === 'GET' && (pathname === '/download' || pathname === '/display')) {
    try {
      const mode = pathname === '/download' ? 'DOWNLOAD' : 'DISPLAY';
      const recordId = urlObj.searchParams.get('recordId');
      const fileName = urlObj.searchParams.get('fileName') || 'attachment.file';

      console.log(`\n[${mode}] Media Request for: ${recordId}`);
      if (!recordId) throw new Error('Record GUID is required.');

      const token = await getAccessToken();
      const resource = process.env.DATAVERSE_URL.replace(/\/$/, "");
      const fieldName = 'cr8b3_gw_attachments';
      
      const endpoints = [
        `api/data/v9.1/cr8b3_gwia_employee_exceptionses(${recordId})/${fieldName}/$value`,
        `api/data/v9.1/cr8b3_gwia_employee_exceptions(${recordId})/${fieldName}/$value`
      ];

      for (const endpoint of endpoints) {
        try {
          const response = await fetch(`${resource}/${endpoint}`, {
            headers: { "Authorization": `Bearer ${token}`, "OData-Version": "4.0", "OData-MaxVersion": "4.0" }
          });

          if (response.ok) {
            const buffer = await response.arrayBuffer();
            const contentType = response.headers.get('Content-Type') || 'application/octet-stream';
            console.log(`    > Binary Relayed ✅`);

            const headers = { 'Content-Type': contentType, 'Content-Length': buffer.byteLength };
            if (mode === 'DOWNLOAD') headers['Content-Disposition'] = `attachment; filename="${asciiFilename(fileName)}"`;

            res.writeHead(200, headers);
            res.end(Buffer.from(buffer));
            return;
          }
        } catch (e) {}
      }
      throw new Error("Dataverse file stream not found.");
    } catch (err) {
      console.error(`\n❌ [MEDIA ERROR]:`, err.message);
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
  console.log(`\n🚀 [FETCH-SERVICE] Dataverse Query Gateway running at http://localhost:${PORT}`);
  console.log(`Endpoints: /proxy, /download, /display\n`);
});
