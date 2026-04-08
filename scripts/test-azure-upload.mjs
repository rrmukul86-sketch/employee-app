import 'dotenv/config';

async function testAzureApp() {
  console.log("--- Azure App Storage Test ---");
  console.log(`Tenant: ${process.env.AZURE_TENANT_ID}`);
  console.log(`Client: ${process.env.AZURE_CLIENT_ID}`);
  console.log(`URL:    ${process.env.DATAVERSE_URL}`);

  try {
    // 1. Get Token
    console.log("\n1. Requesting Access Token...");
    const authUrl = `https://login.microsoftonline.com/${process.env.AZURE_TENANT_ID}/oauth2/v2.0/token`;
    const authBody = new URLSearchParams({
      client_id: process.env.AZURE_CLIENT_ID,
      client_secret: process.env.AZURE_CLIENT_SECRET,
      grant_type: 'client_credentials',
      scope: `${process.env.DATAVERSE_URL}/.default`
    });

    const authRes = await fetch(authUrl, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: authBody.toString()
    });

    const authData = await authRes.json();
    if (!authRes.ok) {
      console.error("❌ Token Request Failed:", authData);
      return;
    }
    console.log("✅ Token Acquired successfully.");

    // 2. Test Connection to Dataverse
    console.log("\n2. Testing Connectivity to Dataverse...");
    const testUrl = `${process.env.AZURE_URL || process.env.DATAVERSE_URL}api/data/v9.1/WhoAmI`;
    const testRes = await fetch(testUrl, {
      headers: { "Authorization": `Bearer ${authData.access_token}` }
    });

    const testData = await testRes.json();
    if (!testRes.ok) {
      console.error("❌ Connectivity Failed (WhoAmI):", testRes.status, testData);
      if (testRes.status === 401) {
        console.log("\n👉 HINT: Make sure the Azure App is registered as an 'Application User' in your Dataverse Environment.");
      }
      return;
    }
    console.log("✅ Connectivity Successful. Dataverse recognized the app.");
    console.log(`   User ID: ${testData.UserId}`);
    console.log(`   Org ID:  ${testData.OrganizationId}`);

    console.log("\n✅ ALL SYSTEMS READY. The Azure App has permission to connect.");
    console.log("If file uploads still fail, double-check that this Application User has a Security Role (e.g., 'System Administrator' or 'Environment Maker') that allows Writing to the Exception table.");

  } catch (err) {
    console.error("\n❌ Unexpected Error:", err.message);
  }
}

testAzureApp();
