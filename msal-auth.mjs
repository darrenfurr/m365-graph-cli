#!/usr/bin/env node

import { createRequire } from 'module';
import fs from 'fs';

const require = createRequire(import.meta.url);
const msal = require('@azure/msal-node');

const TENANT_ID = process.env.MS365_MCP_TENANT_ID;
const CLIENT_ID = process.env.MS365_MCP_CLIENT_ID;
const CLIENT_SECRET = process.env.MS365_MCP_CLIENT_SECRET;

const TOKEN_CACHE_PATH = '/data/.openclaw/workspace/scripts/m365-graph-cli/.m365-token-cache.json';

const config = {
  auth: {
    clientId: CLIENT_ID,
    authority: `https://login.microsoftonline.com/${TENANT_ID}`,
  },
};

const pca = new msal.PublicClientApplication(config);

const deviceCodeRequest = {
  deviceCodeCallback: (response) => {
    console.log('\n🔐 Microsoft Authentication Required\n');
    console.log('━'.repeat(60));
    console.log(`\n📱 Open: ${response.verificationUri}`);
    console.log(`🔢 Code: ${response.userCode}\n`);
    console.log('━'.repeat(60));
    console.log('\n⏳ Waiting for authentication...\n');
  },
  scopes: ['Calendars.Read', 'Mail.Read', 'Mail.ReadWrite', 'offline_access'],
};

async function authenticate() {
  try {
    const response = await pca.acquireTokenByDeviceCode(deviceCodeRequest);
    
    console.log('\n✅ Authentication successful!');
    console.log(`👤 User: ${response.account.username}`);
    
    // Save tokens
    const cache = pca.getTokenCache().serialize();
    fs.writeFileSync(TOKEN_CACHE_PATH, cache);
    console.log(`💾 Tokens saved to: ${TOKEN_CACHE_PATH}\n`);
    
  } catch (error) {
    console.error('\n❌ Authentication failed:', error.message);
    if (error.errorCode) console.error('Error code:', error.errorCode);
    if (error.errorMessage) console.error('Details:', error.errorMessage);
    process.exit(1);
  }
}

authenticate();
