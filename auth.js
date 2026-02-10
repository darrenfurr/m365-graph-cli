#!/usr/bin/env node

/**
 * M365 Device Code Authentication
 * 
 * Usage: node auth.js
 * 
 * Initiates Microsoft device code flow for user authentication.
 * Saves tokens to .m365-token-cache.json for use by m365-cli.js
 */

import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));

// Config from environment
const TENANT_ID = process.env.MS365_MCP_TENANT_ID;
const CLIENT_ID = process.env.MS365_MCP_CLIENT_ID;
const CLIENT_SECRET = process.env.MS365_MCP_CLIENT_SECRET;

// Token cache path
const TOKEN_CACHE_PATH = path.join(__dirname, '.m365-token-cache.json');

// Scopes for calendar and mail access
const SCOPES = 'Calendars.Read Mail.Read Mail.ReadWrite offline_access';

async function initiateDeviceCodeFlow() {
  console.log('🔐 Starting Microsoft Device Code Authentication\n');

  if (!TENANT_ID || !CLIENT_ID) {
    console.error('❌ Missing environment variables:');
    console.error('   MS365_MCP_TENANT_ID');
    console.error('   MS365_MCP_CLIENT_ID');
    process.exit(1);
  }

  // Step 1: Request device code
  const deviceCodeUrl = `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/devicecode`;
  
  const deviceCodeResponse = await fetch(deviceCodeUrl, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({
      client_id: CLIENT_ID,
      scope: SCOPES,
    }).toString(),
  });

  if (!deviceCodeResponse.ok) {
    const error = await deviceCodeResponse.text();
    console.error('❌ Failed to get device code:', error);
    process.exit(1);
  }

  const deviceCode = await deviceCodeResponse.json();

  // Step 2: Display instructions to user
  console.log('━'.repeat(60));
  console.log('\n📱 To sign in, use a web browser to open:\n');
  console.log(`   ${deviceCode.verification_uri}\n`);
  console.log(`   And enter the code: ${deviceCode.user_code}\n`);
  console.log('━'.repeat(60));
  console.log(`\n⏳ Waiting for authentication (expires in ${Math.floor(deviceCode.expires_in / 60)} minutes)...\n`);

  // Step 3: Poll for token
  const tokenUrl = `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`;
  const pollInterval = deviceCode.interval * 1000 || 5000;
  const expiresAt = Date.now() + (deviceCode.expires_in * 1000);

  while (Date.now() < expiresAt) {
    await sleep(pollInterval);

    const tokenBody = {
      client_id: CLIENT_ID,
      grant_type: 'urn:ietf:params:oauth:grant-type:device_code',
      device_code: deviceCode.device_code,
    };

    // Add client_secret if available (for confidential clients)
    if (CLIENT_SECRET) {
      tokenBody.client_secret = CLIENT_SECRET;
    }

    const tokenResponse = await fetch(tokenUrl, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams(tokenBody).toString(),
    });

    const tokenResult = await tokenResponse.json();

    if (tokenResult.error) {
      if (tokenResult.error === 'authorization_pending') {
        // User hasn't completed auth yet, keep polling
        process.stdout.write('.');
        continue;
      } else if (tokenResult.error === 'slow_down') {
        // Server asking us to slow down
        await sleep(5000);
        continue;
      } else if (tokenResult.error === 'expired_token') {
        console.error('\n\n❌ Authentication expired. Please run again.');
        process.exit(1);
      } else {
        console.error(`\n\n❌ Authentication error: ${tokenResult.error_description || tokenResult.error}`);
        process.exit(1);
      }
    }

    // Success! We have tokens
    console.log('\n\n✅ Authentication successful!\n');
    
    // Save tokens in MSAL-compatible cache format
    saveTokenCache(tokenResult);
    
    // Show user info
    await showUserInfo(tokenResult.access_token);
    
    console.log(`\n💾 Tokens saved to: ${TOKEN_CACHE_PATH}`);
    console.log('\n🎉 You can now use m365-cli.js commands!\n');
    return;
  }

  console.error('\n\n❌ Authentication timed out. Please run again.');
  process.exit(1);
}

function saveTokenCache(tokens) {
  const now = Math.floor(Date.now() / 1000);
  const expiresOn = now + tokens.expires_in;
  
  // Create MSAL-compatible cache structure
  const homeAccountId = `${TENANT_ID}.${TENANT_ID}`;
  const environment = 'login.microsoftonline.com';
  
  const cache = {
    Account: {
      [`${homeAccountId}-${environment}-${TENANT_ID}`]: {
        home_account_id: homeAccountId,
        environment: environment,
        realm: TENANT_ID,
        local_account_id: TENANT_ID,
        username: 'user@domain.com', // Will be updated if we can fetch user info
        authority_type: 'MSSTS',
      }
    },
    AccessToken: {
      [`${homeAccountId}-${environment}-accesstoken-${CLIENT_ID}-${TENANT_ID}-${SCOPES.replace(/ /g, ' ')}`]: {
        home_account_id: homeAccountId,
        environment: environment,
        credential_type: 'AccessToken',
        client_id: CLIENT_ID,
        secret: tokens.access_token,
        realm: TENANT_ID,
        target: SCOPES,
        cached_at: now.toString(),
        expires_on: expiresOn.toString(),
        extended_expires_on: (expiresOn + 3600).toString(),
        token_type: 'Bearer',
      }
    },
    RefreshToken: {},
    IdToken: {},
    AppMetadata: {},
  };

  // Add refresh token if present
  if (tokens.refresh_token) {
    cache.RefreshToken[`${homeAccountId}-${environment}-refreshtoken-${CLIENT_ID}----`] = {
      home_account_id: homeAccountId,
      environment: environment,
      credential_type: 'RefreshToken',
      client_id: CLIENT_ID,
      secret: tokens.refresh_token,
    };
  }

  // Add ID token if present
  if (tokens.id_token) {
    cache.IdToken[`${homeAccountId}-${environment}-idtoken-${CLIENT_ID}-${TENANT_ID}---`] = {
      home_account_id: homeAccountId,
      environment: environment,
      credential_type: 'IdToken',
      client_id: CLIENT_ID,
      secret: tokens.id_token,
      realm: TENANT_ID,
    };
  }

  fs.writeFileSync(TOKEN_CACHE_PATH, JSON.stringify(cache, null, 2));
}

async function showUserInfo(accessToken) {
  try {
    const response = await fetch('https://graph.microsoft.com/v1.0/me', {
      headers: { 'Authorization': `Bearer ${accessToken}` },
    });

    if (response.ok) {
      const user = await response.json();
      console.log(`👤 Signed in as: ${user.displayName}`);
      console.log(`📧 Email: ${user.mail || user.userPrincipalName}`);
      
      // Update cache with actual username
      if (fs.existsSync(TOKEN_CACHE_PATH)) {
        const cache = JSON.parse(fs.readFileSync(TOKEN_CACHE_PATH, 'utf-8'));
        const accountKey = Object.keys(cache.Account)[0];
        if (accountKey && cache.Account[accountKey]) {
          cache.Account[accountKey].username = user.userPrincipalName || user.mail;
          cache.Account[accountKey].name = user.displayName;
          fs.writeFileSync(TOKEN_CACHE_PATH, JSON.stringify(cache, null, 2));
        }
      }
    }
  } catch (e) {
    // Non-fatal - just couldn't get user info
  }
}

function sleep(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

// Run
initiateDeviceCodeFlow().catch(err => {
  console.error('❌ Error:', err.message);
  process.exit(1);
});
