#!/usr/bin/env node

/**
 * M365 CLI - Microsoft Graph API for OpenClaw
 * 
 * Usage:
 *   node m365-cli.js calendar --today
 *   node m365-cli.js calendar --tomorrow --json
 *   node m365-cli.js email --unread
 *   node m365-cli.js email --unread --priority --json
 */

import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import { createRequire } from 'module';

const require = createRequire(import.meta.url);
const msal = require('@azure/msal-node');

const __dirname = path.dirname(fileURLToPath(import.meta.url));

// Config from environment
const TENANT_ID = process.env.MS365_MCP_TENANT_ID;
const CLIENT_ID = process.env.MS365_MCP_CLIENT_ID;

// Token cache paths
const TOKEN_CACHE_PATHS = [
  path.join(__dirname, '.m365-token-cache.json'),
  '/data/.npm/_npx/813b81b976932cb5/node_modules/@softeria/ms-365-mcp-server/.token-cache.json',
  '/data/.openclaw/tools/m365/.token-cache.json',
];

const GRAPH_BASE = 'https://graph.microsoft.com/v1.0';

// ============ TOKEN MANAGEMENT ============

function findTokenCache() {
  for (const p of TOKEN_CACHE_PATHS) {
    if (fs.existsSync(p)) return p;
  }
  return TOKEN_CACHE_PATHS[0]; // Default to local
}

function loadCache() {
  const cachePath = findTokenCache();
  if (!fs.existsSync(cachePath)) return null;
  try {
    return JSON.parse(fs.readFileSync(cachePath, 'utf-8'));
  } catch {
    return null;
  }
}

function saveCache(cache) {
  const cachePath = TOKEN_CACHE_PATHS[0]; // Save to local
  fs.writeFileSync(cachePath, JSON.stringify(cache, null, 2));
}

async function getAccessToken() {
  const cachePath = findTokenCache();
  
  if (!fs.existsSync(cachePath)) {
    throw new Error('No token cache found. Run: node msal-auth.mjs');
  }

  const cacheData = fs.readFileSync(cachePath, 'utf-8');
  
  // Create MSAL client with cache
  const pca = new msal.PublicClientApplication({
    auth: {
      clientId: CLIENT_ID,
      authority: `https://login.microsoftonline.com/${TENANT_ID}`,
    },
    cache: {
      cachePlugin: {
        beforeCacheAccess: async (context) => {
          context.tokenCache.deserialize(cacheData);
        },
        afterCacheAccess: async (context) => {
          if (context.cacheHasChanged) {
            fs.writeFileSync(cachePath, context.tokenCache.serialize());
          }
        },
      },
    },
  });

  // Get accounts from cache
  const accounts = await pca.getTokenCache().getAllAccounts();
  if (accounts.length === 0) {
    throw new Error('No account found. Run: node msal-auth.mjs');
  }

  try {
    // Silent token acquisition (auto-refreshes using MSAL)
    const response = await pca.acquireTokenSilent({
      account: accounts[0],
      scopes: ['Calendars.Read', 'Mail.Read', 'Mail.ReadWrite'],
    });

    return response.accessToken;
  } catch (error) {
    // If silent acquisition fails, need interactive re-auth
    throw new Error(`Token acquisition failed: ${error.message}. Run: node msal-auth.mjs`);
  }
}

// ============ GRAPH API ============

async function graphApi(endpoint) {
  const token = await getAccessToken();
  
  const response = await fetch(`${GRAPH_BASE}${endpoint}`, {
    headers: {
      'Authorization': `Bearer ${token}`,
      'Content-Type': 'application/json',
    },
  });

  if (!response.ok) {
    const error = await response.text();
    throw new Error(`Graph API error (${response.status}): ${error.substring(0, 200)}`);
  }

  return response.json();
}

// ============ CALENDAR ============

async function getCalendarEvents(startDate, endDate) {
  const start = startDate.toISOString();
  const end = endDate.toISOString();
  
  const params = new URLSearchParams({
    startDateTime: start,
    endDateTime: end,
    $orderby: 'start/dateTime',
    $top: '50',
    $select: 'id,subject,start,end,location,isAllDay,organizer,webLink',
  });

  const response = await graphApi(`/me/calendarView?${params}`);
  return response.value;
}

function getTodayRange() {
  const start = new Date();
  start.setHours(0, 0, 0, 0);
  const end = new Date(start);
  end.setDate(end.getDate() + 1);
  return { start, end };
}

function getTomorrowRange() {
  const start = new Date();
  start.setDate(start.getDate() + 1);
  start.setHours(0, 0, 0, 0);
  const end = new Date(start);
  end.setDate(end.getDate() + 1);
  return { start, end };
}

function getWeekRange() {
  const start = new Date();
  start.setHours(0, 0, 0, 0);
  const end = new Date(start);
  end.setDate(end.getDate() + 7);
  return { start, end };
}

function formatCalendarEvents(events) {
  if (events.length === 0) return 'No events found.';
  
  let currentDate = '';
  const lines = [];
  
  for (const event of events) {
    const startDate = new Date(event.start.dateTime + 'Z');
    const dateStr = startDate.toLocaleDateString('en-US', {
      weekday: 'short',
      month: 'short',
      day: 'numeric',
    });
    
    if (dateStr !== currentDate) {
      if (currentDate) lines.push('');
      lines.push(`📅 ${dateStr}`);
      lines.push('─'.repeat(40));
      currentDate = dateStr;
    }
    
    const timeStr = event.isAllDay
      ? 'All day'
      : startDate.toLocaleTimeString('en-US', { hour: '2-digit', minute: '2-digit' });
    
    lines.push(`  ${timeStr}  ${event.subject}`);
    if (event.location?.displayName) {
      lines.push(`           📍 ${event.location.displayName}`);
    }
  }
  
  return lines.join('\n');
}

// ============ EMAIL ============

async function getEmails(options = {}) {
  const { unread = false, priority = false, limit = 20, search } = options;
  
  const filters = [];
  if (unread) filters.push('isRead eq false');
  if (priority) filters.push("importance eq 'high'");
  
  const params = new URLSearchParams({
    $top: limit.toString(),
    $orderby: 'receivedDateTime desc',
    $select: 'id,subject,from,receivedDateTime,isRead,importance,bodyPreview,hasAttachments',
  });
  
  if (filters.length > 0) {
    params.append('$filter', filters.join(' and '));
  }
  
  if (search) {
    params.append('$search', `"${search}"`);
  }

  const response = await graphApi(`/me/messages?${params}`);
  return response.value;
}

function formatEmails(emails) {
  if (emails.length === 0) return 'No emails found.';
  
  const lines = [];
  
  for (const email of emails) {
    const date = new Date(email.receivedDateTime);
    const dateStr = date.toLocaleString('en-US', {
      month: 'short',
      day: 'numeric',
      hour: '2-digit',
      minute: '2-digit',
    });
    
    const icon = email.isRead ? '📧' : '📩';
    const priority = email.importance === 'high' ? '❗' : '';
    const attachment = email.hasAttachments ? '📎' : '';
    
    lines.push(`${icon}${priority} ${email.subject} ${attachment}`);
    lines.push(`   From: ${email.from?.emailAddress?.name || email.from?.emailAddress?.address || 'Unknown'}`);
    lines.push(`   Date: ${dateStr}`);
    if (email.bodyPreview) {
      lines.push(`   ${email.bodyPreview.substring(0, 80).replace(/\n/g, ' ')}...`);
    }
    lines.push('');
  }
  
  return lines.join('\n');
}

// ============ CLI ============

function parseArgs() {
  const args = process.argv.slice(2);
  const command = args[0];
  
  const flags = {
    json: args.includes('--json'),
    today: args.includes('--today'),
    tomorrow: args.includes('--tomorrow'),
    week: args.includes('--week'),
    unread: args.includes('--unread'),
    priority: args.includes('--priority'),
  };
  
  // Parse --limit=N
  const limitArg = args.find(a => a.startsWith('--limit='));
  flags.limit = limitArg ? parseInt(limitArg.split('=')[1]) : 20;
  
  // Parse --search="query"
  const searchArg = args.find(a => a.startsWith('--search='));
  flags.search = searchArg ? searchArg.split('=')[1].replace(/^["']|["']$/g, '') : null;
  
  return { command, flags };
}

async function main() {
  const { command, flags } = parseArgs();
  
  try {
    switch (command) {
      case 'calendar': {
        let range;
        if (flags.tomorrow) {
          range = getTomorrowRange();
        } else if (flags.week) {
          range = getWeekRange();
        } else {
          range = getTodayRange();
        }
        
        const events = await getCalendarEvents(range.start, range.end);
        
        if (flags.json) {
          console.log(JSON.stringify(events, null, 2));
        } else {
          console.log(formatCalendarEvents(events));
        }
        break;
      }
      
      case 'email': {
        const emails = await getEmails({
          unread: flags.unread,
          priority: flags.priority,
          limit: flags.limit,
          search: flags.search,
        });
        
        if (flags.json) {
          console.log(JSON.stringify(emails, null, 2));
        } else {
          console.log(formatEmails(emails));
        }
        break;
      }
      
      case 'me': {
        const me = await graphApi('/me');
        if (flags.json) {
          console.log(JSON.stringify(me, null, 2));
        } else {
          console.log(`👤 ${me.displayName}`);
          console.log(`📧 ${me.mail || me.userPrincipalName}`);
          console.log(`🏢 ${me.jobTitle || 'N/A'}`);
        }
        break;
      }
      
      default:
        console.log(`
M365 CLI - Microsoft Graph API for OpenClaw

Usage:
  node m365-cli.js <command> [options]

Commands:
  calendar    Get calendar events
    --today     Today's events (default)
    --tomorrow  Tomorrow's events
    --week      Next 7 days

  email       Get emails
    --unread    Unread only
    --priority  High priority only
    --limit=N   Number of results (default: 20)
    --search="query"  Search emails

  me          Current user info

Options:
  --json      Output as JSON

Examples:
  node m365-cli.js calendar --tomorrow
  node m365-cli.js calendar --tomorrow --json
  node m365-cli.js email --unread --priority
  node m365-cli.js email --unread --json --limit=10
`);
    }
  } catch (error) {
    if (flags.json) {
      console.log(JSON.stringify({ error: error.message }));
    } else {
      console.error('❌ Error:', error.message);
    }
    process.exit(1);
  }
}

main();
