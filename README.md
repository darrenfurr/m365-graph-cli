# M365 Graph CLI

A lightweight CLI for Microsoft Graph API, designed for cron jobs and automation. No MCP server required - direct REST calls to Graph API.

## Features

- 📅 **Calendar** - View today's, tomorrow's, or this week's events
- 📧 **Email** - Read inbox, filter unread/priority
- 🔐 **Auth** - OAuth2 refresh token flow with client credentials
- 📊 **JSON output** - Perfect for scripting and automation

## Requirements

- Node.js 18+ (uses native `fetch`)
- Microsoft Entra (Azure AD) app registration with:
  - `Calendars.Read` (delegated)
  - `Mail.Read` / `Mail.ReadWrite` (delegated)
  - Client secret configured

## Setup

### 1. Clone and configure

```bash
git clone https://github.com/darrenfurr/m365-graph-cli.git
cd m365-graph-cli

# Copy and edit environment variables
cp .env.example .env
```

### 2. Set environment variables

```bash
export MS365_MCP_TENANT_ID="your-tenant-id"
export MS365_MCP_CLIENT_ID="your-client-id"
export MS365_MCP_CLIENT_SECRET="your-client-secret"
```

### 3. Initial authentication

On first run, you'll need a valid refresh token. Either:
- Copy an existing token cache from another auth flow
- Use device code flow to bootstrap (one-time interactive)

Token cache location: `./.m365-token-cache.json`

## Usage

### Calendar

```bash
# Today's events
node m365-cli.js calendar --today

# Tomorrow's events
node m365-cli.js calendar --tomorrow

# This week
node m365-cli.js calendar --week

# JSON output (for parsing)
node m365-cli.js calendar --tomorrow --json
```

### Email

```bash
# Recent emails (default: 20)
node m365-cli.js email

# Unread only
node m365-cli.js email --unread

# High priority unread
node m365-cli.js email --unread --priority

# Limit results
node m365-cli.js email --unread --limit=5

# Search
node m365-cli.js email --search="project update"

# JSON output
node m365-cli.js email --unread --json
```

### User Info

```bash
node m365-cli.js me
node m365-cli.js me --json
```

## JSON Output

All commands support `--json` for machine-readable output:

```bash
# Parse with jq
node m365-cli.js calendar --tomorrow --json | jq '.[].subject'

# Count unread
node m365-cli.js email --unread --json | jq 'length'

# Get first unread subject
node m365-cli.js email --unread --json --limit=1 | jq '.[0].subject'
```

## Cron Integration

Example: Morning brief at 6 AM on weekdays

```bash
# crontab -e
0 6 * * 1-5 cd /path/to/m365-graph-cli && node m365-cli.js calendar --today >> /var/log/morning-brief.log
```

## OpenClaw Integration

For OpenClaw cron jobs:

```javascript
{
  "name": "Morning Calendar Brief",
  "schedule": { "kind": "cron", "expr": "0 6 * * 1-5" },
  "payload": {
    "kind": "agentTurn",
    "message": "Run: node /path/to/m365-cli.js calendar --today --json"
  },
  "sessionTarget": "isolated"
}
```

## Token Management

The CLI uses a cascading token cache lookup:
1. Local `.m365-token-cache.json`
2. mcporter cache (if available)

Tokens auto-refresh using the client secret + refresh token grant.

## API Endpoints Used

| Feature | Graph API Endpoint |
|---------|-------------------|
| Calendar | `GET /me/calendarView` |
| Email | `GET /me/messages` |
| User | `GET /me` |

## Troubleshooting

### "No valid token or refresh token"

Run interactive auth once to bootstrap:
```bash
# Use device code flow from another tool, or
# Copy a valid token cache from mcporter
```

### Token refresh fails

Check that your app registration has:
- Client secret configured and not expired
- Required delegated permissions granted
- Admin consent (if required by org)

## License

MIT

## Author

Built for [OpenClaw](https://github.com/openclaw/openclaw) by Darren Furr
