# Expandi → Microsoft Teams Notification Relay

Receives Expandi webhook events and forwards them as Adaptive Card messages to Microsoft Teams channels via Incoming Webhooks.

## Architecture

```
Expandi fires webhook → Vercel serverless function → Formats to Adaptive Card → POSTs to user's Teams Webhook
```

**Zero database.** The user's Teams webhook URL is Base64url-encoded into the endpoint path. The serverless function decodes it, formats the Expandi payload into a rich Adaptive Card, and POSTs to Teams. Fully stateless.

## Project Structure

```
expandi-teams/
├── api/
│   └── hook/
│       └── [slug].js       ← Serverless function (Vercel route: /api/hook/:slug)
├── lib/
│   └── format.js           ← Adaptive Card message formatter (all 27 event types)
├── public/
│   └── index.html           ← Setup page (paste Teams URL → get endpoint)
├── test/
│   └── fixture-reply.json   ← Sample Expandi webhook payload
├── package.json
├── vercel.json              ← Vercel routing config
└── README.md
```

## How It Works (User Flow)

1. User creates a Teams Incoming Webhook (Channel → Connectors → Incoming Webhook → Configure → Copy URL)
2. User visits the setup page (`/index.html`)
3. Pastes their Teams webhook URL
4. Page generates a unique Expandi webhook endpoint by Base64url-encoding the Teams URL into the path
5. User copies the generated endpoint URL
6. In Expandi: **LinkedIn Settings → Webhooks → Add a webhook** → paste the endpoint URL → select event type
7. Done. Events fire from Expandi → land in Teams.

## Supported Webhook URLs

Both Microsoft Teams webhook formats are supported:

- **Legacy Connectors:** `https://xxx.webhook.office.com/webhookb2/...`
- **New Workflows:** `https://xxx.logic.azure.com:443/workflows/...`

Microsoft is transitioning from Office 365 Connectors to the Workflows app. Both work with this relay.

## Supported Events (27 total)

All Expandi webhook events are supported with event-specific Adaptive Card layouts. See the [Slack version README](https://github.com/cosankov/expandi-slack) for the full event list.

## Deploy to Vercel

```bash
cd expandi/tools/expandi-teams
vercel
```

## Tech Stack

| Layer | Tool |
|---|---|
| Hosting | Vercel (serverless) |
| Runtime | Node.js 18+ (single serverless function) |
| Frontend | Static HTML (no framework) |
| Message format | Adaptive Cards v1.4 |
| Database | None |
