# PWA Call Logger → HubSpot

Post-call data entry tool for Pacific West Academy enrollment team.
Covers 4 playbooks: Initial Inquiry, Application Follow-Up, Applicant Intake, Veteran/GI Bill.

## Deploy to Railway

1. Push this folder to a GitHub repo (private is fine)
2. Go to railway.app → New Project → Deploy from GitHub repo
3. Select this repo
4. Go to **Variables** tab and add:
   ```
   HUBSPOT_TOKEN = pat-na1-your-token-here
   ```
5. Railway auto-deploys. Your team URL will be something like `pwa-call-logger.up.railway.app`

## HubSpot Private App Setup

1. HubSpot → Settings → Integrations → Private Apps → Create private app
2. Name it "PWA Call Logger"
3. Scopes needed:
   - `crm.objects.contacts.read`
   - `crm.objects.contacts.write`
   - `crm.objects.notes.write`
4. Copy the token → paste into Railway Variables as `HUBSPOT_TOKEN`

## Local development

```bash
npm install
HUBSPOT_TOKEN=your-token node server.js
# Open http://localhost:3000
```
