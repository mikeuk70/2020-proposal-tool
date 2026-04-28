# 20.20 Proposal Generator

Upload a brief PDF. Get a branded PowerPoint back.

---

## Deploying to Railway (recommended — 10 minutes)

### 1. Create a GitHub repository

Create a new private repo on GitHub called `2020-proposal-tool`.
Upload all files from this folder into it, including `2020_template_slim_b64.txt`.

### 2. Deploy on Railway

1. Go to **railway.app** and sign up (free)
2. Click **New Project → Deploy from GitHub repo**
3. Select your `2020-proposal-tool` repo
4. Railway detects it as a Python app automatically

### 3. Add your API key

In Railway, go to your project → **Variables** tab → Add:
```
ANTHROPIC_API_KEY = sk-ant-your-key-here
```

### 4. Connect a custom domain (optional)

In Railway → your service → **Settings → Domains**:
- Click **Generate Domain** for a free railway.app URL, OR
- Add a custom domain like `proposals.lawliss.co.uk`
- Add the CNAME record Railway gives you to your DNS provider

### 5. Share the URL

Send the URL to 20.20. That is it.

---

## Files

| File | Purpose |
|------|---------|
| `app.py` | Flask application — all server logic |
| `templates/index.html` | Frontend — single page UI |
| `requirements.txt` | Python dependencies |
| `railway.toml` | Railway deployment config |
| `2020_template_slim_b64.txt` | 20.20 PowerPoint template (base64) |

---

## Environment variables

| Variable | Required | Description |
|----------|----------|-------------|
| `ANTHROPIC_API_KEY` | Yes | Your Anthropic API key |
| `PORT` | Auto | Set by Railway automatically |

---

## Costs

- **Railway**: Free tier available. Hobby plan ($5/month) recommended for reliability.
- **Anthropic API**: Approximately £0.05–0.10 per proposal at current Claude Sonnet pricing.
  At 1–2 proposals per week that is under £1/month.

---

## Notes

- Jobs are stored in memory. Restarting the server clears them (users should download before navigating away).
- The Anthropic API key never reaches the user's browser — all API calls happen server-side.
- The template file is read from disk on each generation so it is never exposed.
