# Railway Deployment Guide — Excel AI Agent

## Prerequisites
- [Railway account](https://railway.app) (free tier works)
- [Railway CLI](https://docs.railway.app/develop/cli): `npm install -g @railway/cli`
- Git repository (GitHub/GitLab)

---

## IMPORTANT: File Storage on Railway

**Railway's filesystem is ephemeral.** Every time your backend service redeploys, the local disk is wiped. This affects uploaded Excel files and processed output files stored in the `uploads/` directory.

**What this means for you:**

| Scenario | Impact |
|----------|--------|
| Redeploy during processing | Active job files may be lost |
| Redeploy after job completes | Output file unavailable for download |
| Service restart | All pending files lost |

**Required fix before public deployment:**

Add a Railway Volume to the backend service for persistent file storage:

```bash
# In Railway dashboard → your backend service → Volumes
# Mount path: /app/uploads
# Size: 1 GB (or as needed)
```

Or for production scale, configure an S3-compatible object store (AWS S3, Cloudflare R2) and update `FilesService` to use it instead of local disk.

**For demo / portfolio use:** Railway Volumes are sufficient. Files persist across redeploys. Free tier includes 1 GB.

---

## Architecture on Railway

```
┌─────────────────┐    ┌─────────────────┐
│  Next.js        │    │  NestJS         │
│  Frontend       │───▶│  Backend        │
│  (Port 3000)    │    │  (Port 3001)    │
└─────────────────┘    └────────┬────────┘
                                │
                    ┌───────────┴───────────┐
                    │                       │
               ┌────▼────┐          ┌──────▼──────┐
               │PostgreSQL│         │    Redis     │
               │(Railway) │         │  (Railway)   │
               └─────────┘          └─────────────┘
```

---

## Step 1 — Push to GitHub

```bash
cd excel-ai-agent
git init
git add .
git commit -m "initial commit"
gh repo create excel-ai-agent --public --source=. --push
```

---

## Step 2 — Create Railway Project

```bash
railway login
railway init
# Choose "Empty Project", name it "excel-ai-agent"
```

---

## Step 3 — Add Database Services

In Railway dashboard → your project → "+ New":

1. **Add PostgreSQL** → Railway auto-sets `DATABASE_URL`
2. **Add Redis** → Railway auto-sets `REDIS_URL`

---

## Step 4 — Deploy Backend

```bash
cd backend

# Create backend service
railway service create excel-ai-backend

# Set environment variables
railway variables set NODE_ENV=production
railway variables set AI_API_KEY=sk-ant-YOUR_KEY_HERE
railway variables set AI_MODEL=claude-sonnet-4-6
railway variables set AI_BASE_URL=https://api.anthropic.com
# DATABASE_URL and REDIS_URL are set automatically by Railway plugins

# Deploy
railway up --service excel-ai-backend
```

**Build command** (Railway auto-detects): `npm install && npm run build && npx prisma generate && npx prisma migrate deploy`

**Start command**: `node dist/main.js`

---

## Step 5 — Deploy Frontend

```bash
cd frontend

# Create frontend service
railway service create excel-ai-frontend

# Set environment variables
# Get your backend URL from: railway domain --service excel-ai-backend
railway variables set NEXT_PUBLIC_API_URL=https://excel-ai-backend-xxxx.railway.app

# Deploy
railway up --service excel-ai-frontend
```

**Build command**: `npm install && npm run build`

**Start command**: `npm start`

---

## Step 6 — Set CORS on Backend

Once you have the frontend Railway URL:

```bash
railway variables set CORS_ORIGIN=https://excel-ai-frontend-xxxx.railway.app \
  --service excel-ai-backend
```

---

## Step 7 — Verify

```bash
# Health check
curl https://excel-ai-backend-xxxx.railway.app/api/health

# Should return:
# { "status": "ok", "service": "excel-ai-backend" }
```

---

## Environment Variables Reference

| Variable | Service | Required | Description |
|----------|---------|----------|-------------|
| `NODE_ENV` | Backend | Yes | Set to `production` |
| `AI_API_KEY` | Backend | Yes | Anthropic API key |
| `AI_MODEL` | Backend | No | Default: `claude-sonnet-4-6` |
| `AI_BASE_URL` | Backend | No | Default: `https://api.anthropic.com` |
| `DATABASE_URL` | Backend | Auto | Set by Railway PostgreSQL plugin |
| `REDIS_URL` / `VALKEY_URL` | Backend | Auto | Set by Railway Redis plugin |
| `CORS_ORIGIN` | Backend | Yes | Your Railway frontend URL |
| `NEXT_PUBLIC_API_URL` | Frontend | Yes | Your Railway backend URL |

---

## Costs (Railway free tier)

| Resource | Free Tier |
|----------|-----------|
| Compute | $5 credit/month (enough for hobby use) |
| PostgreSQL | 1 GB storage free |
| Redis | 25 MB free |
| Bandwidth | 100 GB/month |

For heavy use, Railway Pro is $20/month.

---

## Troubleshooting

**Backend won't start:**
```bash
railway logs --service excel-ai-backend
```

**Redis connection refused:**
- Check that Railway Redis plugin is attached
- Verify `REDIS_URL` is set: `railway variables --service excel-ai-backend`

**Prisma migration failed:**
```bash
railway run --service excel-ai-backend npx prisma migrate deploy
```

**Frontend can't reach backend:**
- Check `NEXT_PUBLIC_API_URL` points to the correct backend URL
- Check `CORS_ORIGIN` on backend includes the frontend URL

---

## Custom Domain (Optional)

In Railway dashboard → your service → Settings → Domains → "Add Custom Domain"

Then update `CORS_ORIGIN` and `NEXT_PUBLIC_API_URL` accordingly.
