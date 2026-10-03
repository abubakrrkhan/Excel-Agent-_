# Authentication Decision Required

## Current State

The application has NO authentication. Access control works as follows:

- Job IDs are random UUIDs (hard to guess, but not zero-risk)
- Download endpoint only serves `output-*` UUID-named files
- Any user who knows a jobId can poll its status and download the result
- No ownership check exists between a user and their jobs/files

---

## Choose Your Deployment Model

### Option A — Single-user / Personal Tool
**Best for:** Portfolio project, personal use, demo for a specific client.

**What it means:**
- Only you use this tool
- Deploy on a private network or VPN
- Or: add IP allowlisting at the infrastructure level (Railway/nginx)

**Changes required:** None in the application code.

**How to restrict access on Railway:**
- Use Railway's private networking (only your team)
- Or add a reverse proxy (Cloudflare Access, Tailscale)

---

### Option B — Static API Key (Recommended for demos)
**Best for:** Sharing with a small team, internship demos, client showcases.

**What it means:**
- One shared API key for all users
- Key is sent in every request header: `X-API-Key: your-secret-key`

**Changes required:**
1. Add `API_KEY` env var
2. Add one NestJS guard (20 lines)
3. Frontend sends `X-API-Key` header (one line in `api.ts`)

**Effort:** ~2 hours. Non-breaking — existing behavior unchanged when no key set.

---

### Option C — JWT Auth with User Ownership
**Best for:** Production SaaS, multiple independent users.

**What it means:**
- Users register/login
- Each job is tied to a userId
- Can only poll/download your own jobs

**Changes required:**
1. Add `users` table to Prisma schema
2. Add `@nestjs/passport` + JWT strategy
3. Add `userId` to job payloads
4. Add ownership checks on all job endpoints
5. Add login/register page to frontend

**Effort:** ~3 days. Breaking change — requires DB migration + frontend changes.

---

## Recommendation

| Scenario | Choose |
|----------|--------|
| Portfolio / demo / PAF-IAST project | **Option A** — no code changes needed |
| Sharing with uncle/supervisor for review | **Option B** — 2 hours of work |
| Real product with multiple clients | **Option C** — 3 days |

**Reply with A, B, or C and implementation begins immediately.**

---

## Security Risk (Honest Statement)

Without auth, exposure is:
- **Low** for a demo/portfolio — UUIDs are effectively unguessable in practice
- **Medium** if deployed publicly — a determined attacker with network access could probe job IDs
- **High** if files contain sensitive financial data and the app is publicly accessible

The current formula protection, path traversal guards, and output-only download restrictions remain in place regardless of auth choice.
