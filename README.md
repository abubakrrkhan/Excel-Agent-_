# Excel AI Agent

An AI-powered Excel automation system that inspects workbooks, locates target cells using evidence, and applies safe, verified modifications.

## Architecture

```
Next.js 14 Frontend
        ↓
POST /api/jobs/submit  (async — instant jobId)
        ↓
BullMQ Queue + Redis
        ↓ background worker
AgentService (multi-round tool-calling)
        ↓
OperationsService.validatePlan()
        ↓
OperationExecutionRouter
        ↓
ExcelService (formula-protected writes)
        ↓
GET /api/jobs/:id  (2s polling)
```

## Requirements

- Node.js 20+
- PostgreSQL 15+
- Redis 7+ (or Valkey)
- npm 9+ (lockfiles included — `npm ci` works out of the box)

## Local Development

```bash
# 1. Clone and install
git clone <repo>
cd excel-ai-app

# 2. Configure environment
cp .env.example .env
# Edit .env — set AI_API_KEY, DATABASE_URL, VALKEY_URL (or REDIS_URL)

# 3. Start dependencies
docker-compose up -d  # starts postgres + valkey (Redis-compatible)

# 4. Backend
cd backend
npm ci                      # uses package-lock.json — reproducible
npx prisma migrate deploy
npm run start:dev

# 5. Frontend (new terminal)
cd frontend
npm ci
npm run dev
```

Open http://localhost:3000

## Environment Variables

| Variable | Required | Description |
|----------|----------|-------------|
| `AI_API_KEY` | Yes | Google Gemini API key (`AIza...`) — free at https://aistudio.google.com/apikey |
| `AI_MODEL` | No | Default: `gemini-2.0-flash` (free tier) |
| `DATABASE_URL` | Yes | PostgreSQL connection string |
| `REDIS_URL` | Railway | Set automatically by Railway Redis plugin. Takes priority over `VALKEY_URL`. |
| `VALKEY_URL` | Self-hosted | Redis/Valkey URL for Docker or bare-metal (`redis://...`). Used when `REDIS_URL` is absent. |
| `VALKEY_HOST` | No | Default: `localhost` (used only when no URL set) |
| `VALKEY_PORT` | No | Default: `6379` (used only when no URL set) |
| `CORS_ORIGIN` | Yes | Frontend URL (e.g. `http://localhost:3000`) |
| `PORT` | No | Backend port (default: `3001`) |
| `FILE_RETENTION_INPUT_HOURS` | No | Hours to keep uploaded input files (default: `24`) |
| `FILE_RETENTION_OUTPUT_HOURS` | No | Hours to keep processed output files (default: `2`) |
| `FILE_CLEANUP_CRON` | No | Cron expression for cleanup schedule (default: `0 3 * * *` = 3 AM daily) |
| `RATE_LIMIT_MAX` | No | Max job submissions per IP per TTL window (default: `20`) |
| `RATE_LIMIT_TTL` | No | Rate limit window in seconds (default: `60`) |

### Redis connection precedence

The backend resolves the Redis URL in this order:

```
1. REDIS_URL   — set automatically by Railway's Redis plugin
2. VALKEY_URL  — for self-hosted Valkey / Docker Compose / bare-metal
3. VALKEY_HOST + VALKEY_PORT  — fallback individual fields
4. redis://localhost:6379     — local development default
```

## API Endpoints

| Method | Path | Description |
|--------|------|-------------|
| GET | `/api/health` | Health check |
| POST | `/api/files/upload` | Upload workbook (get metadata) |
| POST | `/api/jobs/submit` | Submit async job (returns `jobId`) |
| GET | `/api/jobs/:id` | Poll job status |
| GET | `/api/files/download/:filename` | Download processed file |

## Supported Operations

**Write:** SET_CELL, SET_CELL_RANGE, ADD_ROW, DELETE_ROW, ADD_FORMULA, FORMAT_CELL, ADD_COLUMN_HEADER, SORT_SHEET

**Read (inspection):** INSPECT_WORKBOOK_STRUCTURE, DETECT_HEADERS, DETECT_SECTIONS, GET_CELL_CONTEXT

**Hybrid:** SET_CELL_BY_LABEL (finds label, resolves value cell, writes — blocks on ambiguity)

## File Format Support

**Supported:** `.xlsx` (Excel 2007+)

**Not supported:** `.xls` — ExcelJS cannot write `.xls`. Convert to `.xlsx` before uploading.

## File Limits

- Maximum upload size: **10 MB**
- Output files retained for: **1 hour** (configurable)
- Uploaded input files retained for: **24 hours** (configurable)

## Async Job Flow

1. `POST /api/jobs/submit` → returns `{ jobId }` immediately
2. Frontend polls `GET /api/jobs/:id` every 2 seconds
3. Job states: `queued` → `processing` → `completed` / `clarification_needed` / `failed`
4. When `clarification_needed`: re-submit with `clarificationAnswer` body field

## Clarification Flow

When the agent finds a label in multiple locations, it returns `clarification_needed`. Re-submit:

```
POST /api/jobs/submit
  file: [same file]
  instruction: [original instruction]
  clarificationAnswer: [user's selection]
```

The processor combines these into: `"original instruction\n\nUser clarification: answer"`.

## Safety Features

- **Formula protection**: Writing to formula cells is blocked by default. Pass `allowFormulaOverwrite: true` to authorize.
- **Atomic range writes**: `SET_CELL_RANGE` scans the entire range for formula cells before writing any cell.
- **Evidence-based agent**: The AI must call inspection tools and verify target cells before producing a write plan. It cannot invent coordinates.
- **No base64 in queue**: File bytes are saved to disk before queuing. Only a storage reference is in the Redis payload.
- **Download guard**: Only `output-*` files can be downloaded. Input files and other disk files are inaccessible.

## Running Tests

```bash
cd backend
npm test           # run all 518 tests
npm run typecheck  # TypeScript check
npm run build      # production build
```

## Deployment (Render.com)

[![Deploy to Render](https://render.com/images/deploy-to-render-button.svg)](https://render.com/deploy?repo=https://github.com/abubakrrkhan/Excel-Agent-_)

**One-click deploy** creates all 4 services automatically:
- PostgreSQL (`excel-ai-db`) — free
- Redis (`excel-ai-redis`) — free  
- NestJS Backend (`excel-ai-backend`) — free
- Next.js Frontend (`excel-ai-frontend`) — free

**After deploy:** Go to `excel-ai-backend` → Environment → set `AI_API_KEY` to your [Gemini API key](https://aistudio.google.com/apikey).
