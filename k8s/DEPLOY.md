# K3s / Hetzner Deployment Guide — Excel AI Agent

## Architecture

```
Internet
    │
    ▼
[Hetzner VPS — K3s single-node]
    │
    ▼
[nginx Ingress]
    ├─ /api/* ──────→ backend-svc:3001 (NestJS)
    └─ /*  ─────────→ frontend-svc:3000 (Next.js)
                              │
              ┌───────────────┼───────────────┐
              ▼               ▼               ▼
         postgres-svc     valkey-svc      uploads-pvc
         (PostgreSQL)     (Redis/Queue)   (Excel files)
```

---

## Prerequisites

On your Hetzner VPS (Ubuntu 22.04 recommended):

```bash
# Install K3s
curl -sfL https://get.k3s.io | sh -

# Wait for K3s to be ready
sudo k3s kubectl get nodes

# Install nginx ingress controller
# K3s ships with Traefik by default — install nginx instead:
kubectl apply -f https://raw.githubusercontent.com/kubernetes/ingress-nginx/controller-v1.10.0/deploy/static/provider/baremetal/deploy.yaml

# Or use Traefik (already in K3s) — change ingressClassName: nginx → traefik in ingress.yaml
```

On your **local machine** (where you build Docker images):

```bash
git clone <your-repo>
cd excel-ai-agent
```

---

## Step 1 — Configure Your Domain / IP

Edit `k8s/base/configmap.yaml`:

```yaml
data:
  EXTERNAL_URL: "http://YOUR-HETZNER-IP"   # or http://excel-ai.yourdomain.com
```

Edit `k8s/base/ingress.yaml`:

```yaml
  rules:
    - host: YOUR-HETZNER-IP   # or excel-ai.yourdomain.com
```

---

## Step 2 — Create Secrets

**Never commit real secrets.** Run on your Hetzner server:

```bash
kubectl create namespace excel-ai

kubectl create secret generic excel-ai-secrets \
  --namespace=excel-ai \
  --from-literal=DATABASE_URL='postgresql://postgres:YOUR_DB_PASSWORD@postgres-svc:5432/excel_ai_db' \
  --from-literal=AI_API_KEY='sk-ant-YOUR_ANTHROPIC_KEY' \
  --from-literal=POSTGRES_PASSWORD='YOUR_DB_PASSWORD' \
  --from-literal=VALKEY_URL='redis://valkey-svc:6379'
```

Delete the template secret.yaml from the manifests before applying:

```bash
# The secret.yaml in the repo is a TEMPLATE with CHANGE_ME values.
# After creating secrets with kubectl above, delete it from the manifest:
rm k8s/base/secret.yaml
# Remove secret.yaml from k8s/base/kustomization.yaml resources list too.
```

---

## Step 3 — Build Docker Images

From your **local machine** (repo root):

```bash
# Replace with your actual domain or IP
DOMAIN=http://YOUR-HETZNER-IP

make build DOMAIN=$DOMAIN
# This builds:
#   excel-ai-backend:latest
#   excel-ai-frontend:latest  (NEXT_PUBLIC_API_URL baked in = $DOMAIN)
```

---

## Step 4 — Transfer Images to Hetzner

### Option A — Direct import into K3s (no registry needed, simplest)

```bash
# Save images as tarballs
docker save excel-ai-backend:latest  | gzip > /tmp/backend.tar.gz
docker save excel-ai-frontend:latest | gzip > /tmp/frontend.tar.gz

# Transfer to Hetzner
scp /tmp/backend.tar.gz  root@YOUR-HETZNER-IP:/tmp/
scp /tmp/frontend.tar.gz root@YOUR-HETZNER-IP:/tmp/

# On Hetzner — import into K3s containerd
ssh root@YOUR-HETZNER-IP
sudo k3s ctr images import /tmp/backend.tar.gz
sudo k3s ctr images import /tmp/frontend.tar.gz

# Verify
sudo k3s ctr images ls | grep excel-ai
```

### Option B — Private registry (Docker Hub or Hetzner Registry)

```bash
REGISTRY=your-dockerhub-username
make push REGISTRY=$REGISTRY

# On Hetzner: update k8s manifests to use your registry images
# Then: kubectl apply -k k8s/overlays/prod
```

---

## Step 5 — Apply Kubernetes Manifests

On your **Hetzner VPS**:

```bash
# Clone the repo on the server (or scp the k8s/ directory)
git clone <your-repo>
cd excel-ai-agent

kubectl apply -k k8s/overlays/prod

# Watch pods start
kubectl get pods -n excel-ai -w
```

Expected output after ~60 seconds:

```
NAME                        READY   STATUS    RESTARTS
backend-xxx                 1/1     Running   0
frontend-xxx                1/1     Running   0
postgres-xxx                1/1     Running   0
valkey-xxx                  1/1     Running   0
```

---

## Step 6 — Verify Deployment

```bash
# Backend health
kubectl exec -n excel-ai deployment/backend -- \
  wget -qO- http://localhost:3001/api/health

# Expected: {"status":"ok","service":"excel-ai-backend"}

# Frontend accessible
curl -I http://YOUR-HETZNER-IP

# Expected: HTTP/1.1 200 OK
```

---

## Step 7 — Verify Full Workflow

1. Open `http://YOUR-HETZNER-IP` in your browser
2. Upload an Excel file (.xlsx)
3. Enter instruction: "Show me the sheet structure"
4. Click Send — agent should start processing
5. After completion: download the modified file

---

## Environment Variables Reference

### Secrets (kubectl create secret)

| Key | Value |
|-----|-------|
| `DATABASE_URL` | `postgresql://postgres:PASSWORD@postgres-svc:5432/excel_ai_db` |
| `AI_API_KEY` | Your Anthropic API key (`sk-ant-...`) |
| `POSTGRES_PASSWORD` | Strong random password |
| `VALKEY_URL` | `redis://valkey-svc:6379` |

### ConfigMap (k8s/base/configmap.yaml)

| Key | Value |
|-----|-------|
| `EXTERNAL_URL` | `http://YOUR-HETZNER-IP` or `http://your-domain.com` |
| `AI_MODEL` | `claude-sonnet-4-6` (default) |
| `AI_BASE_URL` | `https://api.anthropic.com` (default) |

### Build-time (Makefile)

| Variable | Used for |
|----------|---------|
| `DOMAIN` | `NEXT_PUBLIC_API_URL` baked into frontend image |

`DOMAIN` and `EXTERNAL_URL` must be the same value.

---

## Persistent Storage

Three PVCs are created automatically:

| PVC | Mount | Size | Purpose |
|-----|-------|------|---------|
| `uploads-pvc` | `/app/uploads` | 10Gi | Uploaded + output Excel files |
| `postgres-pvc` | `/var/lib/postgresql/data` | 5Gi | Database |
| `valkey-pvc` | `/data` | 1Gi | Queue + job results |

K3s uses local-path storage by default — data is stored in `/var/lib/rancher/k3s/storage/` on the node.

---

## Prisma Migrations

Migrations run automatically as a Kubernetes initContainer before the backend starts. No manual step needed. If a migration fails:

```bash
kubectl logs -n excel-ai \
  $(kubectl get pod -n excel-ai -l app=backend -o jsonpath='{.items[0].metadata.name}') \
  -c prisma-migrate
```

---

## TLS / HTTPS (Optional)

```bash
# Install cert-manager
kubectl apply -f https://github.com/cert-manager/cert-manager/releases/download/v1.14.4/cert-manager.yaml

# Add your domain to ingress.yaml and uncomment the tls block
# Then apply: kubectl apply -k k8s/overlays/prod
```

---

## Scaling (After Upgrading Storage)

The default PVC uses `ReadWriteOnce` — only one pod can mount it.

To scale backend beyond 1 replica:
1. Replace `ReadWriteOnce` with `ReadWriteMany` (requires Longhorn or NFS)
2. Then update `k8s/overlays/prod/kustomization.yaml` replicas

---

## Troubleshooting

```bash
# All pods status
kubectl get pods -n excel-ai

# Backend logs
kubectl logs -f deployment/backend -n excel-ai

# Frontend logs
kubectl logs -f deployment/frontend -n excel-ai

# Postgres logs
kubectl logs -f deployment/postgres -n excel-ai

# Describe a failing pod
kubectl describe pod -n excel-ai <pod-name>

# Check ingress
kubectl get ingress -n excel-ai
kubectl describe ingress excel-ai-ingress -n excel-ai
```
