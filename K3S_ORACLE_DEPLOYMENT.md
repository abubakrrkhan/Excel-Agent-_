# K3s Deployment Guide — Oracle Always Free + Excel AI Agent

## Architecture

```
Internet (port 80)
        │
        ▼
[Oracle VM — ARM64 Ampere A1]
        │
        ▼
  K3s + Traefik Ingress
        │
        ├─ /api/* ──────▶ backend-svc:3001  (NestJS)
        └─ /*  ─────────▶ frontend-svc:3000 (Next.js)

Internal K8s services (NOT publicly exposed):
  postgres-svc:5432   ← PostgreSQL
  valkey-svc:6379     ← Valkey/Redis (BullMQ queue)

Persistent storage (survives pod restarts):
  uploads-pvc  →  /app/uploads   ← uploaded + output Excel files
  postgres-pvc →  /var/lib/postgresql/data
  valkey-pvc   →  /data
```

---

## Prerequisites (on your LOCAL machine)

- Docker Desktop with buildx enabled
- `ssh` and `scp`
- `git`

---

# PART 1 — ORACLE CLOUD SETUP

## Step 1 — Create Oracle Always Free VM

1. Go to [cloud.oracle.com](https://cloud.oracle.com)
2. Create a free account (credit card required for verification — NOT charged)
3. Go to: **Compute → Instances → Create Instance**
4. Settings:
   ```
   Name:    excel-ai-agent
   Image:   Canonical Ubuntu 22.04
   Shape:   Ampere → VM.Standard.A1.Flex
             OCPU: 2    (free allowance: up to 4)
             RAM:  12GB (free allowance: up to 24GB)
   ```
5. Under **Add SSH Keys** → paste your public key (`~/.ssh/id_rsa.pub`)
6. Click **Create**

## Step 2 — Get Your Server IP

After instance starts, copy the **Public IP Address** from the instance detail page.

```
Example: 130.61.2.15
```

## Step 3 — Open Firewall Ports

Oracle has TWO firewalls. Both must allow port 80.

### 3a — Oracle Security List (Cloud Console)
1. Instance page → **Subnet** → **Security List**
2. Add **Ingress Rule**:
   ```
   Source CIDR:   0.0.0.0/0
   Protocol:      TCP
   Port Range:    80
   ```
3. Also add rule for port 443 (for future HTTPS)

### 3b — Ubuntu firewall (on the server)
```bash
ssh ubuntu@YOUR-ORACLE-IP
sudo iptables -I INPUT -p tcp --dport 80  -j ACCEPT
sudo iptables -I INPUT -p tcp --dport 443 -j ACCEPT
sudo iptables -I INPUT -p tcp --dport 6443 -j ACCEPT  # K3s API
sudo netfilter-persistent save
sudo apt-get install -y iptables-persistent   # if save fails
```

## Step 4 — SSH Into Server

```bash
ssh ubuntu@YOUR-ORACLE-IP
```

---

# PART 2 — K3S INSTALLATION

## Step 5 — Install K3s

```bash
# K3s with Traefik (default — we use Traefik as ingress)
curl -sfL https://get.k3s.io | sh -

# Wait for K3s to start
sudo systemctl status k3s
```

## Step 6 — Configure kubectl

```bash
mkdir -p ~/.kube
sudo cp /etc/rancher/k3s/k3s.yaml ~/.kube/config
sudo chown ubuntu:ubuntu ~/.kube/config
echo 'export KUBECONFIG=~/.kube/config' >> ~/.bashrc
source ~/.bashrc
```

## Step 7 — Verify K3s

```bash
kubectl get nodes
# Expected:
# NAME   STATUS   ROLES                  AGE   VERSION
# vm1    Ready    control-plane,master   1m    v1.x.x+k3s1

kubectl get pods -A
# Expected: traefik running in kube-system
```

---

# PART 3 — BUILD AND TRANSFER DOCKER IMAGES

## Step 8 — Build Images on Your LOCAL Machine

```bash
# Clone your repo
git clone YOUR-REPO-URL
cd excel-ai-agent

# Set your Oracle VM IP
export ORACLE_IP=130.61.2.15   # ← YOUR actual IP

# Build ARM64 images for Oracle Ampere A1
make build-oracle ORACLE_IP=$ORACLE_IP
```

This builds:
- `excel-ai-backend:latest` (ARM64)
- `excel-ai-frontend:latest` (ARM64, with `NEXT_PUBLIC_API_URL=http://ORACLE_IP.nip.io`)

## Step 9 — Transfer Images to Oracle VM

```bash
make push-oracle ORACLE_IP=$ORACLE_IP
```

This:
1. Saves images as `.tar.gz`
2. `scp`s them to the Oracle VM
3. Imports them into K3s containerd

Verify on Oracle VM:
```bash
sudo k3s ctr images ls | grep excel-ai
# Should show: excel-ai-backend:latest, excel-ai-frontend:latest
```

---

# PART 4 — KUBERNETES DEPLOYMENT

All remaining commands run on the **Oracle VM**.

## Step 10 — Create Namespace

```bash
kubectl apply -f k8s/base/namespace.yaml
kubectl get namespace excel-ai
```

## Step 11 — Create Secrets

```bash
# Generate a strong DB password
DB_PASSWORD=$(openssl rand -base64 24 | tr -d '/+=' | head -c 32)
echo "Save this password: $DB_PASSWORD"

# Create the secret (never stored in git)
kubectl create secret generic excel-ai-secrets \
  --namespace=excel-ai \
  --from-literal=DATABASE_URL="postgresql://postgres:${DB_PASSWORD}@postgres-svc:5432/excel_ai_db" \
  --from-literal=AI_API_KEY="sk-ant-YOUR_ANTHROPIC_KEY" \
  --from-literal=POSTGRES_PASSWORD="${DB_PASSWORD}" \
  --from-literal=VALKEY_URL="redis://valkey-svc:6379"

# Verify (values are base64-encoded, not shown in plain text)
kubectl get secret excel-ai-secrets -n excel-ai
```

## Step 12 — Configure Your IP in manifests

```bash
# On the Oracle VM, clone the repo (or scp the k8s/ directory)
git clone YOUR-REPO-URL
cd excel-ai-agent

ORACLE_IP=$(curl -s ifconfig.me)
echo "Your IP: $ORACLE_IP"

# Replace placeholder in configmap and ingress
sed -i "s/YOUR-ORACLE-IP/$ORACLE_IP/g" k8s/base/configmap.yaml
sed -i "s/YOUR-ORACLE-IP/$ORACLE_IP/g" k8s/base/ingress.yaml

# Verify
grep "nip.io" k8s/base/ingress.yaml
# Should show: host: 130.61.2.15.nip.io
```

## Step 13 — Deploy Everything

Apply resources in the correct order. The backend initContainer requires PostgreSQL to be ready before it can run `prisma migrate deploy`, so Postgres and Valkey must be running first.

```bash
# 1. ConfigMap (must exist before backend reads it)
kubectl apply -f k8s/base/configmap.yaml

# 2. PostgreSQL + Valkey
kubectl apply -f k8s/base/postgres.yaml
kubectl apply -f k8s/base/valkey.yaml

# 3. Wait — backend initContainer needs these ready
kubectl rollout status deployment/postgres -n excel-ai --timeout=120s
kubectl rollout status deployment/valkey   -n excel-ai --timeout=60s

# 4. Backend (initContainer runs prisma migrate, then NestJS starts)
kubectl apply -f k8s/base/backend.yaml
kubectl rollout status deployment/backend  -n excel-ai --timeout=180s

# 5. Frontend
kubectl apply -f k8s/base/frontend.yaml
kubectl rollout status deployment/frontend -n excel-ai --timeout=120s

# 6. Ingress (last — only expose traffic once everything is ready)
kubectl apply -f k8s/base/ingress.yaml
```

---

# PART 5 — VERIFICATION

## Step 14 — Check All Pods

```bash
kubectl get pods -n excel-ai
```

Expected output (all Running):
```
NAME                        READY   STATUS    RESTARTS   AGE
postgres-xxx                1/1     Running   0          2m
valkey-xxx                  1/1     Running   0          2m
backend-xxx                 1/1     Running   0          90s
frontend-xxx                1/1     Running   0          60s
```

## Step 15 — Check Services and PVCs

```bash
kubectl get svc -n excel-ai
# Expected:
# backend-svc    ClusterIP  10.x.x.x  <none>  3001/TCP
# frontend-svc   ClusterIP  10.x.x.x  <none>  3000/TCP
# postgres-svc   ClusterIP  10.x.x.x  <none>  5432/TCP
# valkey-svc     ClusterIP  10.x.x.x  <none>  6379/TCP

kubectl get pvc -n excel-ai
# Expected: uploads-pvc, postgres-pvc, valkey-pvc — all Bound

kubectl get ingress -n excel-ai
# Expected: ADDRESS = your Oracle IP
```

## Step 16 — Test Backend Health

```bash
curl http://YOUR-ORACLE-IP.nip.io/api/health
# Expected:
# {"status":"ok","service":"excel-ai-backend","timestamp":"..."}
```

## Step 17 — Open the Frontend

Open in your browser:
```
http://YOUR-ORACLE-IP.nip.io
```

You should see the Excel AI Agent ChatGPT-style interface.

## Step 18 — End-to-End Test

1. Click the **+** button → upload a `.xlsx` file
2. Type: `Show me the sheet structure`
3. Press Enter
4. Watch the Agent log fill in real-time
5. After completion → click **Download file**

---

# TROUBLESHOOTING

### Pod not starting — check events
```bash
kubectl describe pod -n excel-ai <pod-name>
kubectl get events -n excel-ai --sort-by='.lastTimestamp'
```

### Backend not connecting to Postgres
```bash
kubectl logs -n excel-ai deployment/backend -c wait-for-postgres
kubectl logs -n excel-ai deployment/backend -c prisma-migrate
kubectl logs -n excel-ai deployment/backend
```

### Check all logs
```bash
kubectl logs -f deployment/backend   -n excel-ai
kubectl logs -f deployment/frontend  -n excel-ai
kubectl logs -f deployment/postgres  -n excel-ai
kubectl logs -f deployment/valkey    -n excel-ai
```

### Restart a deployment
```bash
kubectl rollout restart deployment/backend -n excel-ai
kubectl rollout status  deployment/backend -n excel-ai
```

### Rollback to previous version
```bash
kubectl rollout undo deployment/backend -n excel-ai
```

### Check PVCs
```bash
kubectl get pvc -n excel-ai
kubectl describe pvc uploads-pvc -n excel-ai
```

### Check ingress
```bash
kubectl get ingress -n excel-ai
kubectl describe ingress excel-ai-ingress -n excel-ai
kubectl logs -n kube-system deployment/traefik
```

### Test Redis/Valkey from inside cluster
```bash
kubectl run redis-test --image=valkey/valkey:7-alpine \
  --restart=Never --rm -it -n excel-ai \
  -- valkey-cli -h valkey-svc ping
# Expected: PONG
```

### Test Postgres from inside cluster
```bash
kubectl run pg-test --image=postgres:16-alpine \
  --restart=Never --rm -it -n excel-ai \
  --env="PGPASSWORD=YOUR_DB_PASSWORD" \
  -- psql -h postgres-svc -U postgres -d excel_ai_db -c "SELECT 1"
# Expected: ?column? = 1
```

---

# UPDATING THE APPLICATION

## Build new image and update

On your **LOCAL machine**:
```bash
# Bump image tag or rebuild
make build-oracle ORACLE_IP=$ORACLE_IP IMAGE_TAG=v2
make push-oracle  ORACLE_IP=$ORACLE_IP IMAGE_TAG=v2

# On Oracle VM — update image reference and restart
kubectl set image deployment/backend \
  backend=excel-ai-backend:v2 -n excel-ai

kubectl rollout status deployment/backend -n excel-ai
```

## Quick restart (same image)
```bash
kubectl rollout restart deployment/backend  -n excel-ai
kubectl rollout restart deployment/frontend -n excel-ai
```

---

# PERSISTENT STORAGE — IMPORTANT

| PVC | Size | What it stores | If deleted |
|-----|------|----------------|-----------|
| `uploads-pvc` | 5Gi | Uploaded + output Excel files | All files lost permanently |
| `postgres-pvc` | 5Gi | All job history and metadata | Database lost permanently |
| `valkey-pvc` | 1Gi | Active job queue | In-flight jobs lost |

**Back up uploads-pvc before any maintenance:**
```bash
# Copy files out of the pod before deleting
kubectl cp excel-ai/$(kubectl get pod -n excel-ai -l app=backend \
  -o jsonpath='{.items[0].metadata.name}'):/app/uploads ./backup-uploads
```

---

# ENVIRONMENT VARIABLES REFERENCE

### Secrets (kubectl create secret)
| Key | Description |
|-----|-------------|
| `DATABASE_URL` | `postgresql://postgres:PASSWORD@postgres-svc:5432/excel_ai_db` |
| `AI_API_KEY` | Your Anthropic API key (`sk-ant-...`) |
| `POSTGRES_PASSWORD` | Random strong password |
| `VALKEY_URL` | `redis://valkey-svc:6379` |

### ConfigMap (k8s/base/configmap.yaml)
| Key | Value |
|-----|-------|
| `EXTERNAL_URL` | `http://YOUR-ORACLE-IP.nip.io` |
| `AI_MODEL` | `claude-sonnet-4-6` |
| `AI_BASE_URL` | `https://api.anthropic.com` |
| `FILE_RETENTION_INPUT_HOURS` | `24` |
| `FILE_RETENTION_OUTPUT_HOURS` | `2` |

### Build-time (Makefile)
| Variable | Used for |
|----------|----------|
| `ORACLE_IP` | nip.io domain + CORS_ORIGIN |
| `DOMAIN` | `NEXT_PUBLIC_API_URL` baked into frontend image |

`DOMAIN` and `EXTERNAL_URL` must be the same value.
