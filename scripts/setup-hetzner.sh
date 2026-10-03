#!/bin/bash
# ─────────────────────────────────────────────────────────────────────────────
# setup-hetzner.sh — Run this ONCE on a fresh Hetzner VPS (Ubuntu 22.04)
#
# Usage:
#   ssh root@YOUR-HETZNER-IP "bash -s" < scripts/setup-hetzner.sh
#
# What it does:
#   1. Installs K3s (Kubernetes)
#   2. Installs nginx ingress controller
#   3. Installs kubectl alias
#   4. Creates the excel-ai namespace
#   5. Shows next steps
# ─────────────────────────────────────────────────────────────────────────────

set -e

GREEN='\033[0;32m'
YELLOW='\033[1;33m'
NC='\033[0m'

log() { echo -e "${GREEN}[setup]${NC} $1"; }
warn() { echo -e "${YELLOW}[warn]${NC} $1"; }

log "Starting Hetzner K3s setup..."

# ── 1. Install K3s ────────────────────────────────────────────────────────────
log "Installing K3s..."
curl -sfL https://get.k3s.io | INSTALL_K3S_EXEC="--disable traefik" sh -
sleep 10

# Make kubectl available without sudo
mkdir -p ~/.kube
cp /etc/rancher/k3s/k3s.yaml ~/.kube/config
chmod 600 ~/.kube/config
export KUBECONFIG=~/.kube/config

# Also set up for the user running this script
echo 'export KUBECONFIG=~/.kube/config' >> ~/.bashrc

log "K3s installed. Node status:"
kubectl get nodes

# ── 2. Install nginx ingress ──────────────────────────────────────────────────
log "Installing nginx ingress controller..."
kubectl apply -f https://raw.githubusercontent.com/kubernetes/ingress-nginx/controller-v1.10.0/deploy/static/provider/baremetal/deploy.yaml

log "Waiting for nginx ingress to be ready..."
kubectl wait --namespace ingress-nginx \
  --for=condition=ready pod \
  --selector=app.kubernetes.io/component=controller \
  --timeout=120s

log "Nginx ingress ready."

# ── 3. Patch nginx to use host ports (needed for bare metal) ─────────────────
kubectl patch svc ingress-nginx-controller \
  -n ingress-nginx \
  -p '{"spec":{"type":"NodePort","externalTrafficPolicy":"Local"}}'

# ── 4. Create namespace ───────────────────────────────────────────────────────
kubectl create namespace excel-ai --dry-run=client -o yaml | kubectl apply -f -

# ── 5. Get server IP ──────────────────────────────────────────────────────────
SERVER_IP=$(curl -s ifconfig.me || hostname -I | awk '{print $1}')

log "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
log "Hetzner K3s setup complete!"
log ""
log "Server IP: ${SERVER_IP}"
log ""
warn "NEXT: Run the deploy script on your LOCAL machine:"
warn ""
warn "  export HETZNER_IP=${SERVER_IP}"
warn "  export AI_API_KEY=sk-ant-YOUR_KEY"
warn "  export DB_PASSWORD=\$(openssl rand -base64 24)"
warn "  bash scripts/deploy.sh"
log "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
