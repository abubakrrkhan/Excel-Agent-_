# ─────────────────────────────────────────────────────────────────────────────
# Makefile — Excel AI Agent
#
# ORACLE ALWAYS FREE QUICK START:
#
#   Step 1 — On your LOCAL machine, set your Oracle VM IP:
#     export ORACLE_IP=130.61.2.15         (your Oracle VM public IP)
#     export AI_API_KEY=sk-ant-YOUR_KEY
#
#   Step 2 — Edit k8s/base/configmap.yaml and ingress.yaml:
#     Replace YOUR-ORACLE-IP with $ORACLE_IP
#
#   Step 3 — Build images for ARM64 (Oracle Ampere A1):
#     make build-oracle ORACLE_IP=$ORACLE_IP
#
#   Step 4 — Copy images to Oracle VM and import into K3s:
#     make push-oracle ORACLE_IP=$ORACLE_IP
#
#   Step 5 — Create secrets and deploy on Oracle VM:
#     (done via SSH — see K3S_ORACLE_DEPLOYMENT.md)
#
# ─────────────────────────────────────────────────────────────────────────────

ORACLE_IP   ?= YOUR-ORACLE-IP
DOMAIN      ?= http://$(ORACLE_IP).nip.io
IMAGE_TAG   ?= latest
PLATFORM    ?= linux/arm64   # Oracle Ampere A1 is ARM64
NAMESPACE   = excel-ai

BACKEND_IMG  = excel-ai-backend:$(IMAGE_TAG)
FRONTEND_IMG = excel-ai-frontend:$(IMAGE_TAG)

.PHONY: help build-oracle build-backend build-frontend \
        push-oracle test type-check verify health \
        logs-backend logs-frontend up down

## ── Oracle Always Free (primary target) ────────────────────────────────────

build-oracle: build-backend build-frontend  ## Build ARM64 images for Oracle
	@echo ""
	@echo "✅ Images built for Oracle ARM64"
	@echo "   Backend:  $(BACKEND_IMG)"
	@echo "   Frontend: $(FRONTEND_IMG)"
	@echo ""
	@echo "Next: make push-oracle ORACLE_IP=$(ORACLE_IP)"

build-backend:  ## Build backend image (ARM64 for Oracle Ampere A1)
	@echo "Building backend for $(PLATFORM)..."
	docker buildx build \
		--platform $(PLATFORM) \
		--file docker/Dockerfile.backend \
		--tag $(BACKEND_IMG) \
		--load \
		.

build-frontend:  ## Build frontend image — DOMAIN is baked in at build time
	@echo "Building frontend for $(PLATFORM) — NEXT_PUBLIC_API_URL=$(DOMAIN)"
	docker buildx build \
		--platform $(PLATFORM) \
		--file docker/Dockerfile.frontend \
		--build-arg NEXT_PUBLIC_API_URL=$(DOMAIN) \
		--tag $(FRONTEND_IMG) \
		--load \
		.

push-oracle:  ## Save images as tarballs and scp to Oracle VM
	@echo "Saving images..."
	docker save $(BACKEND_IMG)  | gzip > /tmp/backend.tar.gz
	docker save $(FRONTEND_IMG) | gzip > /tmp/frontend.tar.gz
	@echo "Copying to Oracle VM $(ORACLE_IP)..."
	scp /tmp/backend.tar.gz  ubuntu@$(ORACLE_IP):/tmp/
	scp /tmp/frontend.tar.gz ubuntu@$(ORACLE_IP):/tmp/
	@echo "Importing into K3s on Oracle VM..."
	ssh ubuntu@$(ORACLE_IP) "sudo k3s ctr images import /tmp/backend.tar.gz && sudo k3s ctr images import /tmp/frontend.tar.gz && echo '✅ Images imported'"

## ── Local development (x86 Docker Compose) ─────────────────────────────────

up:   ## Start locally with Docker Compose (x86)
	docker compose up --build

down: ## Stop local services
	docker compose down

## ── Tests ───────────────────────────────────────────────────────────────────

test:       ## Run backend tests
	cd backend && npm test

type-check: ## TypeScript check (backend + frontend)
	cd backend  && npm run typecheck
	cd frontend && npm run type-check

## ── Verification (run from Oracle VM) ──────────────────────────────────────

verify: ## Wait for all deployments to be ready
	kubectl rollout status deployment/postgres  -n $(NAMESPACE) --timeout=120s
	kubectl rollout status deployment/valkey    -n $(NAMESPACE) --timeout=60s
	kubectl rollout status deployment/backend   -n $(NAMESPACE) --timeout=180s
	kubectl rollout status deployment/frontend  -n $(NAMESPACE) --timeout=120s
	@echo "✅ All deployments ready"

health: ## Test backend health endpoint from inside the cluster
	kubectl run curl-$(shell date +%s) --image=curlimages/curl:latest \
		--restart=Never --rm -it -n $(NAMESPACE) \
		-- curl -s http://backend-svc:3001/api/health

logs-backend:  ## Tail backend logs
	kubectl logs -f deployment/backend -n $(NAMESPACE)

logs-frontend: ## Tail frontend logs
	kubectl logs -f deployment/frontend -n $(NAMESPACE)

## ── Help ────────────────────────────────────────────────────────────────────

help: ## Show available commands
	@grep -E '^[a-zA-Z_-]+:.*?## .*$$' $(MAKEFILE_LIST) | \
		awk 'BEGIN {FS = ":.*?## "}; {printf "  \033[36m%-22s\033[0m %s\n", $$1, $$2}'

.DEFAULT_GOAL := help
