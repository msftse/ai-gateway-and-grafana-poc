# AGENT.md — APIM AI Gateway + Microsoft Foundry Demo

Internal project tracker for a reproducible, in-person customer demo of **Azure API Management as an AI Gateway in front of Microsoft Foundry**. This file records architecture, decisions, conventions, and a changelog. User-facing setup/usage lives in [README.md](README.md).

---

## Summary

A Terraform-provisioned lab that puts API Management (Premium v2) in front of two Microsoft Foundry accounts, demonstrating the full AI-gateway feature set — managed-identity auth, per-team token rate limits and daily quotas, semantic (vector) caching, content safety, backend load balancing, and end-to-end observability — plus a hosted Foundry agent (an Azure assistant) whose Microsoft Learn MCP tool is fronted by the gateway. Ships with an asserting PASS/FAIL test suite, a 12-section Azure Monitor workbook, a presenter runbook, and a standalone slide deck.

---

## Phases (delivery plan)

| Phase | Scope | Status |
|-------|-------|--------|
| 1 | Scaffold — repo layout, Terraform skeleton, policy files | ✅ Complete |
| 2 | Live provisioning — `terraform apply`, agent bootstrap | ✅ Complete |
| 3 | Test suite (6 tests), 16-file runbook, MCP tool script | ✅ Complete |
| 4 | Standalone PowerPoint deck (26 slides) | ✅ Complete |
| 5 | Documentation — AGENT.md + README.md | ✅ Complete |

---

## Architecture

### Provisioned resources (Terraform modules)

| Module | Key resources | Purpose |
|--------|---------------|---------|
| `observability` | Log Analytics workspace, Application Insights | Telemetry sink for gateway logs + custom token metrics |
| `redis` | Azure Managed Redis | Backing store for the semantic cache |
| `content-safety` | Azure AI Content Safety account | Prompt moderation backend for the `llm-content-safety` policy |
| `foundry` | 2× AI Services (Foundry) accounts, project, model deployments | `gpt-4o-mini` (chat, both accounts) + `text-embedding-3-small` (account 1) |
| `apim` | API Management (Premium v2, system-assigned MI), Foundry LLM API, products, subscriptions, backends, named values, diagnostic | The AI gateway itself |
| `workbook` | Azure Monitor workbook | 12-section dashboard (tokens, cost, latency, errors, cache, safety, traces) |

A post-apply `terraform_data.bootstrap_agent` trigger runs [scripts/bootstrap-agent.py](scripts/bootstrap-agent.py) to create the hosted demo agent (an Azure assistant) and its Microsoft Learn MCP tool, fronted by the gateway.

### Request flow

```
client / agent ──▶ APIM (AI gateway) ──▶ Foundry model deployments
                      │                        (chat + embeddings)
                      ├─ managed-identity auth
                      ├─ llm-token-limit (TPM + daily quota, per product)
                      ├─ llm-semantic-cache (vector lookup via Redis + embeddings)
                      ├─ llm-content-safety (prompt moderation)
                      ├─ backend pool (load balance + circuit breaker)
                      └─ llm-emit-token-metric ──▶ Application Insights / Log Analytics
```

### Telemetry model (critical, non-obvious)

- The **custom token metric** (`Total/Prompt/Completion Tokens`) carries dimensions `Test Run ID`, `User ID`, `Product`, `Subscription ID` — these are the slice axes for cost/usage panels.
- The APIM App Insights diagnostic emits a **fixed customDimensions schema** on `requests`/`dependencies` (API Name, Operation Name, Product Name, Subscription Name, Cache, Request Id, Region, Service ID/Name/Type, API Revision, HTTP Method, API Type). It does **not** include arbitrary headers by default.
- To filter request telemetry by a custom header, the header must be added to the diagnostic's `frontend.request.headers` / `backend.request.headers`; APIM then logs it **prefixed** as `Request-<header>` (e.g. `Request-x-test-run-id`, `Request-x-user-id`).
- `customDimensions['Cache']` on `requests` ∈ {`Hit`, `Miss`, `None`} is the **authoritative** per-request semantic-cache outcome (`None` = throttled/blocked requests that never reached cache lookup).
- Dependencies do **not** reliably carry frontend request headers — dependency-based panels must not filter by run id.

---

## Key decisions

- **APIM tier = Premium v2 (`PremiumV2_1`).** v2 SKU is required for the AI gateway (LLM) policies and MCP features. Kept public (no VNet) for demo simplicity.
- **Region = `swedencentral`** — broad availability for Foundry `gpt-4o-mini` + `text-embedding-3-small`, APIM v2, and Azure Managed Redis.
- **Two Foundry accounts** so the load-balancer / circuit-breaker demo has genuinely distinct backends.
- **Semantic cache ON** via Azure Managed Redis + `text-embedding-3-small`; score threshold `0.2` (cosine distance) to catch paraphrases without false matches.
- **Token-limit demo tiers:** product `dev-team-a` (low) at 500 TPM / 10k daily quota so the 429/403 demo triggers fast; `dev-team-b` (high) at 5000 TPM.
- **Demo agent is pre-created** during provisioning so every gateway capability can be tested end-to-end before the live "add a tool" segment.
- **Custom MCP = reference only** — no deployable MCP server code; the runbook describes fronting a Container App / Function App MCP behind APIM. [scripts/add-mcp-tool.py](scripts/add-mcp-tool.py) attaches a public Microsoft Learn MCP to the agent live.
- **Local Terraform state** (no remote backend) — single-presenter lab, torn down after the session.
- **Deck is standalone/general** — no lab-specific resource names, scripts, or screenshots, so it can be reused as a product presentation.

### Decisions grounded in live diagnosis

- **Embeddings backend URL must end `/embeddings`.** The `llm-semantic-cache-lookup` policy POSTs directly to the embeddings backend; a missing `/embeddings` segment 404s and silently skips caching.
- **`llm-token-limit` at product scope is restricted** — only `counter-key`, `tokens-per-minute`, `token-quota`, `token-quota-period`, `retry-after-*` are allowed; `estimate-prompt-tokens` must be `false`. Output-token headers/vars are reserved for API/operation scope.
- **APIM diagnostic setting needs `log_analytics_destination_type = "Dedicated"`** for the `ApiManagementGatewayLogs` resource-specific table (otherwise logs land only in the legacy `AzureDiagnostics` table).
- **Workbook `tiles` visualization needs one row per tile** — unpivot a single multi-column summary row with `| evaluate narrow()`; `$column` columnMatch mode does not render.

---

## Conventions

- **Naming:** `<resource>-<name_prefix>-<suffix>` where `name_prefix` defaults to `aigw` and `suffix` is a random string (e.g. `apim-aigw-7enh2q`).
- **Tags:** every resource carries `SecurityControl=Ignore` (always added automatically), plus `Project`, `Owner`, `Environment`.
- **Python interpreter:** test/bootstrap scripts require the project venv at `.inspect-venv/bin/python` — the system `python3` cannot import the Azure SDK. The venv is gitignored.
- **Gateway config for clients** is written to `clients/.env` by `terraform output` during provisioning; the test scripts read it.
- **Test correlation:** each `run_all.py` run shares one `x-test-run-id` (`demo-<hex>`); paste it into the workbook's "Test Run ID" parameter to scope the dashboard to that run.
- **IaC** under `infrastructure/` per the `terraform-builder` skill. Provider versions: azurerm 4.77, azapi 2.10, random 3.9 (Terraform ≥ 1.x).

---

## Project structure

```
glassbox-session/
├── AGENT.md                          # internal tracker (this file)
├── README.md                         # presenter-facing public docs
├── .gitignore
├── infrastructure/
│   ├── main.tf  variables.tf  outputs.tf  providers.tf
│   ├── terraform.tfvars.example
│   ├── policies/
│   │   ├── policy-api.xml             # API-scope AI gateway policy
│   │   ├── policy-product-high.xml    # dev-team-b limits
│   │   └── policy-product-low.xml     # dev-team-a limits
│   └── modules/
│       ├── observability/            # Log Analytics + App Insights
│       ├── redis/                    # Azure Managed Redis (semantic cache)
│       ├── content-safety/           # Azure AI Content Safety
│       ├── foundry/                  # 2× Foundry accounts, project, deployments
│       ├── apim/                     # API Management + AI gateway config
│       └── workbook/                 # workbook.json + tf wrapper
├── clients/
│   ├── requirements.txt  .env.example  _common.py
│   ├── run_all.py                    # orchestrator: PASS/FAIL summary + Test Run ID
│   └── 01_chat_basic.py … 06_invoke_agent.py
├── scripts/
│   ├── bootstrap-agent.py            # post-apply: create agent + Microsoft Learn MCP tool
│   └── add-mcp-tool.py               # live finale: attach an MCP tool to the agent
├── docs/
│   ├── runbook/                      # README + 00-prereqs … 99-cleanup (16 files)
│   └── assets/                       # screenshot placeholders
└── deck/
    ├── build.js                      # pptxgenjs build script
    ├── package.json  package-lock.json
    └── apim-ai-gateway-foundry.pptx  # generated deck (26 slides)
```

---

## Changelog

### 2026-06-14 — Initial documentation
- Created `AGENT.md` (internal tracker) and `README.md` (presenter docs).
- Recorded architecture, telemetry model, key decisions, and conventions for the completed Phases 1–5.

### 2026-06-14 — Phase 4: standalone deck
- Built [deck/build.js](deck/build.js) (pptxgenjs) and generated `deck/apim-ai-gateway-foundry.pptx` — 26 slides covering all 7 demo topics, Midnight Executive palette, standalone/general content (no lab specifics).
- Added local deck dependencies (`pptxgenjs`, `react`/`react-dom`, `react-icons`, `sharp`) under `deck/`.
- QA: validated slide XML + visual render; fixed slide 22 icon contrast (azure → white on navy circles).

### 2026-06-14 — Phase 3: tests, runbook, workbook fixes
- Built asserting test suite in `clients/` (01 chat, 02 token-limit, 03 semantic cache, 04 load balance, 05 content safety, 06 agent) + `run_all.py` orchestrator.
- Wrote 16-file presenter runbook under `docs/runbook/`.
- Added [scripts/add-mcp-tool.py](scripts/add-mcp-tool.py) for the live "add a tool" finale (attaches Microsoft Learn MCP to the agent).
- Fixed semantic cache (embeddings backend URL `/embeddings`); rewrote test 03 to assert deterministically on shared response id.
- Workbook: fixed run-filter telemetry keys (`Request-x-test-run-id` on requests; `Test Run ID` on custom metrics), rewired cache panels to `customDimensions['Cache']`, fixed tile rendering with `evaluate narrow()`, switched APIM diagnostic to Dedicated table mode + added header logging.

### 2026-06-14 — Phases 1–2: scaffold + live provisioning
- Authored Terraform root + 6 modules and 3 AI-gateway policy files.
- Applied live (42 resources); bootstrapped hosted agent `agent-aigw-demo`; captured outputs to `clients/.env`; deployed the workbook.
- Corrected product-scope `llm-token-limit` attributes and the `bootstrap-agent.py` SDK call (typed `PromptAgentDefinition`, keyword-only `definition`).
