# Azure API Management as an AI Gateway for Microsoft Foundry

A reproducible, in-person demo lab that puts **Azure API Management (Premium v2)** in front of **Microsoft Foundry** to govern every AI call — security, per-team token limits and quotas, semantic caching, content safety, load balancing, and full observability — plus a hosted Foundry agent (an **Azure assistant**) whose **Microsoft Learn MCP** tool is fronted by the gateway.

It ships with everything needed to run a polished customer session:

- **Terraform** that provisions the whole environment in one `apply`.
- An **asserting PASS/FAIL test suite** that exercises every gateway capability.
- A **12-section Azure Monitor workbook** dashboard (tokens, cost, latency, errors, cache, safety, traces).
- A **presenter runbook** with portal click-paths, commands, and expected output.
- A **standalone slide deck** (`deck/apim-ai-gateway-foundry.pptx`).

> Internal architecture, decisions, and changelog live in [AGENT.md](AGENT.md).

---

## Quick start

### Prerequisites

| Requirement | Notes |
|-------------|-------|
| Azure subscription | Owner/Contributor on a resource group; quota for APIM Premium v2, Foundry `gpt-4o-mini` + `text-embedding-3-small`, and Azure Managed Redis in your region |
| Azure CLI | `az login` before provisioning |
| Terraform | ≥ 1.x |
| Python | 3.12 recommended; create a venv and install `clients/requirements.txt` |
| Node.js | Only needed to rebuild the deck (see [Rebuilding the deck](#rebuilding-the-deck)) |

### 1. Configure

```bash
cd infrastructure
cp terraform.tfvars.example terraform.tfvars
# edit terraform.tfvars — at minimum set apim_publisher_email
```

### 2. Provision

```bash
az login
terraform init
terraform apply
```

This creates the resource group, observability, Redis, Content Safety, two Foundry accounts with model deployments, API Management with the AI-gateway configuration, the workbook dashboard, and (by default) a hosted demo agent. Gateway values are written to `clients/.env` automatically.

### 3. Set up the Python client

```bash
python3 -m venv .inspect-venv
.inspect-venv/bin/python -m pip install -r clients/requirements.txt
```

> The test scripts require this venv — the system `python3` cannot import the Azure SDK.

### 4. Run the test suite

```bash
.inspect-venv/bin/python clients/run_all.py
```

You get an all-green PASS/FAIL table, a shared **Test Run ID**, and a deep link to the dashboard. Paste the Test Run ID into the workbook's **Test Run ID** parameter to scope every panel to that run.

### 5. Tear down

```bash
cd infrastructure
terraform destroy
```

---

## What gets deployed

| Component | Resource | Role |
|-----------|----------|------|
| AI gateway | API Management (Premium v2, system-assigned identity) | Single front door for all AI traffic |
| Models | 2× Foundry accounts — `gpt-4o-mini` (chat) + `text-embedding-3-small` (embeddings) | Backends, with a second account for load balancing |
| Semantic cache | Azure Managed Redis | Vector response cache |
| Content safety | Azure AI Content Safety | Prompt moderation backend |
| Observability | Log Analytics + Application Insights | Token metrics, gateway logs, traces |
| Dashboard | Azure Monitor workbook | 12-section operations view |
| Agent | Hosted Foundry agent (Azure assistant) + Microsoft Learn MCP via APIM | End-to-end agent demo |

### Gateway capabilities demonstrated

- Keyless managed-identity auth from APIM to Foundry
- `llm-token-limit` — per-team tokens-per-minute and daily token quota (429 / 403)
- `llm-semantic-cache-*` — vector response caching
- `llm-content-safety` — central prompt moderation
- Backend pool load balancing + circuit breaker
- `llm-emit-token-metric` — token metering with custom dimensions
- Agent tools and MCP servers brokered through the gateway

---

## Configuration

Set these in `infrastructure/terraform.tfvars` (all optional except `apim_publisher_email`):

| Variable | Default | Description |
|----------|---------|-------------|
| `apim_publisher_email` | — | **Required.** Publisher email for API Management |
| `location` | `swedencentral` | Region (must support APIM v2, the Foundry models, and Managed Redis) |
| `name_prefix` | `aigw` | 2–9 char lowercase prefix for all resource names |
| `environment` | `demo` | Environment tag |
| `owner` | `unset` | Owner tag |
| `apim_sku_name` | `PremiumV2_1` | v2 SKU required for AI-gateway LLM policies |
| `chat_model_name` | `gpt-4o-mini` | Chat model deployed on both accounts |
| `embedding_model_name` | `text-embedding-3-small` | Embedding model for the cache |
| `team_a_tokens_per_minute` | `500` | TPM limit for `dev-team-a` (low tier) |
| `team_a_token_quota` | `10000` | Daily token quota for `dev-team-a` |
| `team_b_tokens_per_minute` | `5000` | TPM limit for `dev-team-b` (high tier) |
| `content_safety_block_threshold` | `4` | Severity (0–7) at/above which content is blocked |
| `semantic_cache_score_threshold` | `0.2` | Cosine-distance threshold for cache matches |
| `run_bootstrap_agent` | `true` | Create the demo agent + initial tool after apply |
| `tags` | `{}` | Extra tags (`SecurityControl=Ignore` is always added) |

> `terraform.tfvars` is gitignored. Every resource is tagged `SecurityControl=Ignore` automatically.

---

## Scripts & CLI usage

| Command | Purpose |
|---------|---------|
| `.inspect-venv/bin/python clients/run_all.py` | Run all 6 tests, print PASS/FAIL + Test Run ID + dashboard link |
| `.inspect-venv/bin/python clients/01_chat_basic.py` | Single chat completion through the gateway |
| `.inspect-venv/bin/python clients/02_load_test.py` | Token rate limit + daily quota (429/403) |
| `.inspect-venv/bin/python clients/03_semantic_cache.py` | Semantic (vector) cache hit |
| `.inspect-venv/bin/python clients/04_load_balance.py` | Backend pool load balancing |
| `.inspect-venv/bin/python clients/05_content_safety.py` | Prompt content-safety screening |
| `.inspect-venv/bin/python clients/06_invoke_agent.py` | Hosted agent (Azure assistant) + Microsoft Learn MCP (via APIM) |
| `.inspect-venv/bin/python scripts/add-mcp-tool.py` | Live finale: (re)publish the agent's Microsoft Learn MCP tool (via APIM) |

---

## Presenting the demo

Follow the runbook in order — each file has a portal click-path, the exact command, and the expected "money shot":

➡️ **[docs/runbook/README.md](docs/runbook/README.md)** — order of play, timings, and 16 step files (`00-prereqs` … `99-cleanup`).

The slide deck for the conceptual portion is at **`deck/apim-ai-gateway-foundry.pptx`** (general/standalone — no lab specifics, reusable as a product presentation).

### Rebuilding the deck

```bash
cd deck
npm install
node build.js          # regenerates apim-ai-gateway-foundry.pptx
```

---

## Project structure

```
glassbox-session/
├── AGENT.md                          # internal tracker
├── README.md                         # this file
├── infrastructure/
│   ├── main.tf  variables.tf  outputs.tf  providers.tf
│   ├── terraform.tfvars.example
│   ├── policies/                     # 3 AI-gateway policy files
│   └── modules/
│       ├── observability/            # Log Analytics + App Insights
│       ├── redis/                    # Azure Managed Redis (semantic cache)
│       ├── content-safety/           # Azure AI Content Safety
│       ├── foundry/                  # 2× Foundry accounts, project, deployments
│       ├── apim/                     # API Management + AI gateway config
│       └── workbook/                 # workbook.json + tf wrapper
├── clients/
│   ├── requirements.txt  .env.example  _common.py
│   ├── run_all.py                    # test orchestrator
│   └── 01_chat_basic.py … 06_invoke_agent.py
├── scripts/
│   ├── bootstrap-agent.py            # post-apply: create agent + Microsoft Learn MCP tool
│   └── add-mcp-tool.py               # live finale: attach an MCP tool
├── docs/
│   ├── runbook/                      # presenter runbook (16 files)
│   └── assets/                       # screenshot placeholders
└── deck/
    ├── build.js                      # pptxgenjs build script
    ├── package.json  package-lock.json
    └── apim-ai-gateway-foundry.pptx  # generated deck (26 slides)
```

---

## References

- [API Management — AI gateway capabilities](https://learn.microsoft.com/azure/api-management/genai-gateway-capabilities)
- [API Management GenAI policy reference](https://learn.microsoft.com/azure/api-management/api-management-policies#ai-gateway)
- [Microsoft Foundry documentation](https://learn.microsoft.com/azure/ai-foundry/)
- [Model Context Protocol (MCP)](https://modelcontextprotocol.io/)
- [Azure Monitor workbooks](https://learn.microsoft.com/azure/azure-monitor/visualize/workbooks-overview)
