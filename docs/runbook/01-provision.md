# 01 — Provision the lab

*Do this before the meeting. `terraform apply` takes ~15–25 minutes (APIM
Premium v2 is the long pole).*

## Steps

### 1. Initialize and review

```bash
cd infrastructure
cp terraform.tfvars.example terraform.tfvars   # first run only; edit if needed
terraform init
terraform plan -out tfplan
```

Review the plan: ~40+ resources across the resource group, Log Analytics + App
Insights, Azure Managed Redis, two Foundry accounts with three model
deployments, Content Safety, APIM Premium v2 with the AI-gateway policies,
products/subscriptions, the backend pool, RBAC role assignments, and the
workbook.

### 2. Apply

```bash
terraform apply tfplan
```

A post-apply bootstrap step (`scripts/bootstrap-agent.py`) creates the hosted
Foundry agent **`agent-aigw-demo`** — an Azure assistant whose Microsoft Learn
MCP tool is fronted by the gateway.

### 3. Capture outputs into the test env

The apply writes [clients/.env](../../clients/.env) automatically. Confirm it:

```bash
cd ..
grep -E "APIM_GATEWAY_URL|PROJECT_ENDPOINT|AGENT_NAME|WORKBOOK_PORTAL_URL" clients/.env
```

### 4. Sanity-check the deployed resources

```bash
az resource list -g rg-aigw-demo -o table
```

## What you should see

- `terraform apply` ends with `Apply complete!` and an outputs block including
  `apim_gateway_url`, `foundry_project_endpoint`, `workbook_portal_url`, and the
  (sensitive) subscription keys.
- `clients/.env` is populated with live values.
- The agent bootstrap prints an agent id like `agent-aigw-demo:1`.

> **Naming:** resources use a random suffix (e.g. `apim-aigw-7enh2q`). The exact
> names are in the Terraform outputs and `clients/.env`; the runbook refers to
> them via those variables so it stays correct across re-provisions.

## Screenshot

`docs/assets/01-provision-apply-complete.png`

➡️ Next: [02-foundry-models.md](02-foundry-models.md)
