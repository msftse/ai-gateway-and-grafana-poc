# 00 — Prerequisites

*Do this before the meeting. ~15 minutes plus provisioning time.*

## What you need

| Requirement | Detail |
|---|---|
| Azure subscription | Owner or Contributor + User Access Administrator (the lab assigns RBAC roles to APIM's managed identity). |
| Region quota | `swedencentral` with capacity for **APIM Premium v2**, **Azure Managed Redis (Balanced B1)**, and Foundry **gpt-4o-mini** + **text-embedding-3-small** deployments. |
| Azure CLI | `az version` ≥ 2.60. |
| Terraform | `terraform version` ≥ 1.9. |
| Python | The project venv at `.inspect-venv/` (created from `clients/requirements.txt`). |

## Steps

### 1. Sign in

```bash
az login
az account set --subscription "<your-subscription-id>"
az account show -o table
```

### 2. Confirm Foundry / OpenAI quota in the region

```bash
az cognitiveservices usage list --location swedencentral \
  --query "[?contains(name.value, 'gpt-4o-mini') || contains(name.value, 'text-embedding-3-small')]" \
  -o table
```

Expect non-zero available capacity for both model families.

### 3. Create the Python venv for the test scripts

```bash
python3 -m venv .inspect-venv
.inspect-venv/bin/python -m pip install --upgrade pip
.inspect-venv/bin/python -m pip install -r clients/requirements.txt
```

> The system `python3` (3.9.x on macOS) cannot import `azure.ai.projects`. Always
> invoke the tests with `.inspect-venv/bin/python`.

### 4. Verify the toolchain

```bash
terraform version
az version
.inspect-venv/bin/python -c "import azure.ai.projects, openai; print('SDK OK')"
```

## What you should see

- `az account show` prints the intended subscription.
- Quota query returns rows with available capacity.
- `SDK OK` printed by the last command.

## Screenshot

`docs/assets/00-prereqs-az-account.png`

➡️ Next: [01-provision.md](01-provision.md)
