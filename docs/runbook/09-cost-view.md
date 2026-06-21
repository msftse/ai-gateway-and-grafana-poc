# 09 — View & manage costs (5 min)

**Goal:** Turn token telemetry into **chargeback/showback**. Show per-team,
per-subscription, per-user cost from the `llm-emit-token-metric` data — no
guessing from invoices.

## Portal click-path

1. Open the **Workbook** (link printed by `run_all.py`, or `WORKBOOK_PORTAL_URL`
   in `clients/.env`).
2. Set the **Test Run ID** parameter to this run (or clear it for all traffic).
3. **Executive summary** tiles — Total tokens, **estimated cost (USD)**, avg
   latency, error rate, cache-hit %, throttled requests.
4. **Cost view** section:
   - Cost broken down by **Product**, **Subscription**, **User ID**, and API
     operation.
   - The editable **price table** parameter (input/output price per 1K tokens per
     model) — change a number and watch the cost recompute live.

## How the cost is derived

The base policy emits `llm-emit-token-metric` (namespace `ai-gateway`) with the
prompt/completion/cached token split and the dimensions API ID, Subscription ID,
Product, User ID, Test Run ID. The workbook multiplies those token counts by the
price table to produce USD cost per any dimension.

## Talking points

- **Showback** (visibility) and **chargeback** (billing) come straight from
  gateway metrics — no per-app instrumentation.
- Filtering by Product/Subscription gives true **per-team** cost.
- Cached tokens are tracked separately, so you can quantify the **savings** the
  semantic cache delivers.

## CLI (optional — raw metric)

```bash
az monitor metrics list \
  --resource "$(az apim show -g rg-aigw-demo -n apim-aigw-7enh2q --query id -o tsv)" \
  --namespace "ai-gateway" --metric "Total Tokens" -o table
```

## Screenshot

`docs/assets/09-cost-by-team.png`

➡️ Next: [10-traces-logs.md](10-traces-logs.md)
