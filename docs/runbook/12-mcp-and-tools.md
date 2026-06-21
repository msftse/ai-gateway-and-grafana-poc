# 12 — MCP & tools: add a capability live (8 min)

**The finale.** Two parts: (a) explain the APIM-as-MCP-gateway architecture, then
(b) **add a new tool to the already-running agent in real time** — proving the
agent is governed, traced, and metered through the gateway *and* can gain new
capabilities live.

---

## Part A — APIM in front of tools & MCP (architecture, ~3 min)

Show the diagram (deck slide "Agent tools, MCPs & custom MCPs through the
gateway") and make three points:

1. **REST-as-MCP** — API Management can expose an existing REST API as an **MCP
   server**, so any MCP-capable agent can call it with central auth, quotas, and
   logging.
2. **MCP passthrough** — APIM can front a remote MCP server, applying the same
   AI-gateway policies (rate limit, content safety, metrics) to tool traffic.
   **This lab provisions one live:** the public Microsoft Learn MCP server is
   exposed through APIM (`type = "mcp"`, streamable HTTP) and the demo agent
   calls it through the gateway.
3. **Custom MCP** — a Container App / Function App MCP server can sit behind APIM
   for governance. (Reference architecture only in this lab — no custom MCP is
   provisioned.)

Tie back: the agent's **model** calls already traverse the gateway; its **tool**
calls can too. One control plane for models *and* tools.

---

## Part B — The Microsoft Learn MCP tool, fronted by the gateway (~5 min)

The agent `agent-aigw-demo` is an **Azure assistant** whose tool is the
**Microsoft Learn MCP** server — and that MCP server is exposed *through APIM*
(MCP passthrough, `type = "mcp"`, streamable HTTP). It was wired in during
provisioning, so the agent's **model** traffic *and* its **MCP tool** traffic
both traverse the same gateway. Foundry forbids sensitive headers on MCP tools,
so the APIM subscription key is stored in a **Foundry project connection**
(Custom keys, key `api-key`) and the tool references it via
`project_connection_id`; the service injects the `api-key` header on each call.

The MCP endpoint lives at `MCP_SERVER_URL` in `clients/.env`:

```
https://<apim>.azure-api.net/ms-learn-mcp
```

### Republish the tool live (versioning demo)

Show the "change a running, governed agent live" moment by publishing a new
agent version with the (gateway-routed) MCP tool:

```bash
.inspect-venv/bin/python scripts/add-mcp-tool.py
```

Point it at a different MCP server to attach another one:

```bash
.inspect-venv/bin/python scripts/add-mcp-tool.py \
  --server-url https://learn.microsoft.com/api/mcp \
  --label microsoft_learn \
  --description "Search and fetch official Microsoft Learn documentation."
```

### What you should see

```
Published MCP tool 'microsoft_learn' (https://<apim>.azure-api.net/ms-learn-mcp) [via APIM gateway].
Agent 'agent-aigw-demo' is now at version 2 (id=agent-aigw-demo:2).
Tools on this version: microsoft_learn (Microsoft Learn MCP).
```

### Confirm in the Foundry portal

1. [ai.azure.com](https://ai.azure.com) → project `proj-aigw` → **Agents** →
   `agent-aigw-demo`.
2. Show the **Microsoft Learn MCP** tool whose server URL points at the **APIM
   gateway** (`/ms-learn-mcp`), not directly at learn.microsoft.com, and the
   **Custom keys** project connection (`learn-mcp-gateway`) that holds the
   `api-key`.
3. In the APIM instance → **APIs → MCP Servers**, show the **Microsoft Learn
   MCP** server (the passthrough) governed by the gateway.

### Exercise the tool

Run the agent test and ask something the MCP tool can answer:

```bash
.inspect-venv/bin/python clients/06_invoke_agent.py
```

Or, in the Foundry **Agents playground**, ask: *"Use Microsoft Learn to find the
APIM policy that limits LLM tokens."* The agent calls the MCP tool **through the
gateway** and answers from live docs.

### Show governance still applies

In the workbook (filtered to your Test Run ID): the agent's model calls appear in
**Token consumption** / **Cost** / **Live trace explorer**, and the MCP tool
calls traverse APIM (visible in **ApiManagementGatewayLogs** / App Insights for
the `ms-learn-mcp` API) — model *and* tool traffic governed and observable from
one control plane.

> **Approval flow (optional):** the script sets `require_approval="never"` for an
> unattended demo. Change it to `"always"` to show the human-in-the-loop tool
> approval step before the agent may call the MCP server.

## Talking points

- The agent's **tool** traffic is governed by the **same gateway** as its model
  traffic — Microsoft Learn MCP is fronted by APIM, not called directly.
- Versioning means the change is **auditable and reversible** — roll back
  instantly.
- Models, agents, and MCP servers all sit behind **one** gateway.

## Screenshot

`docs/assets/12-agent-v2-two-tools.png`

➡️ Next: [99-cleanup.md](99-cleanup.md)
