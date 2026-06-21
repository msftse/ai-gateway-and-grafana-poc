# 08b — Hosted agent, end to end (6 min)

**Goal:** Tie it together. Invoke a **hosted Foundry agent** — an Azure assistant
whose tool is the **Microsoft Learn MCP** server fronted by the gateway — and
show that the *same* gateway policies (managed-identity auth, token metrics,
cache, content safety, load balancing) apply to agent traffic too.

## Command

```bash
.inspect-venv/bin/python clients/06_invoke_agent.py
```

## What it does

Connects to the Foundry project (`PROJECT_ENDPOINT`) and the pre-created agent
**`agent-aigw-demo`** using `AIProjectClient`. It asks an Azure question that
should trigger the agent's **Microsoft Learn MCP** tool (routed through APIM) and
asserts the agent completes a turn (ideally grounded in Microsoft Learn docs).

> SDK notes (azure-ai-projects 2.x): use
> `AIProjectClient(endpoint, credential, allow_preview=True)`, then
> `project.get_openai_client(agent_name=...)` and `client.responses.create(...)`.
> The `model` passed to `responses.create` must be the agent's **model**
> (`gpt-4o-mini`), not the agent name.

## What you should see

```
-> Connect to project and target agent 'agent-aigw-demo'
-> Ask an Azure question that should trigger the Microsoft Learn MCP tool
  agent answer: "The Azure API Management policy that limits LLM tokens is llm-token-limit ..."
  [PASS] expected=agent completes a turn (and ideally uses its tool)  actual=answered=True, tool-evidence=True
```

## Show it in the dashboard

In the workbook (filtered to your Test Run ID): the agent's underlying model
calls appear in **Token consumption**, **Cost**, and **Live trace explorer** —
proving agent traffic is governed and metered identically to direct calls.

## Talking points

- Agents are just another client of the gateway — they inherit every policy.
- The tool the agent calls (Microsoft Learn MCP) is itself fronted by APIM for
  auth, quotas, and observability — model traffic *and* tool traffic on one
  gateway.

## Screenshot

`docs/assets/08b-agent-answer.png`

➡️ Next: [09-cost-view.md](09-cost-view.md)
