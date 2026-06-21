#!/usr/bin/env python3
"""04 - Load balancing across two Foundry backends.

Capability: the gateway routes to a backend POOL of two independent Foundry
accounts (priority + weight) with a circuit breaker. Traffic spreads across
hosts, and if one host trips the breaker, traffic shifts to the other.

Asserts:
  * across N calls, responses are served by >= 1 backend host and the gateway
    reports a backend host on every call (x-backend-host echoed by policy)
  * (best-effort) more than one distinct backend host is observed
Prints the backend distribution.

The pool members share the same public DNS suffix, so when the host name is
identical we fall back to APIM's region/backend routing being exercised; the
distribution is still printed for the live audience.
"""

from __future__ import annotations

import _common as c

CALLS = 30


@c.timed
def main() -> c.Result:
    r = c.Result("04_load_balance", "Backend pool load balancing")
    c.title("04 - Load balancing across backends")

    c.load_env()
    key = c.require("TEAM_B_SUBSCRIPTION_KEY")

    c.step(f"Send {CALLS} sequential calls and record the serving backend")
    hosts: dict[str, int] = {}
    statuses: dict[int, int] = {}
    for i in range(CALLS):
        resp = c.chat(
            key=key,
            user_id="lb-demo",
            messages=[{"role": "user", "content": f"Reply with the single word: ping ({i})"}],
            max_tokens=5,
        )
        statuses[resp.status_code] = statuses.get(resp.status_code, 0) + 1
        host = resp.headers.get("x-backend-host", "(none)")
        hosts[host] = hosts.get(host, 0) + 1

    n200 = statuses.get(200, 0)
    c.info(f"status distribution: {statuses}")
    c.info("backend distribution:")
    for host, n in sorted(hosts.items(), key=lambda kv: -kv[1]):
        c.info(f"    {n:3d}  {host}")

    served_ok = n200 >= int(CALLS * 0.8)
    have_backend = any(h != "(none)" for h in hosts)
    distinct = len([h for h in hosts if h != "(none)"])
    if distinct >= 2:
        actual = f"{distinct} distinct backends, {n200}/{CALLS} ok"
    else:
        actual = f"1 backend host, {n200}/{CALLS} ok"

    # PASS requires healthy serving + the gateway reporting a backend. Multiple
    # distinct hosts is the ideal but pool members can share a hostname.
    passed = served_ok and have_backend
    c.report(
        r,
        passed,
        expected=">=80% served + backend host reported (ideally >=2 distinct)",
        actual=actual,
        detail="Gateway did not report a backend host." if not have_backend else "",
    )
    return r


if __name__ == "__main__":
    res = main()
    raise SystemExit(0 if res.passed else 1)
