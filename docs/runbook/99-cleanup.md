# 99 — Cleanup

*Run after the session to remove all billed resources (APIM Premium v2 and Azure
Managed Redis are the meaningful standing costs).*

## Tear down everything

```bash
cd infrastructure
terraform destroy
```

Review the destruction plan and confirm. This removes the resource group's
contents created by this lab: APIM, both Foundry accounts + deployments, Redis,
Content Safety, Log Analytics, App Insights, the workbook, products,
subscriptions, backends, and RBAC role assignments.

## Verify nothing is left

```bash
az resource list -g rg-aigw-demo -o table
az group show -n rg-aigw-demo -o table   # may still exist if managed elsewhere
```

If the resource group was created by Terraform, `destroy` removes it too; if it
pre-existed, delete it manually only if you own it:

```bash
az group delete -n rg-aigw-demo --yes --no-wait
```

## Local cleanup (optional)

```bash
cd ..
rm -f clients/.env            # contains live subscription keys
rm -rf .bootstrap             # agent bootstrap status
```

> The `.inspect-venv/` virtualenv can be kept for the next run or removed with
> `rm -rf .inspect-venv`.

## Notes

- **Soft-deleted Cognitive Services:** Foundry/Content Safety accounts may enter
  a soft-deleted state. If you need to immediately recreate with the same names,
  purge them:
  ```bash
  az cognitiveservices account list-deleted -o table
  az cognitiveservices account purge \
    --location swedencentral -g rg-aigw-demo -n <account-name>
  ```
- **APIM deletion** can take several minutes to fully release the name.

✅ Demo complete.
