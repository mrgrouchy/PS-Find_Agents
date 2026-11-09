# Find Entra agents

Read-only Microsoft Graph PowerShell inventory for Entra Agent ID identities.
Requires PowerShell 7.2+ and Microsoft.Graph.Authentication. Targets the Global cloud.

```powershell
Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
Copy-Item ./config.example.json ./config.local.json
# Edit config.local.json and enter your tenant GUID or domain.
./Find-EntraAgents.ps1 -CsvPath ./agents.csv
```

The default tenant is read from `config.local.json` beside the script.
Override it when needed with `-TenantId '<another-tenant-id>'`.
The local config and generated CSV reports are ignored by Git.
The script returns PowerShell objects and optionally writes a CSV. For an on-screen table:

```powershell
./Find-EntraAgents.ps1 |
    Format-Table Kind, DisplayName, ObjectId, AccountEnabled, BlueprintAppId -AutoSize
```

Use `-UseDeviceCode` for device sign-in. Interactive authentication uses a process-scoped
Graph session. The script does not install modules or grant consent automatically.

## Permissions

Have an administrator consent to the read permissions needed for your options:

| Option | Microsoft Graph permission |
| --- | --- |
| Default inventory | `AgentIdentity.Read.All` |
| `-IncludeBlueprints` | Additionally `AgentIdentityBlueprintPrincipal.Read.All` |
| `-IncludeLegacyCandidates` | Additionally `Application.Read.All` |

Microsoft documents **Agent ID Administrator** as the least privileged supported role
for nonowners using the delegated agent list APIs. Use an appropriately authorized
account for tenant-wide results; owning a blueprint alone is not proof of tenant-wide access.

For unattended use, establish an app-only `Connect-MgGraph` session using a certificate
or managed identity with the corresponding application permissions and admin consent,
then pass `-UseExistingConnection -TenantId '<tenant-guid>'`. Existing sessions must
target the Global cloud and the same tenant. API authorization is enforced by Graph.

## Coverage

- `AgentIdentity` rows are confirmed typed Entra Agent ID objects, including disabled identities.
- `-IncludeBlueprints` adds tenant blueprint principals as `BlueprintPrincipal` rows.
  These are infrastructure and should not be counted as individual agents.
- `-IncludeLegacyCandidates` scans all service principals and flags display names/tags
  matching `agent|copilot|bot|foundry` (case insensitive). These `LegacyCandidate` rows
  require review. Customize with `-CandidatePattern`. They can include false positives
  and miss agents whose metadata does not match.
- This is not a complete inventory of every Copilot Studio, Foundry, third-party,
  Agent Registry, or unregistered agent. Older agents can use ordinary service principals;
  other agents have no dedicated directory identity. Inventory those platforms separately.
- All collection pages are followed. Query errors terminate before export. If a previous
  CSV exists after a failed run, it remains an old report; do not treat it as current.
  A successful empty scan writes a header-only CSV. Export paths must have an existing parent directory.

## Microsoft references

- [List agent identities](https://learn.microsoft.com/en-us/graph/api/agentidentity-list?view=graph-rest-1.0)
- [List blueprint principals](https://learn.microsoft.com/en-us/graph/api/agentidentityblueprintprincipal-list?view=graph-rest-1.0)
- [Agent identities and ordinary service principals](https://learn.microsoft.com/en-us/entra/agent-id/agent-service-principals)

The script uses the documented v1.0 endpoints. No name heuristic is treated as proof
that an application is an agent.
