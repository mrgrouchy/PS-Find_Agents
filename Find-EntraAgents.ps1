#Requires -Version 7.2
<#
.SYNOPSIS
Inventories Microsoft Entra Agent ID identities using Microsoft Graph.
.DESCRIPTION
Reads every page of the v1.0 agentIdentity collection. Optional switches include
blueprint principals (infrastructure, not individual agents) and heuristic legacy
service principal candidates. Does not change any tenant objects.
.EXAMPLE
./Find-EntraAgents.ps1 -TenantId 'contoso.onmicrosoft.com' -CsvPath ./agents.csv
.EXAMPLE
./Find-EntraAgents.ps1 -TenantId '<tenant-guid>' -IncludeBlueprints -IncludeLegacyCandidates
.EXAMPLE
./Find-EntraAgents.ps1 -TenantId '<tenant-guid>' -UseExistingConnection
#>
[CmdletBinding()]
param(
    [ValidateNotNullOrEmpty()]
    [string]$TenantId,
    [switch]$IncludeBlueprints,
    [switch]$IncludeLegacyCandidates,
    [ValidateNotNullOrEmpty()]
    [string]$CandidatePattern = '(?i)agent|copilot|bot|foundry',
    [string]$CsvPath,
    [switch]$UseDeviceCode,
    [switch]$UseExistingConnection
)

$ErrorActionPreference = 'Stop'
if (-not $TenantId) {
    $configPath = Join-Path $PSScriptRoot 'config.local.json'
    if (Test-Path -LiteralPath $configPath) {
        $config = Get-Content -LiteralPath $configPath -Raw | ConvertFrom-Json
        $TenantId = $config.TenantId
    }
    if ([string]::IsNullOrWhiteSpace($TenantId)) {
        throw "Specify -TenantId or set TenantId in config.local.json (see config.example.json)."
    }
}
if (-not (Get-Module -ListAvailable Microsoft.Graph.Authentication)) {
    throw 'Install the dependency first: Install-Module Microsoft.Graph.Authentication -Scope CurrentUser'
}
Import-Module Microsoft.Graph.Authentication -ErrorAction Stop

# Compile before authentication so an invalid custom expression fails immediately.
$candidateRegex = [regex]::new($CandidatePattern)
$scopes = @('AgentIdentity.Read.All')
if ($IncludeBlueprints) { $scopes += 'AgentIdentityBlueprintPrincipal.Read.All' }
if ($IncludeLegacyCandidates) { $scopes += 'Application.Read.All' }

if (-not $UseExistingConnection) {
    $connect = @{
        TenantId = $TenantId
        Scopes = $scopes
        ContextScope = 'Process'
        NoWelcome = $true
        Environment = 'Global'
    }
    if ($UseDeviceCode) { $connect.UseDeviceCode = $true }
    Connect-MgGraph @connect | Out-Null
}
$context = Get-MgContext
if (-not $context) { throw 'No Graph connection. Connect-MgGraph first or omit -UseExistingConnection.' }
if ($context.Environment -ne 'Global') { throw 'This script currently supports the Global Graph environment only.' }
# A GUID is required when reusing a session so its tenant can be checked unambiguously.
if ($UseExistingConnection) {
    $parsedTenant = [guid]::Empty
    if (-not [guid]::TryParse($TenantId, [ref]$parsedTenant)) {
        throw '-UseExistingConnection requires the tenant GUID rather than its domain name.'
    }
    if ([guid]$context.TenantId -ne $parsedTenant) { throw 'The connected tenant differs from -TenantId.' }
}

function Get-GraphCollection {
    param([Parameter(Mandatory)][string]$Uri)
    do {
        # Invoke-MgGraphRequest uses the Graph SDK retry handler for throttling.
        try {
            $page = Invoke-MgGraphRequest -Method GET -Uri $Uri -OutputType Hashtable -ErrorAction Stop
        }
        catch {
            throw "Graph inventory failed at '$Uri'. Check admin consent, Agent ID roles, and API availability. No complete report was produced. $($_.Exception.Message)"
        }
        if (-not $page.ContainsKey('value')) { throw "Unexpected collection response from '$Uri'." }
        foreach ($item in $page['value']) { $item }
        $Uri = $page['@odata.nextLink']
    } while ($Uri)
}

function New-InventoryRow {
    param($Item, [string]$Kind, [string]$Evidence)
    [pscustomobject][ordered]@{
        TenantId = $context.TenantId
        Kind = $Kind
        DisplayName = $Item['displayName']
        ObjectId = $Item['id']
        AppId = $Item['appId']
        BlueprintAppId = $Item['agentIdentityBlueprintId']
        AccountEnabled = $Item['accountEnabled']
        CreatedDateTime = $Item['createdDateTime']
        ServicePrincipalType = $Item['servicePrincipalType']
        Tags = @($Item['tags']) -join ';'
        Evidence = $Evidence
    }
}

$baseUri = 'https://graph.microsoft.com/v1.0'
$rows = [System.Collections.Generic.List[object]]::new()
$knownIds = [System.Collections.Generic.HashSet[string]]::new()
$identities = @(Get-GraphCollection "$baseUri/servicePrincipals/microsoft.graph.agentIdentity")
foreach ($item in $identities) {
    [void]$knownIds.Add($item['id'])
    $rows.Add((New-InventoryRow $item 'AgentIdentity' 'Graph agentIdentity object type'))
}

if ($IncludeBlueprints) {
    foreach ($item in (Get-GraphCollection "$baseUri/servicePrincipals/microsoft.graph.agentIdentityBlueprintPrincipal")) {
        [void]$knownIds.Add($item['id'])
        $rows.Add((New-InventoryRow $item 'BlueprintPrincipal' 'Graph agentIdentityBlueprintPrincipal object type; not an individual agent'))
    }
}

if ($IncludeLegacyCandidates) {
    Write-Warning 'Legacy candidates are name/tag matches, not confirmed agents. This scan can miss agents and include unrelated applications.'
    $uri = $baseUri + '/servicePrincipals?$select=id,appId,displayName,accountEnabled,createdDateTime,servicePrincipalType,tags&$top=100'
    foreach ($item in (Get-GraphCollection $uri)) {
        if ($knownIds.Contains($item['id'])) { continue }
        # Exclude blueprint infrastructure even when -IncludeBlueprints was not requested.
        if ($item['@odata.type'] -eq '#microsoft.graph.agentIdentityBlueprintPrincipal') { continue }
        $searchText = @($item['displayName']) + @($item['tags']) -join ' '
        if ($candidateRegex.IsMatch($searchText)) {
            $rows.Add((New-InventoryRow $item 'LegacyCandidate' "Display name or tags match regex: $CandidatePattern"))
        }
    }
}

$results = @($rows | Sort-Object Kind, DisplayName, ObjectId)
if ($CsvPath) {
    if ($results.Count -gt 0) {
        $results | Export-Csv -LiteralPath $CsvPath -NoTypeInformation -Encoding utf8
    }
    else {
        # Replace any old report with a header-only file for a successful empty inventory.
        $emptyItem = @{}
        $header = (New-InventoryRow $emptyItem '' '' | ConvertTo-Csv -NoTypeInformation)[0]
        Set-Content -LiteralPath $CsvPath -Value $header -Encoding utf8
    }
    Write-Information "CSV saved to $CsvPath" -InformationAction Continue
}
Write-Information "Found $($identities.Count) confirmed Agent ID identities; $($results.Count) total report rows." -InformationAction Continue
Write-Warning 'Coverage: Entra Agent ID identities only, plus any requested blueprint principals and heuristic candidates. Agents without an Entra Agent ID identity are not comprehensively discoverable through this query.'
$results
