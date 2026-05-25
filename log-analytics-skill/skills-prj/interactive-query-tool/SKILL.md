---
name: interactive-query-tool
description: Query and filter Azure Log Analytics Office365 audit logs with precise conditions, displaying results in an interactive HTML viewer with client-side filtering, search, and column management. Use when searching for specific users (by email, userId, UPN), investigating particular events, auditing specific activities, filtering logs by conditions (workload, operation, date, IP), exploring individual log entries in detail, or building ad-hoc log queries. Always use azure_log_query.ps1 to fetch data first, then embed results in an interactive HTML.
---

# Interactive Query Tool

Query, filter, and explore Office365 Audit logs with a self-contained interactive HTML viewer.

## Prerequisites

- `azure_log_query.ps1` script in repo root — the only data source
- Azure authentication via `Connect-AzAccount` (AzureChinaCloud)
- Workspace ID: `703a5771-97fc-4bf3-a585-f607d18c4479`

## Workflow

### Step 1: Probe Schema

**Always** probe the current schema first — fields change as the ingestion pipeline evolves.

```bash
pwsh ./azure_log_query.ps1 -Query "AuditGeneralDCR_CL | take 1" -Hours 24
```

Capture all available column names from the response. Use these to build filters dynamically.

### Step 2: Fetch Data with Conditions

Query logs with specific filters:

```bash
# All logs for a user
pwsh ./azure_log_query.ps1 -Query "AuditGeneralDCR_CL | where UserUPN contains 'user@domain.com' | sort by TimeGenerated desc" -Hours 168

# Specific workload + operation
pwsh ./azure_log_query.ps1 -Query "AuditGeneralDCR_CL | where Workload == 'PowerBI' and Operation contains 'Refresh' | sort by TimeGenerated desc" -Hours 72

# Failed operations only
pwsh ./azure_log_query.ps1 -Query "AuditGeneralDCR_CL | where IsSuccess == false | sort by TimeGenerated desc" -Hours 24

# Specific IP address
pwsh ./azure_log_query.ps1 -Query "AuditGeneralDCR_CL | where ClientIP contains '192.168' | sort by TimeGenerated desc" -Hours 168

# Time window + user + workload combined
pwsh ./azure_log_query.ps1 -Query "AuditGeneralDCR_CL | where TimeGenerated > datetime('2026-05-06') and UserUPN contains 'user@' and Workload == 'SharePoint' | sort by TimeGenerated desc" -Hours 96
```

**Common filter patterns:**
- `UserUPN contains 'email@'` — filter by user email (most human-readable)
- `UserId contains 'xxx'` — filter by GUID
- `Workload == 'PowerBI'` — exact workload match (values: `PowerBI`, `Exchange`, `SharePoint`, `OneDrive`, etc.)
- `Operation contains 'keyword'` — partial operation match
- `ClientIP contains 'xxx'` — filter by IP prefix
- `IsSuccess == true/false` — success/failure filter
- `TimeGenerated > datetime('YYYY-MM-DD')` — time filter (use alongside `-Hours` parameter)
- `Activity contains 'xxx'` — filter by activity name

### Step 3: Build Interactive HTML Viewer

Create a **self-contained** HTML file with all data embedded inline.

**Core features:**
1. **Search bar** — free-text search across all fields
2. **Column filter dropdowns** — one per column, populated with unique values from the dataset
3. **Date range picker** — filter by CreationTime/TimeGenerated range
4. **Column management** — show/hide columns (important for 100+ column tables)
5. **Sortable table** — click headers to sort ascending/descending
6. **Pagination** — customizable page size (25/50/100 per page)
7. **Row detail modal** — click a row to see all fields in key-value layout
8. **Export button** — copy visible data as CSV

**Data embedding:**
Embed the fetched JSON data directly in a `<script>` block:
```html
<script>
const AUDIT_DATA = [/* JSON array of all fetched records */];
const FIELD_ORDER = ["UserUPN","UserId","Workload","Operation","TimeGenerated","CreationTime","ClientIP","IsSuccess","Activity","RecordType"];
</script>
```

**Technical requirements:**
- Single `.html` file — self-contained, no external CDN or network requests
- All CSS inline or in `<style>` block
- All JavaScript in `<script>` block
- Use modern vanilla JS only (no jQuery, no frameworks)
- Dark theme by default
- Responsive layout
- Filter state should update in real-time without page reload

**UI Layout:**
```
┌──────────────────────────────────────────────────────────────┐
│  Header: "Audit Log Explorer — N records, Time Range"       │
├──────────────────────────────────────────────────────────────┤
│  🔍 Search  │  Workload ▼  │  Operation ▼  │  User ▼  │ 📅   │
├──────────────────────────────────────────────────────────────┤
│  Column Manager: ☑ UserUPN  ☑ Workload  ☑ Operation ... ▼   │
├──────────────────────────────────────────────────────────────┤
│  ┌─────────┬──────────┬──────────┬──────────┬──────────┐     │
│  │ UserUPN │ Workload │ Operation│ TimeGene│ ClientIP │  ↓  │
│  ├─────────┼──────────┼──────────┼──────────┼──────────┤     │
│  │ user@   │ PowerBI  │ Refresh  │ 05-07   │ 10.x.x.x │  ↓  │
│  │ user2@  │ Exchange │ Send     │ 05-07   │ 10.x.x.x │  ↓  │
│  └─────────┴──────────┴──────────┴──────────┴──────────┘     │
├──────────────────────────────────────────────────────────────┤
│  ← Page 1 of 12 │ 25 per page ▼                    Export 📋 │
└──────────────────────────────────────────────────────────────┘
```

**Column ordering strategy:**
With 100+ columns, display a curated set by default. Order columns in this priority:

1. Primary identifiers: `UserUPN`, `UserId`, `Workload`, `Operation`, `Activity`, `TimeGenerated`, `CreationTime`, `ClientIP`, `IsSuccess`
2. Common metadata: `RecordType`, `OrganizationId`, `TenantId`
3. Workload-specific fields (detect from data):
   - PowerBI: `DatasetName`, `ReportName`, `WorkspaceName`
   - Exchange: `MailboxGuid`, `Subject`, `MessageId`
   - SharePoint/OneDrive: `SourceFileName`, `SiteUrl`, `ObjectSource`
4. All remaining fields — available in row detail modal and column manager

### Step 4: Verify Output

- Open the HTML file and test:
  - Search functionality works across all fields
  - Filter dropdowns populate with correct unique values
  - Date range filter works
  - Column show/hide works
  - Sorting works on all visible columns
  - Pagination works correctly
  - Row detail modal shows all fields
  - Export produces valid CSV
- Confirm file size is reasonable (< 50 MB for typical datasets)
- Verify no external network requests (check browser dev tools Network tab)

## Edge Cases

- **Empty results**: Show friendly "no matching records" message with clear filters button
- **Large datasets** (>5000 rows): Add a warning that rendering may be slow; offer to aggregate first via KQL before fetching
- **Null/empty fields**: Show as `-` rather than blank space; exclude columns where all values are null
- **Long text fields**: Truncate in table view with `...`, show full value in detail modal
- **Special characters in KQL**: Escape single quotes with `''`, wrap strings in quotes
