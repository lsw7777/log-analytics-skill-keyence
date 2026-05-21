---
name: interactive-query-tool
description: Query and filter Azure Log Analytics logs with precise conditions, displaying results in an interactive HTML viewer with client-side filtering, search, and column management. Supports multiple DCR_CL tables. Use when searching for specific users (by email, userId, UPN), investigating particular events, auditing specific activities, filtering logs by conditions (workload, operation, date, IP), exploring individual log entries in detail, or building ad-hoc log queries. Always use azure_log_query.ps1 to fetch data first, then embed results in an interactive HTML.
---

# Interactive Query Tool

Query, filter, and explore Azure Log Analytics logs with a self-contained interactive HTML viewer.

## Supported Tables

| Table Name | Description |
|-----------|-------------|
| `AuditGeneralDCR_CL` | Office 365 通用审计日志 |
| `SharePointAuditDCR_CL` | SharePoint 审计日志 |
| `MessageTraceDataDCR_CL` | 邮件追踪数据 |
| `AssignedLicensesDCR_CL` | 已分配许可证信息 |
| `AzureADUsersDCR_CL` | Azure AD 用户信息 |
| `MailboxStatisticsDCR_CL` | 邮箱统计信息 |
| `WQCLogDCR_CL` | WQC 日志 |

## Prerequisites

- `azure_log_query.ps1` script — the only data source
- Azure authentication via `Connect-AzAccount` (AzureChinaCloud)
- Workspace ID: `703a5771-97fc-4bf3-a585-f607d18c4479`

## Workflow

### Step 1: Probe Schema

**Always** probe the current schema first — fields change as the ingestion pipeline evolves.

```bash
# Probe AuditGeneralDCR_CL
pwsh ./azure_log_query.ps1 -TableName "AuditGeneralDCR_CL" -Query "AuditGeneralDCR_CL | take 1" -Hours 24

# Probe SharePointAuditDCR_CL
pwsh ./azure_log_query.ps1 -TableName "SharePointAuditDCR_CL" -Query "SharePointAuditDCR_CL | take 1" -Hours 24

# Probe MessageTraceDataDCR_CL
pwsh ./azure_log_query.ps1 -TableName "MessageTraceDataDCR_CL" -Query "MessageTraceDataDCR_CL | take 1" -Hours 24
```

Capture all available column names from the response. Use these to build filters dynamically.

### Step 2: Fetch Data with Conditions

Query logs with specific filters:

```bash
# AuditGeneralDCR_CL - All logs for a user
pwsh ./azure_log_query.ps1 -TableName "AuditGeneralDCR_CL" -Query "AuditGeneralDCR_CL | where UserUPN contains 'user@domain.com' | sort by TimeGenerated desc" -Hours 168

# AuditGeneralDCR_CL - Specific workload + operation
pwsh ./azure_log_query.ps1 -TableName "AuditGeneralDCR_CL" -Query "AuditGeneralDCR_CL | where Workload == 'PowerBI' and Operation contains 'Refresh' | sort by TimeGenerated desc" -Hours 72

# AuditGeneralDCR_CL - Failed operations only
pwsh ./azure_log_query.ps1 -TableName "AuditGeneralDCR_CL" -Query "AuditGeneralDCR_CL | where IsSuccess == false | sort by TimeGenerated desc" -Hours 24

# SharePointAuditDCR_CL - Specific site access
pwsh ./azure_log_query.ps1 -TableName "SharePointAuditDCR_CL" -Query "SharePointAuditDCR_CL | where SiteUrl contains 'sites/finance' | sort by TimeGenerated desc" -Hours 168

# MessageTraceDataDCR_CL - Failed emails
pwsh ./azure_log_query.ps1 -TableName "MessageTraceDataDCR_CL" -Query "MessageTraceDataDCR_CL | where Status == 'Failed' | sort by TimeGenerated desc" -Hours 24
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
