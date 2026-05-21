---
name: bulk-log-analyzer
description: Analyze bulk Azure Log Analytics logs and generate comprehensive HTML reports. Supports multiple DCR_CL tables. Data is exported to a temporary CSV file first to handle large datasets efficiently, then analyzed locally. Use when analyzing Office365 audit logs, viewing log statistics, generating log reports, or exploring activity patterns across time periods, users, or operations. Triggers for bulk log analysis, activity summaries, statistics generation, compliance reports, or any request to view multiple logs together.
---

# Bulk Log Analyzer

Analyze Azure Log Analytics logs at scale and generate interactive, self-contained HTML reports.

## Supported Tables

| Table Name | Short Name | Description |
|-----------|-----------|-------------|
| `AuditGeneralDCR_CL` | `General` | Office 365 通用审计日志 |
| `SharePointAuditDCR_CL` | `SPAudit` | SharePoint 审计日志 |
| `MessageTraceDataDCR_CL` | `MsgTrace` | 邮件追踪数据 |
| `AssignedLicensesDCR_CL` | `Licenses` | 已分配许可证信息 |
| `AzureADUsersDCR_CL` | `AADUsers` | Azure AD 用户信息 |
| `MailboxStatisticsDCR_CL` | `Mailbox` | 邮箱统计信息 |
| `WQCLogDCR_CL` | `WQC` | WQC 日志 |

**Usage:** Add `-TableName <name>` parameter to `azure_log_query.ps1`, or use `-Query` for full KQL control.

## Prerequisites

- `azure_log_query.ps1` script in repo root — the only data source
- Azure authentication via `Connect-AzAccount` (AzureChinaCloud)
- Workspace ID: `703a5771-97fc-4bf3-a585-f607d18c4479`
- Temp directory: `~\AppData\Local\Temp\opencode` for staging data files

## Workflow

### Step 1: Probe Schema

**Always** probe the current schema before writing analysis code, as fields evolve over time.

```bash
# Probe AuditGeneralDCR_CL
powershell -File ".\azure_log_query.ps1" -TableName "AuditGeneralDCR_CL" -Query "AuditGeneralDCR_CL | take 1" -Hours 48

# Probe SharePointAuditDCR_CL
powershell -File ".\azure_log_query.ps1" -TableName "SharePointAuditDCR_CL" -Query "SharePointAuditDCR_CL | take 1" -Hours 48

# Probe MessageTraceDataDCR_CL
powershell -File ".\azure_log_query.ps1" -TableName "MessageTraceDataDCR_CL" -Query "MessageTraceDataDCR_CL | take 1" -Hours 48
```

This returns a single record with all available columns. Note the field structure — fields vary by table type.

### Step 2: Export Data to Staging File

**Data may be large**, so always export to a temporary staging file first, then analyze the file rather than re-querying Azure.

**Naming Convention:**
- Format: `{ShortName}_YYYYMMDD.csv`
- Always include `YYYYMMDD` (the date the user asked to analyze, not the export date)
- The file is placed in `$env:USERPROFILE\AppData\Local\Temp\opencode\`
- Examples:
  - `General_20260507.csv` — 用户要求分析 2026-05-07 的 AuditGeneral 数据
  - `SPAudit_20260507.csv` — 用户要求分析 2026-05-07 的 SharePointAudit 数据
  - `MsgTrace_20260507.csv` — 用户要求分析 2026-05-07 的 MessageTrace 数据
  - `General_202605_last7d.csv` — 最后 7 天

**Always use `-ExportCsv` in `azure_log_query.ps1`:**
```bash
# AuditGeneralDCR_CL - Yesterday's data
powershell -File ".\azure_log_query.ps1" -TableName "AuditGeneralDCR_CL" -Query "AuditGeneralDCR_CL | where TimeGenerated >= datetime(2026-05-07T00:00:00Z) and TimeGenerated < datetime(2026-05-08T00:00:00Z) | sort by TimeGenerated desc" -Hours 48 -ExportCsv "$env:USERPROFILE\AppData\Local\Temp\opencode\General_20260507.csv"

# SharePointAuditDCR_CL - Last 24 hours
powershell -File ".\azure_log_query.ps1" -TableName "SharePointAuditDCR_CL" -Query "SharePointAuditDCR_CL | sort by TimeGenerated desc" -Hours 24 -ExportCsv "$env:USERPROFILE\AppData\Local\Temp\opencode\SPAudit_20260508.csv"

# MessageTraceDataDCR_CL - Custom KQL filter
powershell -File ".\azure_log_query.ps1" -TableName "MessageTraceDataDCR_CL" -Query "MessageTraceDataDCR_CL | where Status == 'Failed' | sort by TimeGenerated desc" -Hours 168 -ExportCsv "$env:USERPROFILE\AppData\Local\Temp\opencode\MsgTrace_202605_failed.csv"

# Using -TableName shorthand (auto-builds query)
powershell -File ".\azure_log_query.ps1" -TableName "AuditGeneralDCR_CL" -Hours 24 -ExportCsv "$env:USERPROFILE\AppData\Local\Temp\opencode\General_20260508.csv"
```

**KQL tips:**
- Use `| sort by TimeGenerated desc` for most recent first
- Add `| where Workload == "PowerBI"` to filter by workload
- Add `| take N` to limit result count (e.g., `| take 1000`)
- Add `| where UserId contains "example@domain.com"` to filter by user
- For very large datasets, consider filtering in KQL before export to reduce file size
- Use explicit `datetime()` range filters instead of `-Hours` for precise date coverage

**Data staging benefits:**
- Avoids repeated Azure queries (saves time and reduces API load)
- Enables re-analysis without re-authentication
- Large datasets can be chunked or paginated across multiple runs
- Analysis can proceed even if Azure connectivity is lost

**File naming rule:** Every time `azure_log_query.ps1` is executed, the output MUST be exported with `-ExportCsv` using a date-stamped filename. Never run the script without exporting.

### Step 3: Load and Analyze from Staging File

Read the exported CSV file and compute statistics **from the file**, not from Azure, :

```bash
# Example: Load the staging file and check row count
powershell -Command "(Import-Csv -Path '$env:USERPROFILE\AppData\Local\Temp\opencode\General_20260507.csv' -Encoding UTF8).Count"
```

All subsequent analysis (aggregation, group-by, risk detection) reads from this CSV file using PowerShell `Import-Csv`.

**Core Metrics:**
- Total event count
- Time range covered from user-specified period (use the date the user asked to analyze, NOT derived from data TimeGenerated; display as "查询时间段: YYYY-MM-DD")
- Unique users count
- Unique operations count
- Workload distribution

**Top N Tables (N=10 default):**
- Top users by activity count
- Top operations by frequency
- Top workloads by activity
- Top ClientIPs by request count
- Activity timeline (events per hour/day)
- Success/failure ratio (IsSuccess field)

**Risk & Security Metrics (Mandatory — always compute):**
- **IsSuccess analysis**: count of true/false/empty — investigate events with `IsSuccess == false` (failed operations) as potential security concerns
- **Suspicious IPs**: external IPs (non-RFC1918, non-10.x) accessing admin endpoints or exporting data; flag IPs accessing across multiple workloads
- **Off-hours activity (00:00-07:00 local time)**: users active outside business hours, especially ExportReport/SensitivityLabel operations
- **Service account activity**: GUID-formatted UserIds like `00000009-*` — distinguish automated scheduled vs suspicious trigger
- **High-privilege operations**: ExportReport, Search, EditDataset, Delete events — rank users performing these
- **Sensitive data events**: SensitivityLabeledFileOpened/Renamed, IrmContent — flag unusual patterns
- **Failed access attempts**: `IsSuccess == false` combined with operations like ViewReport, ExportReport, EditDataset — indicate unauthorized access attempts
- **Admin/Management events**: Workload == SecurityComplianceCenter, PowerPlatform admin actions
- **IP velocity**: single IP associated with more than 5 distinct users — possible shared proxy/VPN

If the dataset is too large for a single file, split across multiple files (e.g., `logs_part1.csv`, `logs_part2.csv`) and aggregate results across all files.

### Step 4: Generate HTML Report

Create a **self-contained** HTML report using the **standard template** below. The template structure is fixed — always use this layout to ensure reports are consistent across runs. Only the data and section visibility change based on findings.

**Standard HTML Template Structure:**

```html
<div class="navbar">
  <!-- Left: Title -->
  <!-- Right: Language toggle (中文 | 日本語) -->
</div>
<div class="header">
  <!-- Title, subtitle -->
  <!-- Meta tags: 查询时间范围 (from user-specified time, NOT data TimeGenerated) -->
  <!-- Example: "查询时间段：2026-05-07" or "Query Period: 2026-05-07" -->
  <!-- Total Records, Source file info -->
</div>
<div class="section" id="glossary-section">
  <!-- Operation Glossary: table mapping raw operation values to Chinese/Japanese explanations -->
  <!-- Collapsible by default — "Show Glossary" button -->
  <!-- Structure: Operation (EN) | 说明 (CN) | 説明 (JP) | Count -->
</div>
<div class="summary-grid">
  <!-- 6 cards: Total Events, Unique Users, Unique Ops, Workloads, Success, Failed -->
  <!-- All labels in current language via data-i18n attributes -->
</div>
<div class="section">
  <!-- Activity Timeline (horizontal bar chart) -->
</div>
<div class="section">
  <!-- Workload Distribution (SVG donut chart) -->
</div>
<div class="section">
  <!-- Top 15 Users (bar chart) -->
</div>
<div class="section">
  <!-- Top 15 Operations (bar chart) — each operation name has inline tooltip on hover showing glossary definition -->
</div>
<div class="section">
  <!-- Top 10 Client IPs (bar chart) -->
</div>
<div class="section">
  <!-- Success/Failure Ratio (segmented bar + legend) -->
</div>
<div class="section" id="risk-section">
  <!-- Risk Analysis (conditional — render if any findings exist) -->
  <!-- Sub-sections: Failed Operations, Suspicious IPs, Off-hours Activity,
       High-privilege Operations, Sensitive Data Events, IP Velocity -->
</div>
<div class="section">
  <!-- Detailed Data Table with pagination (preview - first 500 rows) -->
  <!-- Operation column cells show tooltip with glossary explanation on hover -->
</div>
```

**Technical Requirements (fixed across all reports):**
- Single `.html` file — self-contained, no external network requests
- CSS variables defined in `:root` for consistent theming — use the same color palette:
  ```css
  --bg-primary: #0d1117; --bg-secondary: #161b22; --bg-tertiary: #21262d;
  --border: #30363d; --text-primary: #e6edf3; --text-secondary: #8b949e;
  --accent: #58a6ff; --accent-green: #3fb950; --accent-red: #f85149;
  --accent-yellow: #d29922; --accent-purple: #bc8cff; --accent-orange: #f0883e;
  --accent-cyan: #39d2c0;
  ```
- Chart colors: `#58a6ff, #3fb950, #bc8cff, #f0883e, #f85149, #39d2c0` (cycle through)
- Font: `-apple-system, BlinkMacSystemFont, 'Segoe UI', 'Noto Sans SC', 'Noto Sans JP', Helvetica, Arial, sans-serif`
- Border radius: `8px`, padding: `24px` body
- Charts: inline CSS/SVG only — no external libraries
  - Bar charts: flexbox bars with percentage widths
  - Donut charts: SVG `<circle>` with `stroke-dasharray`
  - Timeline: horizontal bars
- Tables: sticky headers, clickable column sorting, client-side pagination (50 rows/page)
- Responsive layout using `grid` and `flexbox`
- Max width: `1400px`, centered

**Language Toggle (i18n) Implementation:**
- Default language: Chinese (中文)
- HTML uses `data-i18n="key"` attributes on all UI text elements
- Embed a JavaScript i18n dictionary object with two language sets:
  ```javascript
  const i18n = {
    zh: {
      "totalEvents": "总事件数", "uniqueUsers": "唯一用户", "uniqueOps": "唯一操作",
      "workloads": "工作负载", "success": "成功", "failed": "失败",
      "activityTimeline": "活动时间线", "workloadDist": "工作负载分布",
      "topUsers": "活跃用户排行", "topOps": "操作类型排行",
      "topIPs": "客户端 IP 排行", "successRatio": "成功/失败比率",
      "riskAnalysis": "风险分析", "detailedTable": "详细数据",
      "showGlossary": "显示术语表", "hideGlossary": "隐藏术语表",
      "metric": "指标", "value": "值", "severity": "严重程度",
      "unknown": "未知", "previous": "上一页", "next": "下一页",
      "riskIndicators": "个风险指标已检出",
    },
    ja: {
      "totalEvents": "総イベント数", "uniqueUsers": "ユニークユーザー", "uniqueOps": "ユニーク操作",
      "workloads": "ワークロード", "success": "成功", "failed": "失敗",
      "activityTimeline": "アクティビティタイムライン", "workloadDist": "ワークロード分布",
      "topUsers": "アクティブユーザーランキング", "topOps": "操作タイプランキング",
      "topIPs": "クライアント IP ランキング", "successRatio": "成功/失敗比率",
      "riskAnalysis": "リスク分析", "detailedTable": "詳細データ",
      "showGlossary": "用語集を表示", "hideGlossary": "用語集を非表示",
      "metric": "指標", "value": "値", "severity": "重要度",
      "unknown": "不明", "previous": "前へ", "next": "次へ",
      "riskIndicators": "件のリスク指標が検出されました",
    }
  };
  ```
- Toggle button switches `currentLang` variable and updates all `data-i18n` elements:
  ```javascript
  function switchLang(lang) {
    currentLang = lang;
    document.querySelectorAll('[data-i18n]').forEach(el => {
      el.textContent = i18n[lang][el.getAttribute('data-i18n')] || el.textContent;
    });
    refreshGlossary(); // glossary switches language too
  }
  ```

**Operation Glossary Implementation:**
- Embed an `operationsGlossary` object mapping raw operation names to explanations:
  ```javascript
  const operationsGlossary = {
    "ViewReport":        { zh: "查看 PowerBI 报表", ja: "PowerBI レポートを閲覧", count: 3223 },
    "GetWorkspaces":     { zh: "获取工作区列表", ja: "ワークスペース一覧を取得", count: 930 },
    "RefreshDataset":    { zh: "刷新数据集", ja: "データセットを更新", count: 593 },
    "ExportReport":      { zh: "导出报表 (有风险)", ja: "レポートをエクスポート (リスクあり)", count: 173 },
    "EditDataset":       { zh: "编辑数据集", ja: "データセットを編集", count: 36 },
    "Search":            { zh: "执行搜索", ja: "検索を実行", count: 56 },
    "Import":            { zh: "导入内容", ja: "コンテンツをインポート", count: 41 },
    "MessageSend":       { zh: "发送Teams消息", ja: "Teamsメッセージを送信", count: 36 },
    "SensitivityLabeledFileOpened":       { zh: "打开敏感标签文件", ja: "機密ラベル付きファイルを開く", count: 522 },
    "SensitivityLabeledFileRenamed":      { zh: "重命名敏感标签文件", ja: "機密ラベル付きファイルの名前を変更", count: 35 },
    "MessageReadReceiptReceived": { zh: "收到已读回执", ja: "既読確認を受信", count: 565 },
    "GetSnapshots":      { zh: "获取快照", ja: "スナップショットを取得", count: 67 },
    "RunEmailSubscription": { zh: "运行邮件订阅", ja: "メールサブスクリプションを実行", count: 41 },
    "ApiEndpointCallEvent": { zh: "API端点调用", ja: "APIエンドポイント呼び出し", count: 28 },
  };
  ```
- **Only include operations that are present in the current dataset** — build the glossary dynamically from actual data
- Glossary section is collapsible (hidden by default, show/hide toggle button)
- Every operation cell in tables and charts gets a tooltip showing the glossary definition on hover
- Tooltip style: small floating box with CN and JP text, appears on hover
- Workload names also get glossary entries (PowerBI → Power BI 报表平台 / Power BI レポートプラットフォーム)

**Risk Section Implementation:**
- Render only if risk metrics contain non-zero/interesting findings
- Each risk item is a table: columns = metric | value | severity (Low/Medium/High)
- Severity coloring: Low = yellow, Medium = orange, High = red
- Include a summary count at the top: "N risk indicators detected"
- List specific users/IPs/operations with context
- All risk labels use `data-i18n` for language toggle support
- Severity labels: `data-i18n="low"` / `data-i18n="medium"` / `data-i18n="high"`

### Step 5: Verify Output

- Check the HTML file renders correctly
- Verify all statistics match the source data
- Ensure the file is truly self-contained (no external URLs)
- Confirm tables render correctly (pagination or truncation for large datasets)
- Test language toggle (中文 ↔ 日本語) — all UI text and glossary entries switch properly
- Verify tooltips appear on hover for operation names
- Confirm glossary is collapsible and shows correct CN/JP explanations
- Staging file can be kept for reuse or cleaned up after report generation
