Log Analytics Skill 文件说明
============================

更新时间: 2026-06-11

一、当前 PowerShell 脚本
----------------------

1. scripts/main.ps1
   主入口脚本。负责解析用户参数、选择日志表和时间范围、管理本地查询缓存、逐表调用 Log Analytics 查询脚本、收集 CSV 数据源，并最终调用 HTML 报告生成脚本。

   常用方式:
   - .\scripts\main.ps1
   - .\scripts\main.ps1 -TableName "SigninLogs" -CustomStart "2026-06-10T00:00:00" -CustomEnd "2026-06-10T03:00:00"
   - .\scripts\main.ps1 -ForceRefresh
   - .\scripts\main.ps1 -ClearCache

2. scripts/query-log-analytics.ps1
   Azure Log Analytics 查询脚本。负责加载 Az.Accounts / Az.OperationalInsights 模块，登录 Azure China Cloud，按传入的表名和时间范围生成或执行 KQL，并把查询结果导出为 CSV。

   该脚本通常由 main.ps1 调用，也可以单独用于排查 Azure 登录、模块修复或自定义 KQL 查询。

3. scripts/generate-html-report.ps1
   报告生成脚本。读取 main.ps1 收集到的一个或多个 CSV，按风险规则聚合登录失败、可疑成功登录、删除/禁用操作、Service Principal 权限变更、许可证、邮箱容量、DCR 采集错误和 Intune 审计风险，并生成自包含 HTML 报告。

   当前报告 HTML/CSS 直接内嵌在脚本中，不再依赖 scripts/template 目录。

4. scripts/log-analyzer-shared.ps1
   共享函数库。保存支持的日志表清单、表名解析、时间范围处理、路径生成、缓存校验、可信 IP 读取、Microsoft Service Tags 缓存、字段标准化、KQL 生成和报告辅助函数。

   main.ps1、query-log-analytics.ps1 和 generate-html-report.ps1 都会 dot-source 这个文件。

二、其他关键文件和目录
--------------------

1. SKILL.md
   Skill 的使用说明和维护说明。用户入口、常用参数、默认处理表、风险关注点和 KQL 逻辑说明都在这里。

2. scripts/config/TrustedLocation_KJ.txt
   本地可信 IP/CIDR 配置。报告中的可疑 IP 判断会排除这里定义的可信地址。

3. scripts/config/TrustedLocation_IDC_Ali.txt
   另一组本地可信 IP/CIDR 配置。用途同上。

4. scripts/cache/
   项目内缓存目录。当前主要运行缓存默认写到用户临时目录:
   %USERPROFILE%\AppData\Local\Temp\opencode\cache

5. final_report_*.html / final_report_merged_*.html
   生成的 HTML 报告文件。单表报告使用 final_report_<TableName>_* 命名，多表合并报告使用 final_report_merged_* 命名。

6. scripts/LAW-logs.xlsx
   当前工作区中的日志相关 Excel 文件。它不是 PowerShell 主流程的运行依赖，若作为人工参考或输入样例使用，需要另行确认。

7. scripts/archive/
   归档和说明目录。用于保存历史说明、迁移记录或当前这份 readme.txt。归档文档不参与 main.ps1 的运行链路。

三、整体工作流程
----------------

1. 用户运行 scripts/main.ps1。

2. main.ps1 解析参数。
   如果用户没有传入明确时间范围，会交互式选择最近 3 小时、最近 1 天或最近 n 天。用户也可以通过 AnalysisDate、StartDate/EndDate 或 CustomStart/CustomEnd 指定时间。

3. main.ps1 确定目标日志表。
   如果没有指定 TableName，则使用 log-analyzer-shared.ps1 中定义的默认支持表；如果传入逗号分隔表名，则逐个解析。

4. main.ps1 为每张表生成 CSV、缓存和报告路径。

5. main.ps1 检查本地缓存。
   如果缓存命中且未过期，则直接复制缓存 CSV 作为本次数据源。若缓存缺失、过期、元数据不匹配或 payload 无效，则重新查询。

6. 需要查询时，main.ps1 调用 scripts/query-log-analytics.ps1。
   查询脚本负责 Azure China Cloud 登录、Az 模块加载、按表构造风险预过滤 KQL，并把结果写成 CSV。

7. main.ps1 收集所有成功生成的 CSV。
   如果没有任何可用 CSV，会停止并报错；否则继续生成报告。

8. main.ps1 调用 scripts/generate-html-report.ps1。
   报告脚本读取 CSV，标准化字段，计算各类风险指标，聚合重复事件，生成 HTML 页面。

9. main.ps1 输出报告路径和 file:// URL。
   默认会打开生成的 HTML 报告；传入 -NoOpen 时只生成文件，不自动打开。

四、维护注意事项
----------------

1. 新增或删除支持表时，优先修改 scripts/log-analyzer-shared.ps1 中的表清单、字段映射和 KQL 生成逻辑。

2. 修改风险展示或报告布局时，主要修改 scripts/generate-html-report.ps1。

3. 修改 Azure 登录、模块加载、查询执行或 CSV 导出时，主要修改 scripts/query-log-analytics.ps1。

4. 修改缓存、时间范围、入口参数或多表编排时，主要修改 scripts/main.ps1。

5. 删除文件前应先用 rg 检查引用关系，重点检查剩余 .ps1、SKILL.md 和当前说明文档。

6. 获取License数量时，调用的API是：https://microsoftgraph.chinacloudapi.cn/v1.0/subscribedSkus



不要使用DeviceCodeFlow登录，要使用AuthCodeFlow登录
需要拿Graph，需要Token，不同的人需要不同的App，一个人需要一个clientids
MS Graph PowerShell是国际版，而中国使用21v
organization权限，offline权限
