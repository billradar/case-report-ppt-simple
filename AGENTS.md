# 病例汇报 PPT 改造（case-report-ppt）

安全改造临床病例汇报、出科汇报和教学病例 PowerPoint 的知识库。核心原则：**模板保真、原位替换、布局字体不变**。

本目录是一套跨 agent 的 skill：`SKILL.md` 为 Agent Skills 入口（Claude Code / opencode），本 `AGENTS.md` 为通用入口（Copilot / Cursor / Windsurf / Cline / Codex 等）。

## 核心知识位置

核心知识在 `references/` 中，按任务按需读取，不要一次全读：

| 情形 | 读取文件 |
| --- | --- |
| 安全工作流（保护原件、先复制后编辑、恢复） | `references/workflow.md` |
| 病例资料、诊断、检查、隐私、医学一致性 | `references/medical-content.md` |
| 布局、字体、缩进、动画、呈现方式 | `references/layout-design.md` |
| PowerPoint MCP / PowerShell COM 技术坑 | `references/powerpoint-technical.md` |
| 完成前审查、渲染检查、交付 | `references/validation.md` |
| 页面改造取舍与文字示例 | `references/examples.md` |
| 现成高质量病例、选病例 / 扩病例库 | `references/case-library.md` 与 `cases/` |

## 核心原则（摘要）

- **模板保真**：只在原位置改内容，不改布局、字体或模板；不改页面尺寸/主题/母版/页眉页脚/页码体系，除非用户明确授权
- **先复制、后编辑**：OneDrive/PowerPoint 的 AutoSave 会实时写回当前文档，任何编辑前必须先另存/复制副本
- **医学内容据资料整理**：优先采用操作者自行提供的详细病例资料，以其为唯一事实来源；不编造病史/检查/治疗，不确定项标注"待核实"
- **去标识化**：移除姓名、住院号、证件、联系方式等（含截图、页眉页脚、备注）
- **异常标红**：诊断关键阳性信息和异常指标用 `#EE0000`，保留单位、参考范围、时间
- **标题不重复**：同一章节标题全篇只出现一次
- **固定结构**：原页面顺序、列表和表格行列数、对象位置尺寸、字体字号及动画保持不变；仅原位替换文字与局部标红

## 模板

- 通用母版模板：`templates/病例汇报_通用模板.pptx`（二进制，被 .gitignore 忽略，需本地保留）
- 模板选择：用户指定已有 PPT 时用它改；未指定/要求按模板生成时，复制 `templates/` 内母版到目标目录后使用，母版永不修改

## 工具依赖（重要）

- 首选 **ppt-mcp**（PowerPoint MCP）：https://github.com/ykuwai/ppt-mcp ，以 `uvx ppt-mcp` 启动（需 uv 与本机 Microsoft PowerPoint）
- 准备阶段先探测：调用 `ppt_get_app_info` / `ppt_list_presentations`；**已装则跳过，直接继续**
- **未装则安装**：确认/安装 `uv`；OpenCode 写入 `opencode.json` 的 `mcp.powerpoint = {"type":"local","command":["uvx","ppt-mcp"],"enabled":true}`（其他客户端用 `{"mcpServers":{"powerpoint":{"command":"uvx","args":["ppt-mcp"]}}}`），然后**提示用户重启 OpenCode/会话后继续**
- 仍不可用时退回 PowerShell COM / python-pptx；无自动化能力时只整理内容、建立病例数据源、给出可执行修改清单，**不伪称已修改 PPT**

## 完成标准

交付时说明：输出路径、修改范围、待临床核实项、是否通过质检；保留原件，交付后删除工作副本。

## 特殊版执行方式

逐页读取现有占位内容，依照用户提供的病例事实做一对一替换。不要推断缺失患者事实、主动重构叙事或制作新图形；内容放不下时复制现有同类页面扩页。标红仍依 `references/medical-content.md`：关键阳性事实、异常数值与最终诊断关键短语用 `#EE0000`，仅改对应字符颜色。替换后比较布局、字体与原件；无合适同类页面可复制且无法保留原格式的内容不写入，列入交付说明。
