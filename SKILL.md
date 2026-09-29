---
name: case-report-ppt
description: 安全改造临床病例汇报、出科汇报和教学病例 PowerPoint（PPT/PPTX）。当需要换科室或病种、在既有页面和文本框内择要撰写病例内容、严格保留布局字体、用诊断证据链标红异常指标、设置克制动画/切换、或对医学逻辑和最终幻灯片进行质检时使用。
compatibility: opencode
metadata:
  format: agent-skills
  locale: zh-CN
---

# 病例汇报 PPT 改造

遵循“模板保真、择要撰写、布局字体不变”。优先保护原文件和患者隐私；真实病例仅据已提供资料整理；虚构教学病例可生成相关且合理的检查结果或替代临床判断。

## 约定常量（全篇唯一来源）

- 通用模板：`templates/病例汇报_通用模板.pptx`（本 skill 目录内的母版模板，相对本 SKILL.md 路径）。使用时用文件系统复制到目标 PPT 目录，不直接编辑母版。

## 开始

1. **准备：检查是否已安装 ppt-mcp（PowerPoint MCP）；未装则安装并让用户重启后继续**。
   - 项目与依赖：https://github.com/ykuwai/ppt-mcp ，以 `uvx ppt-mcp` 启动；需要 [uv](https://docs.astral.sh/uv/getting-started/installation/) 与本机 Microsoft PowerPoint（Windows/macOS）。
   - 探测：调用一次 `ppt_get_app_info` 或 `ppt_list_presentations`。**有返回 → 已安装，跳过安装，直接继续**。
   - **未安装 → 安装**：
     1. 确认 `uv` 可用（`uv --version`），缺失则先安装 uv；
     2. 写入 MCP 客户端配置：OpenCode（`opencode.json`）用 `{"$schema":"https://opencode.ai/config.json","mcp":{"powerpoint":{"type":"local","command":["uvx","ppt-mcp"],"enabled":true}}}`；其他客户端用 `{"mcpServers":{"powerpoint":{"command":"uvx","args":["ppt-mcp"]}}}`；
     3. **提示用户重启 OpenCode / 会话**（MCP 仅在启动时加载），重启后继续。
   - 重启后仍不可用：退回 PowerShell COM / python-pptx；确无自动化能力时，只整理内容与可执行修改清单，不伪称已改 PPT。
2. 先用 OpenCode 的文件读取能力加载 [workflow.md](references/workflow.md)，并在任何编辑前创建、验证安全副本。
3. 盘点源 PPT 的页面、母版、背景、版式、可编辑元素、字体和已有动画；建立唯一的病例数据源（优先采用操作者自行提供的详细病例资料；未提供时才据科室/病种生成去标识化教学病例）。
4. 按任务读取需要的参考文件：

| 情形 | 必读文件 |
| --- | --- |
| 病例资料、诊断、检查、隐私或医学一致性 | [medical-content.md](references/medical-content.md) |
| 重排版、字体、缩进、动画、切换或可视化表达 | [layout-design.md](references/layout-design.md) |
| PowerPoint MCP、PowerShell COM、格式保留或兼容性问题 | [powerpoint-technical.md](references/powerpoint-technical.md) |
| 完成前审查、渲染检查、恢复与交付 | [validation.md](references/validation.md) |
| 页面改造取舍或文字示例 | [examples.md](references/examples.md) |
| 选现成高质量病例 / 换科室病种 | [case-library.md](references/case-library.md) 与 `cases/` |

**优先用病例库**：先在 `cases/` 中按科室/病种选用现成高质量病例（见 [case-library.md](references/case-library.md)），不足时再生成。

从项目根目录使用本技能时，参考文件位于 `.opencode/skills/case-report-ppt/references/`。不要假设 OpenCode 已安装 PowerPoint MCP；先检查当前可用工具。无 PowerPoint 自动化能力时，只整理内容、提出可执行修改清单或请求用户提供可编辑环境，不伪称已修改 PPT。

## 不可突破的边界

- 不改页面尺寸/比例、主题、主题色、核心品牌字体、主背景、Logo、母版核心元素、页眉页脚或页码体系，除非用户明确授权。
- 不删除/重排原页面、文本框、表格行列或图形；扩页只复制现有同类页面；不移动、缩放对象，不修改字体、字号、字重、段落格式、动画或切换。仅允许对诊断证据的准确字符范围改色为 `#EE0000`。
- 不在原件上试错；只编辑已验证为当前激活文档的副本。
- 真实病例不将虚拟结果写成患者事实；缺失或矛盾项标“待核实”。教学病例可生成与一个诊断有关的合理检查值，明确其生成来源于交付说明，不编造指南结论或把生成值声称为真实记录。
- 删除或遮蔽可识别个人信息；不把病例材料上传至未经授权的外部服务。

## 默认决策

- 模板选择：用户指定了已有 PPT 时以该 PPT 为模板改造；未指定、或要求「根据模板生成」时，把 skill 内通用模板 `templates/病例汇报_通用模板.pptx` 复制到目标目录后作为起点（复制而非直接编辑母版）。
- `layout_change_level = 0`：按原页用途在现有对象内撰写内容。保持原页面顺序、对象位置与尺寸、字体属性、缩进、表格结构和已有动画/切换。
- 根据病例资料筛选有诊断或处置价值的体征、检查、关键阴性及病程，自己组织简明叙述；不用逐条搬运原始记录，也不用一对一替换占位词。真实病例事实可回链；教学病例生成的检查结果另行记录来源。
- 评分量表仅在已提供评分及足够依据、且对本病例确有价值时呈现；NIHSS 不是通用模板，不要求每种病都配一个评分，不根据不全的资料猜分。
- 按原有文本框容量写简短事实；放不下时复制现有同类页面，紧接对应页面插入并分配内容；不缩字、不扩框。无合适页面可复制时列入交付说明的“未写入/待补充”，不要丢失关键事实或强行塞入。
- 不自动选择可视化形式、不重组诊断链、不创建卡片/时间线、不凭空扩充真实病例事实。用户明确要求某一变化时再单独处理。
- 诊断可只有一个明确的最终诊断；不为填满模板而增加并发诊断或强制罗列鉴别诊断。
- 沿用主分支标红规则：诊断关键阳性事实、异常数值、最终诊断关键短语使用 `#EE0000`，仅给准确字符范围改色；单位、参考范围、检查时间和解释保留在原位置附近。不要把整段或整表标红。
- 每次明显改动后做页面级检查；完成后按 validation 清单逐页渲染核验。

## 交付

- 文件名固定为 `NAME-科室-病例汇报.pptx`（NAME 为汇报人、科室为轮转科室；不再附加病种、日期等），工作/恢复副本用 `NAME-科室-病例汇报_working.pptx`。
- 幻灯片（正文、页脚、备注、图片替代文本）一律不写“教学病例模拟数据”“模拟数据”“仅用于演示”等模拟/演示字样；确需说明数据性质或脱敏情况时，只写在交付说明里，不写进幻灯片。
- 保存为明确命名的最终副本；交付后删除工作副本，只保留最终版与原件，目标目录不留 `_working` 文件。汇报时说明：输出路径、修改范围、待临床核实项，以及是否通过最终质检。
