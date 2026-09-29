# PowerPoint MCP 与 PowerShell COM 注意事项

## 首选工具：ppt-mcp

- 首选 **ppt-mcp**（PowerPoint MCP）：https://github.com/ykuwai/ppt-mcp ，以 `uvx ppt-mcp` 启动，需要 [uv](https://docs.astral.sh/uv/getting-started/installation/) 与本机 Microsoft PowerPoint（Windows/macOS）；本 skill 中的 `ppt_*` 工具即来自它。
- 探测：调用 `ppt_get_app_info` / `ppt_list_presentations`。**有返回 → 已安装，跳过安装，直接继续**。
- **未安装 → 安装，然后请用户重启后继续**：① 确认 `uv` 可用（`uv --version`），缺失先装 uv；② OpenCode 写入 `opencode.json`：`{"$schema":"https://opencode.ai/config.json","mcp":{"powerpoint":{"type":"local","command":["uvx","ppt-mcp"],"enabled":true}}}`（其他客户端：`{"mcpServers":{"powerpoint":{"command":"uvx","args":["ppt-mcp"]}}}`）；③ **提示用户重启 OpenCode / 会话**后继续。
- 重启后仍不可用：退回 PowerShell COM / python-pptx；确无自动化能力时，只整理内容与可执行修改清单，不伪称已改 PPT。

## 编辑原则

- 先探测能力与对象，再写入。不同 MCP、PowerPoint 版本和受保护对象的行为不同；不要假设某个工具或属性必然可用。
- 精确定位到 slide、shape 和 text range，优先局部格式化，避免对整页/整形状做破坏性重置。
- 每一小批修改后读取回写结果或查看预览，确认页码、对象、内容和样式均正确。

## MCP 常见风险

- 类似 `ppt_set_text` 的全量写入通常会替换原文本，且可能重置字体、段落、项目符号、富文本、超链接或运行级颜色。先保存格式信息；能用局部文本替换或格式范围 API 时优先使用。
- 对异常值标红时，仅格式化对应字符范围；验证范围索引、中文字符和多行文本，避免把邻近单位、参考范围或整段错误染色。
- shape 名称可能为空、重复或包含中文；优先使用稳定的 slide 编号和 `Shape.Id`，并在每次会话重新枚举对象，不长期依赖集合索引。
- 预览/截图只能验证视觉，不能保证文字、动画、超链接或备注正确；视觉检查与结构检查都要做。
- JSON 字符串含中文引号、反斜杠、换行或公式时，先使用稳妥的转义/序列化；不要手拼易损 JSON。

## PowerShell COM 注意事项

- 仅在 Windows PowerPoint 已安装、用户授权且 MCP 无法完成任务时使用 COM；操作前仍须确认工作副本已激活。
- COM 集合常为 1 基索引；通过显式计数与索引遍历，避免假定 `foreach` 在所有集合上行为一致。
- 文本通常在 `Shape.TextFrame` / `TextFrame2`；段落格式在 `ParagraphFormat` / `ParagraphFormat2`。对象模型和属性支持会随版本变化，读取后再写入。
- 首行缩进优先设置段落级 `FirstLineIndent`（及需要时的左缩进），不要靠字符串空格。仅对需要的段落/文本范围应用。
- 动画与切换 API 可能对组合对象、占位符或特定版本不一致。先用一页测试；失败则保留静态、可读的页面而非反复修改。
- COM 进程可能残留。仅释放本次创建的 COM 引用；不要强杀用户正在使用的 PowerPoint 进程。

## 兼容性与恢复

- 另存 `.pptx` 后重新打开检查。出现“修复演示文稿”提示时，保留损坏副本和修复报告，回到最近恢复点定位问题。
- 警惕 SmartArt、嵌入对象、媒体、图表、公式、组合形状、母版占位符和动画触发器；这些对象优先最小编辑。
- 若能力不足，交付不依赖该能力的稳定版本，并如实标注未应用的动画或需要人工完成的步骤。
