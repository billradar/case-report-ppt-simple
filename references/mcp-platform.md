# 按系统选择演示文稿 MCP

## Linux：优先探测 WPS MCP

项目：[Xiao-rx/wps-mcp-server](https://github.com/Xiao-rx/wps-mcp-server)。先检查客户端是否已提供 `wps` MCP；有工具时直接调用并验证能力。未配置时，从可信仓库安装到本机独立目录，按其 README 安装 Python 3.10+、`mcp`、`python-docx`、`openpyxl`、`python-pptx`。例如客户端的 stdio 配置为：

```json
{"mcpServers":{"wps":{"command":"python","args":["/absolute/path/to/wps-mcp-server/wps_mcp.py"],"env":{"WPS_WORKSPACE":"/absolute/path/to/workdir"}}}}
```

不同客户端的 MCP 配置结构不同，OpenCode 按其当前版本改为本地 MCP 的 `command` 数组。配置后按客户端的加载机制重连。

**能力门槛**：该项目当前演示工具是 `create_presentation`、`add_slide`、`add_text_to_slide`、`set_slide_layout`，通过 `python-pptx` 在 `WPS_WORKSPACE` 读写 `.pptx`。它不操作 WPS 桌面窗口，也不提供已有文本对象的精确替换、局部字符标红、复制现有页面和动画保留工具。因此仅在其实际工具能完成某一步时调用，不用 `create_presentation` 覆盖现有模板；不能靠它单独完成本技能的保版式改造或用户可见的 WPS 操作。切换至有相应能力的工具或文件级编辑并如实说明可见性限制。若用户要求全过程在窗口可见且没有可用 GUI 自动化，就停在该门槛，不能暗中转为文件编辑。

## Windows / macOS：PowerPoint MCP

项目：[ykuwai/ppt-mcp](https://github.com/ykuwai/ppt-mcp)，按项目要求在安装了 Microsoft PowerPoint 的本机配置 `uvx ppt-mcp`。先探测实际可用的 `ppt_*` 工具，确认能读取、局部编辑、着色、复制页面、保存及回读。只有窗口与用户同桌面会话时才能承诺可见过程。Windows 无 MCP 时可考虑 PowerShell COM；macOS 按实际可用自动化能力判断。

## 共通规则

MCP 的存在不证明功能足够，也不证明用户能看见窗口。始终先复制源 PPT，枚举实际工具、验证作用目标和可见性；先做单页试改与回读，再分批处理。使用文件级替代工具须明确说明动画、页码和视觉验证范围，不把后台文件改动描述为桌面窗口操作。
