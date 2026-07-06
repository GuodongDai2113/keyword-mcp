# keyword-mcp 配置方式

在项目根目录的 `.opencode.json` 中配置 MCP 服务器。

## 最小配置

```json
{
  "mcpServers": {
    "keyword-mcp": {
      "command": "node",
      "args": ["dist/index.js"]
    }
  }
}
```

## 完整配置（含环境变量）

```json
{
  "mcpServers": {
    "keyword-mcp": {
      "command": "node",
      "args": ["dist/index.js"],
      "env": {
        "KEYWORD_MCP_PROJECTS_ROOT": ".",
        "KEYWORD_MCP_DEFAULT_MARKET": "global",
        "KEYWORD_MCP_DEFAULT_LANGUAGE": "zh"
      }
    }
  }
}
```

## 环境变量

| 变量 | 默认值 | 说明 |
|---|---|---|
| `KEYWORD_MCP_PROJECTS_ROOT` | `"."` (当前目录) | 工作簿存放根目录 |
| `KEYWORD_MCP_DEFAULT_MARKET` | `"global"` | 新项目默认市场 |
| `KEYWORD_MCP_DEFAULT_LANGUAGE` | `"en"` | 新项目默认语言 |

## 使用前准备

```powershell
npm install
npm run build
```

构建后 `dist/index.js` 即为 MCP 入口，`.opencode.json` 中的 `command` 和 `args` 指向它即可。
