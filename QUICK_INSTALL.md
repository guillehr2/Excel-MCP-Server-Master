# 🚀 Excel MCP Server - Quick Install

## Installation in 30 seconds

### 1️⃣ Open your Claude Desktop config file:

**Windows:**
```
%APPDATA%\Claude\claude_desktop_config.json
```

**macOS:**
```
~/Library/Application Support/Claude/claude_desktop_config.json
```

### 2️⃣ Add this configuration:

```json
{
  "mcpServers": {
    "excel-master": {
      "command": "npx",
      "args": [
        "-y",
        "@guillehr2/excel-mcp-server@latest"
      ]
    }
  }
}
```

### 3️⃣ Restart Claude Desktop

That's it! You now have access to all Excel tools in Claude.

## 🎯 Test it works

Ask Claude: "Create an Excel file with sales data and a chart"

## 📚 Learn more

- Full documentation: [README.md](README.md)
- Examples: [docs/examples.md](docs/examples.md)
- NPM Package: https://www.npmjs.com/package/@guillehr2/excel-mcp-server

---

**Created by Guillem Hermida** | [GitHub](https://github.com/guillehr2)
