---
layout: default
permalink: /mechanism/
---

# 生成の仕組み

## AIエージェントOpenCode

### MCPサーバー

- Filesystem MCP Server

```sh
dnf -y update
dnf install nodejs npm -y
npm install -g opencode-ai
npm install -g @modelcontextprotocol/server-filesystem

mkdir -p ~/.config/opencode && cat <<'EOF' > ~/.config/opencode/opencode.json
{
  "mcp": {
    "filesystem": {
      "type": "local",
      "command": [
        "npx",
        "-y",
        "@modelcontextprotocol/server-filesystem",
        "/work"
      ],
      "enabled": true
    }
  }
}
EOF
```
