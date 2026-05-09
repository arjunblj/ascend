# MCP setup

Use the stdio server at `apps/mcp/src/index.ts`.

**Cursor** (`~/.cursor/mcp.json` or project `.cursor/mcp.json`):

```json
{
  "mcpServers": {
    "ascend": {
      "command": "bun",
      "args": ["run", "apps/mcp/src/index.ts"],
      "cwd": "<path-to-ascend-repo-root>"
    }
  }
}
```

**Claude Desktop**: merge the same `ascend` entry under `mcpServers` in `claude_desktop_config.json`, with an absolute `cwd`.

Restart the host app after editing config.
