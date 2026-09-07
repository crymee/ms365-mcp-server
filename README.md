# ms365-mcp-server

WorkOS-protected Microsoft 365 MCP server for Opssoul.

Runtime defaults:

- Allows only `binh.nguyen@rba-asia.com` through WorkOS.
- Runs `@softeria/ms-365-mcp-server@0.149.1` in delegated, read-only, org mode.
- Keeps the HTTP server alive even when the Microsoft delegated token cache is not bootstrapped yet.

Required runtime env:

```env
MS365_ALLOWED_EMAILS=binh.nguyen@rba-asia.com
MS365_MCP_EXPECTED_USERNAME=binh.nguyen@rba-asia.com
MS365_MCP_USE_KEYTAR=0
```

Optional runtime env:

```env
MS365_MCP_TENANT_ID=
MS365_MCP_CLIENT_ID=
MS365_MCP_CLIENT_SECRET=
MS365_MCP_TOKEN_CACHE_PATH=
MS365_MCP_SELECTED_ACCOUNT_PATH=
MS365_MCP_AUTH_CACHE_COMMAND=
MS365_MCP_AUTH_CACHE_COMMAND_TIMEOUT_MS=
MS365_MCP_KEYVAULT_URL=
MS365_READ_ONLY=1
MS365_DISCOVERY=1
MS365_TOON=1
```
