import "dotenv/config";

import { MCPClient } from "@mcp-use/client";
import { MCPServer } from "mcp-use";
import { oauthWorkOSProvider } from "mcp-use/oauth/workos";

const expectedUsername =
  process.env.MS365_MCP_EXPECTED_USERNAME ??
  "binh.nguyen@rba-asia.com";

const server = new MCPServer({
  name: "ms365-mcp-server",
  title: "Opssoul Microsoft 365 MCP Server",
  version: "1.0.0",
  description: "WorkOS-protected Microsoft 365 MCP server for Opssoul",
  oauth: oauthWorkOSProvider(),
});

function getAllowedEmails(): string[] {
  return (
    process.env.MS365_ALLOWED_EMAILS ??
    process.env.MS365_MCP_EXPECTED_USERNAME ??
    expectedUsername
  )
    .split(",")
    .map((email) => email.trim().toLowerCase())
    .filter(Boolean);
}

function getAuthenticatedEmail(ctx: any): string | undefined {
  const auth = ctx.auth as any;

  return (
    (auth?.extra?.user?.email as string | undefined) ??
    (auth?.extra?.payload?.email as string | undefined) ??
    (auth?.extra?.payload?.preferred_username as string | undefined) ??
    (auth?.user?.email as string | undefined) ??
    (auth?.payload?.email as string | undefined) ??
    (auth?.payload?.preferred_username as string | undefined)
  )?.toLowerCase();
}

async function requireAllowedEmail(ctx: any, next: any) {
  const email = getAuthenticatedEmail(ctx);
  const allowedEmails = getAllowedEmails();

  if (!email || !allowedEmails.includes(email)) {
    console.warn(
      `AUTH DENIED: ${email ?? "unknown"} is not allowed to use MS365`,
    );

    throw new Error(
      "Forbidden: this WorkOS user is not allowed to use MS365",
    );
  }

  return next();
}

server.use("mcp:tools/call", requireAllowedEmail);
server.use("mcp:resources/read", requireAllowedEmail);
server.use("mcp:prompts/get", requireAllowedEmail);

function getMs365Args(): string[] {
  const args = [
    "-y",
    "@softeria/ms-365-mcp-server@0.149.1",
    "--org-mode",
    "--expected-username",
    expectedUsername,
  ];

  if (process.env.MS365_READ_ONLY !== "0") {
    args.push("--read-only");
  }

  if (process.env.MS365_DISCOVERY !== "0") {
    args.push("--discovery");
  }

  if (process.env.MS365_TOON !== "0") {
    args.push("--toon");
  }

  return args;
}

async function connectMs365() {
  const client = new MCPClient({
    mcpServers: {
      ms365: {
        command: "npx",
        args: getMs365Args(),
        env: {
          MS365_MCP_ORG_MODE: "true",
          MS365_MCP_EXPECTED_USERNAME: expectedUsername,
          MS365_MCP_USE_KEYTAR: process.env.MS365_MCP_USE_KEYTAR ?? "0",
          ...(process.env.MS365_MCP_TENANT_ID
            ? { MS365_MCP_TENANT_ID: process.env.MS365_MCP_TENANT_ID }
            : {}),
          ...(process.env.MS365_MCP_CLIENT_ID
            ? { MS365_MCP_CLIENT_ID: process.env.MS365_MCP_CLIENT_ID }
            : {}),
          ...(process.env.MS365_MCP_CLIENT_SECRET
            ? { MS365_MCP_CLIENT_SECRET: process.env.MS365_MCP_CLIENT_SECRET }
            : {}),
          ...(process.env.MS365_MCP_TOKEN_CACHE_PATH
            ? {
                MS365_MCP_TOKEN_CACHE_PATH:
                  process.env.MS365_MCP_TOKEN_CACHE_PATH,
              }
            : {}),
          ...(process.env.MS365_MCP_SELECTED_ACCOUNT_PATH
            ? {
                MS365_MCP_SELECTED_ACCOUNT_PATH:
                  process.env.MS365_MCP_SELECTED_ACCOUNT_PATH,
              }
            : {}),
          ...(process.env.MS365_MCP_AUTH_CACHE_COMMAND
            ? {
                MS365_MCP_AUTH_CACHE_COMMAND:
                  process.env.MS365_MCP_AUTH_CACHE_COMMAND,
              }
            : {}),
          ...(process.env.MS365_MCP_AUTH_CACHE_COMMAND_TIMEOUT_MS
            ? {
                MS365_MCP_AUTH_CACHE_COMMAND_TIMEOUT_MS:
                  process.env.MS365_MCP_AUTH_CACHE_COMMAND_TIMEOUT_MS,
              }
            : {}),
          ...(process.env.MS365_MCP_KEYVAULT_URL
            ? { MS365_MCP_KEYVAULT_URL: process.env.MS365_MCP_KEYVAULT_URL }
            : {}),
        },
      },
    },
  });

  const connection = await client.connect("ms365");

  return Object.create(connection, {
    info: {
      value: {
        ...connection.info,
        server: {
          ...connection.info.server,
          name: "ms365",
        },
      },
      enumerable: true,
    },
  });
}

async function withTimeout<T>(
  promise: Promise<T>,
  timeoutMs: number,
  message: string,
): Promise<T> {
  let timeout: ReturnType<typeof setTimeout> | undefined;

  const timeoutPromise = new Promise<never>((_, reject) => {
    timeout = setTimeout(() => reject(new Error(message)), timeoutMs);
  });

  try {
    return await Promise.race([promise, timeoutPromise]);
  } finally {
    if (timeout) {
      clearTimeout(timeout);
    }
  }
}

if (process.env.MCP_BUILD !== "1" && process.env.MS365_ENABLED !== "0") {
  try {
    const connection = await withTimeout(
      connectMs365(),
      Number(process.env.MS365_STARTUP_TIMEOUT_MS ?? 20000),
      "MS365 upstream startup timed out",
    );

    await server.proxy(connection);
  } catch (error) {
    console.error(
      "MS365 upstream disabled: failed to connect during server startup",
      error,
    );
  }
}

export default server;
