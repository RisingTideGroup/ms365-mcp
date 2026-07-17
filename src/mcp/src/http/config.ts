/**
 * Configuration for Lokka HTTP mode (delegated OAuth via Entra ID).
 * All values come from environment variables; fail fast on missing required config.
 */

export interface HttpConfig {
  tenantId: string;
  clientId: string;          // Entra app registration (the protected API + OBO confidential client)
  clientSecret?: string;     // OBO requires a confidential client: secret or certificate
  certificatePath?: string;  // PEM (private key + cert) as alternative to secret
  scopeName: string;         // custom API scope name, e.g. "access_as_user"
  appIdUri: string;          // Application ID URI; MUST equal the public MCP URL for Entra resource-indicator alignment
  baseUrl: string;           // public HTTPS base URL of this server, e.g. https://lokka.risingtidegroup.net
  port: number;
  allowedOrigins: string[];  // optional browser-origin allowlist (empty = only non-browser clients)
  allowedUserOids: string[]; // optional per-user allowlist (empty = any user who can consent/sign in)
  allowedGroupId?: string;   // optional: require membership of this group id (checked via token 'groups' claim)
  useGraphBeta: boolean;
}

function required(name: string): string {
  const v = process.env[name];
  if (!v) {
    throw new Error(`Missing required environment variable: ${name}`);
  }
  return v;
}

function csv(name: string): string[] {
  return (process.env[name] || "")
    .split(",")
    .map((s) => s.trim())
    .filter(Boolean);
}

export function loadHttpConfig(): HttpConfig {
  const tenantId = required("TENANT_ID");
  const clientId = required("CLIENT_ID");
  const clientSecret = process.env.CLIENT_SECRET;
  const certificatePath = process.env.CERTIFICATE_PATH;

  if (!clientSecret && !certificatePath) {
    throw new Error(
      "HTTP mode requires a confidential client for the On-Behalf-Of flow. Set CLIENT_SECRET or CERTIFICATE_PATH."
    );
  }

  const baseUrl = (process.env.BASE_URL || `http://localhost:${process.env.PORT || 3000}`).replace(/\/$/, "");

  return {
    tenantId,
    clientId,
    clientSecret,
    certificatePath,
    scopeName: process.env.SCOPE_NAME || "access_as_user",
    // Default the App ID URI to <BASE_URL>/mcp so it matches the RFC 8707 resource
    // indicator MCP clients send. Override with APP_ID_URI if yours differs.
    appIdUri: (process.env.APP_ID_URI || `${baseUrl}/mcp`).replace(/\/$/, ""),
    baseUrl,
    port: parseInt(process.env.PORT || "3000", 10),
    allowedOrigins: csv("ALLOWED_ORIGINS"),
    allowedUserOids: csv("ALLOWED_USER_OIDS"),
    allowedGroupId: process.env.ALLOWED_GROUP_ID || undefined,
    useGraphBeta: process.env.USE_GRAPH_BETA !== "false",
  };
}
