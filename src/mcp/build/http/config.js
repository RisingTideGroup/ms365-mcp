function required(name) {
  const v = process.env[name];
  if (!v) {
    throw new Error(`Missing required environment variable: ${name}`);
  }
  return v;
}
function csv(name) {
  return (process.env[name] || "").split(",").map((s) => s.trim()).filter(Boolean);
}
function loadHttpConfig() {
  const tenantId = required("TENANT_ID");
  const clientId = required("CLIENT_ID");
  const clientSecret = process.env.CLIENT_SECRET;
  const certificatePath = process.env.CERTIFICATE_PATH;
  if (!clientSecret && !certificatePath) {
    throw new Error(
      "HTTP mode requires a confidential client for the On-Behalf-Of flow. Set CLIENT_SECRET or CERTIFICATE_PATH."
    );
  }
  const baseUrl = (process.env.BASE_URL || `http://localhost:${process.env.PORT || 3e3}`).replace(/\/$/, "");
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
    allowedGroupId: process.env.ALLOWED_GROUP_ID || void 0,
    useGraphBeta: process.env.USE_GRAPH_BETA !== "false"
  };
}
export {
  loadHttpConfig
};
