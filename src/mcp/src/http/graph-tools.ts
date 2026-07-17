/**
 * Per-session MCP server factory for HTTP mode.
 *
 * The Lokka-Microsoft tool logic is ported from main.ts (stdio mode) with one
 * structural change: instead of module-global authManager/graphClient, every
 * session gets its own Graph client and Azure credential backed by the
 * session's OBO broker + the connecting user's assertion. Sessions are fully
 * isolated — no token can leak between users.
 *
 * main.ts is intentionally left untouched so stdio mode and upstream merges
 * keep working.
 */

import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { Client, PageIterator, PageCollection } from "@microsoft/microsoft-graph-client";
import { logger } from "../logger.js";
import { OboBroker, GRAPH_SCOPE, AZURE_RM_SCOPE } from "./obo.js";
import type { VerifiedUser } from "./token-verify.js";

export interface SessionContext {
  user: VerifiedUser;
  broker: OboBroker;
  useGraphBeta: boolean;
  /** Updated on every authenticated request so long-lived sessions keep a fresh assertion. */
  refreshAssertion(token: string, expiresAt: number): void;
  getAssertion(): string;
}

export function createSessionContext(user: VerifiedUser, broker: OboBroker, useGraphBeta: boolean): SessionContext {
  let assertion = user.rawToken;
  let assertionExp = user.expiresAt;
  return {
    user,
    broker,
    useGraphBeta,
    refreshAssertion(token: string, expiresAt: number) {
      assertion = token;
      assertionExp = expiresAt;
    },
    getAssertion() {
      if (assertionExp && assertionExp * 1000 < Date.now()) {
        logger.info(`Assertion for user ${user.oid} is expired; relying on client to re-authenticate`);
      }
      return assertion;
    },
  };
}

// ---------------------------------------------------------------------------
// SSRF-safe Azure URL helpers (ported verbatim in behavior from main.ts)
// ---------------------------------------------------------------------------

function validateAzurePath(path: string): void {
  if (!path) {
    throw new Error("Path cannot be empty");
  }
  const forbiddenPatterns: Array<{ pattern: RegExp; reason: string }> = [
    { pattern: /@/, reason: "contains @ (host-escape character)" },
    { pattern: /\/{2,}/, reason: "contains double slashes (protocol-relative URL)" },
    { pattern: /^https?:\/\//i, reason: "is an absolute URL" },
    { pattern: /\\/g, reason: "contains backslashes" },
  ];
  for (const { pattern, reason } of forbiddenPatterns) {
    if (pattern.test(path)) {
      logger.error(`Invalid Azure path rejected: ${reason}. Path: ${path}`);
      throw new Error(`Invalid path: ${reason}. Path must be a relative path starting with '/'.`);
    }
  }
  if (!path.startsWith("/")) {
    throw new Error("Invalid path: must start with '/'. Paths must be relative, not absolute URLs.");
  }
  try {
    const testUrl = new URL(path, "https://management.azure.com");
    if (testUrl.hostname !== "management.azure.com") {
      throw new Error(`Invalid path: would resolve to different host (${testUrl.hostname}). Path must target management.azure.com.`);
    }
  } catch (e: any) {
    if (e.message.includes("would resolve to different host")) {
      throw e;
    }
    throw new Error("Invalid path: cannot be parsed as a valid URL component.");
  }
}

function buildAzureUrl(
  subscriptionId: string | undefined,
  path: string,
  apiVersion: string,
  queryParams?: Record<string, string>
): string {
  const urlObj = new URL("https://management.azure.com");
  let pathname = "";
  if (subscriptionId) {
    pathname += `/subscriptions/${subscriptionId}`;
  }
  pathname += path;
  urlObj.pathname = pathname;
  urlObj.searchParams.set("api-version", apiVersion);
  if (queryParams) {
    for (const [key, value] of Object.entries(queryParams)) {
      urlObj.searchParams.append(key, String(value));
    }
  }
  return urlObj.toString();
}

// ---------------------------------------------------------------------------
// Server factory
// ---------------------------------------------------------------------------

export function createLokkaServer(ctx: SessionContext): McpServer {
  const server = new McpServer({
    name: "Lokka-Microsoft",
    version: "0.3.0-http",
  });

  const defaultGraphApiVersion: "v1.0" | "beta" = ctx.useGraphBeta ? "beta" : "v1.0";

  // Per-session Graph client: every request acquires (cached) a delegated
  // Graph token via OBO from this user's assertion.
  const graphClient = Client.initWithMiddleware({
    authProvider: {
      getAccessToken: async () => ctx.broker.getDownstreamToken(ctx.getAssertion(), GRAPH_SCOPE),
    },
  });

  interface RequestParams {
    apiType: "graph" | "azure";
    path: string;
    method: "get" | "post" | "put" | "patch" | "delete";
    apiVersion?: string;
    subscriptionId?: string;
    queryParams?: Record<string, string>;
    body?: any;
    graphApiVersion: "v1.0" | "beta";
    fetchAll: boolean;
    consistencyLevel?: string;
  }

  async function execute(p: RequestParams) {
    const { apiType, path, method, apiVersion, subscriptionId, queryParams, body, graphApiVersion, fetchAll, consistencyLevel } = p;
    const effectiveGraphApiVersion = !ctx.useGraphBeta ? "v1.0" : graphApiVersion;
    logger.info(`[${ctx.user.preferredUsername || ctx.user.oid}] ${method.toUpperCase()} ${apiType}:${path} (graphApiVersion=${effectiveGraphApiVersion}, fetchAll=${fetchAll})`);
    let determinedUrl: string | undefined;

    try {
      let responseData: any;

      if (apiType === "graph") {
        determinedUrl = `https://graph.microsoft.com/${effectiveGraphApiVersion}`;
        let request = graphClient.api(path).version(effectiveGraphApiVersion);
        if (queryParams && Object.keys(queryParams).length > 0) {
          request = request.query(queryParams);
        }
        if (consistencyLevel) {
          request = request.header("ConsistencyLevel", consistencyLevel);
        }
        switch (method.toLowerCase()) {
          case "get":
            if (fetchAll) {
              const firstPageResponse: PageCollection = await request.get();
              const odataContext = firstPageResponse["@odata.context"];
              let allItems: any[] = firstPageResponse.value || [];
              const callback = (item: any) => {
                allItems.push(item);
                return true;
              };
              const pageIterator = new PageIterator(graphClient, firstPageResponse, callback);
              await pageIterator.iterate();
              responseData = { "@odata.context": odataContext, value: allItems };
            } else {
              responseData = await request.get();
            }
            break;
          case "post":
            responseData = await request.post(body ?? {});
            break;
          case "put":
            responseData = await request.put(body ?? {});
            break;
          case "patch":
            responseData = await request.patch(body ?? {});
            break;
          case "delete":
            responseData = await request.delete();
            if (responseData === undefined || responseData === null) {
              responseData = { status: "Success (No Content)" };
            }
            break;
          default:
            throw new Error(`Unsupported method: ${method}`);
        }
      } else {
        // Azure Resource Management via delegated OBO token
        determinedUrl = "https://management.azure.com";
        if (!apiVersion) {
          throw new Error("API version is required for Azure Resource Management queries");
        }
        validateAzurePath(path);

        const azureToken = await ctx.broker.getDownstreamToken(ctx.getAssertion(), AZURE_RM_SCOPE);
        const url = buildAzureUrl(subscriptionId, path, apiVersion, queryParams);
        const headers: Record<string, string> = {
          Authorization: `Bearer ${azureToken}`,
          "Content-Type": "application/json",
        };
        const requestOptions: RequestInit = { method: method.toUpperCase(), headers };
        if (["POST", "PUT", "PATCH"].includes(method.toUpperCase())) {
          requestOptions.body = body ? JSON.stringify(body) : JSON.stringify({});
        }

        if (fetchAll && method === "get") {
          let allValues: any[] = [];
          let currentUrl: string | null = url;
          while (currentUrl) {
            const pageToken = await ctx.broker.getDownstreamToken(ctx.getAssertion(), AZURE_RM_SCOPE);
            const pageHeaders = { ...headers, Authorization: `Bearer ${pageToken}` };
            const pageResponse = await fetch(currentUrl, { method: "GET", headers: pageHeaders });
            const pageText = await pageResponse.text();
            let pageData: any;
            try {
              pageData = pageText ? JSON.parse(pageText) : {};
            } catch {
              pageData = { rawResponse: pageText };
            }
            if (!pageResponse.ok) {
              throw new Error(`API error (${pageResponse.status}) during Azure RM pagination on ${currentUrl}: ${JSON.stringify(pageData)}`);
            }
            if (pageData.value && Array.isArray(pageData.value)) {
              allValues = allValues.concat(pageData.value);
            } else if (currentUrl === url && !pageData.nextLink) {
              allValues.push(pageData);
            }
            currentUrl = pageData.nextLink || null;
          }
          responseData = { allValues };
        } else {
          const apiResponse = await fetch(url, requestOptions);
          const responseText = await apiResponse.text();
          try {
            responseData = responseText ? JSON.parse(responseText) : {};
          } catch {
            responseData = { rawResponse: responseText };
          }
          if (!apiResponse.ok) {
            throw new Error(`API error (${apiResponse.status}) for Azure RM: ${JSON.stringify(responseData)}`);
          }
        }
      }

      let resultText = `Result for ${apiType} API (${apiType === "graph" ? effectiveGraphApiVersion : apiVersion}) - ${method} ${path}:\n\n`;
      resultText += JSON.stringify(responseData, null, 2);
      if (!fetchAll && method === "get") {
        const nextLinkKey = apiType === "graph" ? "@odata.nextLink" : "nextLink";
        if (responseData && responseData[nextLinkKey]) {
          resultText += `\n\nNote: More results are available. To retrieve all pages, add the parameter 'fetchAll: true' to your request.`;
        }
      }
      return { content: [{ type: "text" as const, text: resultText }] };
    } catch (error: any) {
      logger.error(`Error in Lokka tool (user: ${ctx.user.oid}, apiType: ${apiType}, path: ${path}, method: ${method}):`, error);
      if (!determinedUrl) {
        determinedUrl = apiType === "graph" ? `https://graph.microsoft.com/${effectiveGraphApiVersion}` : "https://management.azure.com";
      }
      const errorBody = error.body ? (typeof error.body === "string" ? error.body : JSON.stringify(error.body)) : "N/A";
      return {
        content: [{
          type: "text" as const,
          text: JSON.stringify({
            error: error instanceof Error ? error.message : String(error),
            statusCode: error.statusCode || "N/A",
            errorBody,
            attemptedBaseUrl: determinedUrl,
          }),
        }],
        isError: true,
      };
    }
  }

  // Shared schema fragments
  const commonShape = {
    apiType: z.enum(["graph", "azure"]).describe("Type of Microsoft API to query. Options: 'graph' for Microsoft Graph (Entra) or 'azure' for Azure Resource Management."),
    path: z.string().describe("The Azure or Graph API URL path to call (e.g. '/users', '/groups', '/subscriptions')"),
    apiVersion: z.string().optional().describe("Azure Resource Management API version (required for apiType Azure)"),
    subscriptionId: z.string().optional().describe("Azure Subscription ID (for Azure Resource Management)."),
    queryParams: z.record(z.string()).optional().describe("Query parameters for the request"),
    graphApiVersion: z.enum(["v1.0", "beta"]).optional().default(defaultGraphApiVersion).describe(`Microsoft Graph API version to use (default: ${defaultGraphApiVersion})`),
  };

  // ---- READ tool: structurally incapable of writes (method is hardcoded to GET) ----
  server.registerTool(
    "microsoft-read",
    {
      title: "Microsoft Graph/Azure — Read",
      description:
        "Read-only queries (HTTP GET) against Microsoft Graph (Entra, Intune, Teams, SharePoint) and Azure Resource Management, acting as the signed-in user (delegated permissions). This tool cannot modify anything. For Graph GET requests using advanced query parameters ($filter, $count, $search, $orderby), set 'consistencyLevel: \"eventual\"'.",
      inputSchema: {
        ...commonShape,
        fetchAll: z.boolean().optional().default(false).describe("Set to true to automatically fetch all pages for list results (e.g., users, groups). Default is false."),
        consistencyLevel: z.string().optional().describe("Graph API ConsistencyLevel header. ADVISED to be set to 'eventual' for Graph GET requests using advanced query parameters ($filter, $count, $search, $orderby)."),
      },
      annotations: {
        readOnlyHint: true,
        openWorldHint: true,
      },
    },
    async (p: any) => execute({ ...p, method: "get", body: undefined })
  );

  // ---- WRITE tool: all mutating verbs live here, and only here ----
  server.registerTool(
    "microsoft-write",
    {
      title: "Microsoft Graph/Azure — Write (MUTATES TENANT)",
      description:
        "Mutating operations (POST, PUT, PATCH, DELETE) against Microsoft Graph and Azure Resource Management, acting as the signed-in user. This CHANGES tenant configuration or data. Before calling, state clearly to the user what will change, on which object, and with what payload. Never batch unrelated changes into one call.",
      inputSchema: {
        ...commonShape,
        method: z.enum(["post", "put", "patch", "delete"]).describe("Mutating HTTP method to use"),
        body: z.record(z.string(), z.any()).optional().describe("The request body (for POST, PUT, PATCH)"),
      },
      annotations: {
        readOnlyHint: false,
        destructiveHint: true,
        openWorldHint: true,
      },
    },
    async (p: any) => execute({ ...p, fetchAll: false, consistencyLevel: undefined })
  );

  server.registerTool(
    "get-auth-status",
    {
      title: "Auth status",
      description: "Shows who is signed in to this Lokka session and how authentication works (delegated OAuth with On-Behalf-Of).",
      inputSchema: {},
      annotations: { readOnlyHint: true },
    },
    async () => ({
      content: [{
        type: "text" as const,
        text: JSON.stringify({
          mode: "delegated-http",
          signedInUser: ctx.user.preferredUsername || ctx.user.name || ctx.user.oid,
          userObjectId: ctx.user.oid,
          tenantId: ctx.user.tid,
          apiScopes: ctx.user.scopes,
          tools: {
            "microsoft-read": "GET only — safe to auto-approve in your MCP client",
            "microsoft-write": "POST/PUT/PATCH/DELETE — keep on per-call approval in your MCP client",
          },
          note: "All Graph/Azure calls run as this user via the On-Behalf-Of flow. Effective permissions = the intersection of the app's delegated permissions and this user's own privileges.",
        }, null, 2),
      }],
    })
  );

  return server;
}
