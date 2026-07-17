#!/usr/bin/env node
import express from "express";
import { randomUUID } from "crypto";
import { StreamableHTTPServerTransport } from "@modelcontextprotocol/sdk/server/streamableHttp.js";
import { isInitializeRequest } from "@modelcontextprotocol/sdk/types.js";
import fetch from "isomorphic-fetch";
import { logger } from "../logger.js";
import { loadHttpConfig } from "./config.js";
import { TokenVerifier, AuthError } from "./token-verify.js";
import { OboBroker } from "./obo.js";
import { createLokkaServer, createSessionContext } from "./graph-tools.js";
global.fetch = global.fetch || fetch;
const config = loadHttpConfig();
const verifier = new TokenVerifier(config);
const broker = new OboBroker(config);
const sessions = /* @__PURE__ */ new Map();
const app = express();
app.set("trust proxy", true);
app.use(express.json({ limit: "4mb" }));
app.use((req, res, next) => {
  const origin = req.headers.origin;
  if (origin && config.allowedOrigins.length > 0 && !config.allowedOrigins.includes(origin)) {
    res.status(403).json({ error: "origin_not_allowed" });
    return;
  }
  if (origin && (config.allowedOrigins.length === 0 || config.allowedOrigins.includes(origin))) {
    res.setHeader("Access-Control-Allow-Origin", origin);
    res.setHeader("Access-Control-Allow-Headers", "Authorization, Content-Type, Mcp-Session-Id, Mcp-Protocol-Version");
    res.setHeader("Access-Control-Allow-Methods", "GET, POST, DELETE, OPTIONS");
    res.setHeader("Access-Control-Expose-Headers", "Mcp-Session-Id, WWW-Authenticate");
  }
  if (req.method === "OPTIONS") {
    res.status(204).end();
    return;
  }
  next();
});
const prm = {
  resource: `${config.baseUrl}/mcp`,
  authorization_servers: [`https://login.microsoftonline.com/${config.tenantId}/v2.0`],
  scopes_supported: [`api://${config.clientId}/${config.scopeName}`],
  bearer_methods_supported: ["header"],
  resource_name: "Lokka (delegated Microsoft Graph MCP)"
};
app.get("/.well-known/oauth-protected-resource", (_req, res) => {
  res.json(prm);
});
app.get("/.well-known/oauth-protected-resource/mcp", (_req, res) => {
  res.json(prm);
});
app.get("/healthz", (_req, res) => {
  res.json({ ok: true });
});
function unauthorized(res, e) {
  res.status(e.status).set(
    "WWW-Authenticate",
    `Bearer error="${e.code}", error_description="${e.message.replace(/"/g, "'")}", resource_metadata="${config.baseUrl}/.well-known/oauth-protected-resource"`
  ).json({
    jsonrpc: "2.0",
    error: { code: -32001, message: e.message },
    id: null
  });
}
async function authenticate(req, res) {
  try {
    return await verifier.verify(req.headers.authorization);
  } catch (e) {
    if (e instanceof AuthError) {
      unauthorized(res, e);
    } else {
      logger.error("Unexpected auth error", e);
      unauthorized(res, new AuthError(401, "invalid_token", "Authentication failed"));
    }
    return null;
  }
}
app.post("/mcp", async (req, res) => {
  const user = await authenticate(req, res);
  if (!user) return;
  const sessionId = req.headers["mcp-session-id"];
  try {
    if (sessionId && sessions.has(sessionId)) {
      const session = sessions.get(sessionId);
      if (session.ownerOid !== user.oid) {
        unauthorized(res, new AuthError(403, "access_denied", "Session belongs to a different user"));
        return;
      }
      session.ctx.refreshAssertion(user.rawToken, user.expiresAt);
      await session.transport.handleRequest(req, res, req.body);
      return;
    }
    if (!sessionId && isInitializeRequest(req.body)) {
      const ctx = createSessionContext(user, broker, config.useGraphBeta);
      const transport = new StreamableHTTPServerTransport({
        sessionIdGenerator: () => randomUUID(),
        onsessioninitialized: (sid) => {
          sessions.set(sid, { transport, ctx, ownerOid: user.oid, createdAt: /* @__PURE__ */ new Date() });
          logger.info(`Session ${sid} initialized for ${user.preferredUsername || user.oid}`);
        }
      });
      transport.onclose = () => {
        if (transport.sessionId) {
          sessions.delete(transport.sessionId);
          logger.info(`Session ${transport.sessionId} closed`);
        }
      };
      const server = createLokkaServer(ctx);
      await server.connect(transport);
      await transport.handleRequest(req, res, req.body);
      return;
    }
    res.status(400).json({
      jsonrpc: "2.0",
      error: { code: -32e3, message: "Bad Request: no valid session id and not an initialize request" },
      id: null
    });
  } catch (e) {
    logger.error("Error handling MCP request", e);
    if (!res.headersSent) {
      res.status(500).json({ jsonrpc: "2.0", error: { code: -32603, message: "Internal server error" }, id: null });
    }
  }
});
async function handleSessionRequest(req, res) {
  const user = await authenticate(req, res);
  if (!user) return;
  const sessionId = req.headers["mcp-session-id"];
  if (!sessionId || !sessions.has(sessionId)) {
    res.status(400).send("Invalid or missing session ID");
    return;
  }
  const session = sessions.get(sessionId);
  if (session.ownerOid !== user.oid) {
    unauthorized(res, new AuthError(403, "access_denied", "Session belongs to a different user"));
    return;
  }
  session.ctx.refreshAssertion(user.rawToken, user.expiresAt);
  await session.transport.handleRequest(req, res);
}
app.get("/mcp", handleSessionRequest);
app.delete("/mcp", handleSessionRequest);
setInterval(() => {
  const cutoff = Date.now() - 8 * 60 * 60 * 1e3;
  for (const [sid, s] of sessions) {
    if (s.createdAt.getTime() < cutoff) {
      try {
        s.transport.close();
      } catch {
      }
      sessions.delete(sid);
    }
  }
}, 15 * 60 * 1e3).unref();
app.listen(config.port, () => {
  logger.info(`Lokka HTTP (delegated OAuth) listening on :${config.port}`);
  logger.info(`Resource: ${config.baseUrl}/mcp`);
  logger.info(`Authorization server: https://login.microsoftonline.com/${config.tenantId}/v2.0`);
  logger.info(`Required scope: api://${config.clientId}/${config.scopeName}`);
});
