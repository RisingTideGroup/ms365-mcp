/**
 * Bearer token validation for the Lokka HTTP endpoint.
 *
 * The MCP server is an OAuth 2.1 protected resource. Entra ID is the
 * authorization server. Every request must carry a bearer access token whose
 * audience is THIS app registration (not Graph!). We validate:
 *   - signature against the tenant's JWKS
 *   - issuer (v1.0 sts.windows.net and v2.0 login.microsoftonline.com both accepted)
 *   - audience (client id or api://client-id)
 *   - tenant id claim
 *   - scope contains our custom API scope
 *   - optional user-oid / group allowlists
 *
 * The raw token is retained as the OBO assertion for Graph token exchange.
 */

import { createRemoteJWKSet, jwtVerify, JWTPayload } from "jose";
import { logger } from "../logger.js";
import type { HttpConfig } from "./config.js";

export interface VerifiedUser {
  oid: string;
  tid: string;
  name?: string;
  preferredUsername?: string;
  scopes: string[];
  rawToken: string; // OBO assertion
  expiresAt: number; // epoch seconds
}

export class TokenVerifier {
  private jwks: ReturnType<typeof createRemoteJWKSet>;
  private config: HttpConfig;
  private validIssuers: string[];
  private validAudiences: string[];

  constructor(config: HttpConfig) {
    this.config = config;
    // The tenant-scoped JWKS endpoint serves keys for both v1 and v2 tokens.
    this.jwks = createRemoteJWKSet(
      new URL(`https://login.microsoftonline.com/${config.tenantId}/discovery/v2.0/keys`)
    );
    this.validIssuers = [
      `https://login.microsoftonline.com/${config.tenantId}/v2.0`, // v2 tokens
      `https://sts.windows.net/${config.tenantId}/`,               // v1 tokens
    ];
    this.validAudiences = [config.clientId, `api://${config.clientId}`, config.appIdUri];
  }

  async verify(authorizationHeader: string | undefined): Promise<VerifiedUser> {
    if (!authorizationHeader || !authorizationHeader.toLowerCase().startsWith("bearer ")) {
      throw new AuthError(401, "missing_token", "Authorization: Bearer <token> header is required");
    }
    const token = authorizationHeader.slice(7).trim();

    let payload: JWTPayload;
    try {
      const result = await jwtVerify(token, this.jwks, {
        issuer: this.validIssuers,
        audience: this.validAudiences,
      });
      payload = result.payload;
    } catch (e: any) {
      logger.error("Token validation failed", e?.message || e);
      throw new AuthError(401, "invalid_token", `Token validation failed: ${e?.message || "unknown error"}`);
    }

    const tid = payload["tid"] as string | undefined;
    if (tid !== this.config.tenantId) {
      throw new AuthError(401, "invalid_token", "Token tenant does not match this server's tenant");
    }

    const oid = payload["oid"] as string | undefined;
    if (!oid) {
      throw new AuthError(401, "invalid_token", "Token missing oid claim");
    }

    // Delegated tokens carry 'scp'. App-only tokens carry 'roles' — we reject
    // those by design: this endpoint is delegated-only so every call maps to a human.
    const scp = (payload["scp"] as string | undefined) || "";
    const scopes = scp.split(" ").filter(Boolean);
    if (!scopes.includes(this.config.scopeName)) {
      throw new AuthError(
        403,
        "insufficient_scope",
        `Token must contain the '${this.config.scopeName}' scope (delegated). App-only tokens are not accepted.`
      );
    }

    if (this.config.allowedUserOids.length > 0 && !this.config.allowedUserOids.includes(oid)) {
      throw new AuthError(403, "access_denied", "User is not on the allowlist for this server");
    }

    if (this.config.allowedGroupId) {
      const groups = (payload["groups"] as string[] | undefined) || [];
      if (!groups.includes(this.config.allowedGroupId)) {
        throw new AuthError(
          403,
          "access_denied",
          "User is not a member of the required group (ensure the app registration emits the 'groups' claim)"
        );
      }
    }

    return {
      oid,
      tid,
      name: payload["name"] as string | undefined,
      preferredUsername: payload["preferred_username"] as string | undefined,
      scopes,
      rawToken: token,
      expiresAt: (payload.exp as number) || 0,
    };
  }
}

export class AuthError extends Error {
  status: number;
  code: string;
  constructor(status: number, code: string, message: string) {
    super(message);
    this.status = status;
    this.code = code;
  }
}
