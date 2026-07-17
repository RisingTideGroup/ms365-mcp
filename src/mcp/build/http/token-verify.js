import { createRemoteJWKSet, jwtVerify } from "jose";
import { logger } from "../logger.js";
class TokenVerifier {
  jwks;
  config;
  validIssuers;
  validAudiences;
  constructor(config) {
    this.config = config;
    this.jwks = createRemoteJWKSet(
      new URL(`https://login.microsoftonline.com/${config.tenantId}/discovery/v2.0/keys`)
    );
    this.validIssuers = [
      `https://login.microsoftonline.com/${config.tenantId}/v2.0`,
      // v2 tokens
      `https://sts.windows.net/${config.tenantId}/`
      // v1 tokens
    ];
    this.validAudiences = [config.clientId, `api://${config.clientId}`, config.appIdUri];
  }
  async verify(authorizationHeader) {
    if (!authorizationHeader || !authorizationHeader.toLowerCase().startsWith("bearer ")) {
      throw new AuthError(401, "missing_token", "Authorization: Bearer <token> header is required");
    }
    const token = authorizationHeader.slice(7).trim();
    let payload;
    try {
      const result = await jwtVerify(token, this.jwks, {
        issuer: this.validIssuers,
        audience: this.validAudiences
      });
      payload = result.payload;
    } catch (e) {
      logger.error("Token validation failed", e?.message || e);
      throw new AuthError(401, "invalid_token", `Token validation failed: ${e?.message || "unknown error"}`);
    }
    const tid = payload["tid"];
    if (tid !== this.config.tenantId) {
      throw new AuthError(401, "invalid_token", "Token tenant does not match this server's tenant");
    }
    const oid = payload["oid"];
    if (!oid) {
      throw new AuthError(401, "invalid_token", "Token missing oid claim");
    }
    const scp = payload["scp"] || "";
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
      const groups = payload["groups"] || [];
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
      name: payload["name"],
      preferredUsername: payload["preferred_username"],
      scopes,
      rawToken: token,
      expiresAt: payload.exp || 0
    };
  }
}
class AuthError extends Error {
  status;
  code;
  constructor(status, code, message) {
    super(message);
    this.status = status;
    this.code = code;
  }
}
export {
  AuthError,
  TokenVerifier
};
