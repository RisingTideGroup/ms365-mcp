import { ClientSecretCredential, ClientCertificateCredential, InteractiveBrowserCredential, DeviceCodeCredential } from "@azure/identity";
import jwt from "jsonwebtoken";
import { logger } from "./logger.js";
import { LokkaClientId, LokkaDefaultTenantId, LokkaDefaultRedirectUri } from "./constants.js";
const ONE_HOUR_IN_MS = 60 * 60 * 1e3;
function parseJwtScopes(token) {
  try {
    const decoded = jwt.decode(token);
    if (!decoded || typeof decoded !== "object") {
      logger.info("Failed to decode JWT token");
      return [];
    }
    const scopesString = decoded.scp;
    if (typeof scopesString === "string") {
      return scopesString.split(" ").filter((scope) => scope.length > 0);
    }
    const roles = decoded.roles;
    if (Array.isArray(roles)) {
      return roles;
    }
    logger.info("No scopes found in JWT token");
    return [];
  } catch (error) {
    logger.error("Error parsing JWT token for scopes", error);
    return [];
  }
}
class TokenCredentialAuthProvider {
  credential;
  constructor(credential) {
    this.credential = credential;
  }
  async getAccessToken() {
    const token = await this.credential.getToken("https://graph.microsoft.com/.default");
    if (!token) {
      throw new Error("Failed to acquire access token");
    }
    return token.token;
  }
}
class ClientProvidedTokenCredential {
  accessToken;
  expiresOn;
  constructor(accessToken, expiresOn) {
    if (accessToken) {
      this.accessToken = accessToken;
      this.expiresOn = expiresOn || new Date(Date.now() + ONE_HOUR_IN_MS);
    } else {
      this.expiresOn = /* @__PURE__ */ new Date(0);
    }
  }
  async getToken(scopes) {
    if (!this.accessToken || !this.expiresOn || this.expiresOn <= /* @__PURE__ */ new Date()) {
      logger.error("Access token is not available or has expired");
      return null;
    }
    return {
      token: this.accessToken,
      expiresOnTimestamp: this.expiresOn.getTime()
    };
  }
  updateToken(accessToken, expiresOn) {
    this.accessToken = accessToken;
    this.expiresOn = expiresOn || new Date(Date.now() + ONE_HOUR_IN_MS);
    logger.info("Access token updated successfully");
  }
  isExpired() {
    return !this.expiresOn || this.expiresOn <= /* @__PURE__ */ new Date();
  }
  getExpirationTime() {
    return this.expiresOn || /* @__PURE__ */ new Date(0);
  }
  // Getter for access token (for internal use by AuthManager)
  getAccessToken() {
    return this.accessToken;
  }
}
var AuthMode = /* @__PURE__ */ ((AuthMode2) => {
  AuthMode2["ClientCredentials"] = "client_credentials";
  AuthMode2["ClientProvidedToken"] = "client_provided_token";
  AuthMode2["Interactive"] = "interactive";
  AuthMode2["Certificate"] = "certificate";
  return AuthMode2;
})(AuthMode || {});
class AuthManager {
  credential = null;
  config;
  constructor(config) {
    this.config = config;
  }
  async initialize() {
    switch (this.config.mode) {
      case "client_credentials" /* ClientCredentials */:
        if (!this.config.tenantId || !this.config.clientId || !this.config.clientSecret) {
          throw new Error("Client credentials mode requires tenantId, clientId, and clientSecret");
        }
        logger.info("Initializing Client Credentials authentication");
        this.credential = new ClientSecretCredential(
          this.config.tenantId,
          this.config.clientId,
          this.config.clientSecret
        );
        break;
      case "client_provided_token" /* ClientProvidedToken */:
        logger.info("Initializing Client Provided Token authentication");
        this.credential = new ClientProvidedTokenCredential(
          this.config.accessToken,
          this.config.expiresOn
        );
        break;
      case "certificate" /* Certificate */:
        if (!this.config.tenantId || !this.config.clientId || !this.config.certificatePath) {
          throw new Error("Certificate mode requires tenantId, clientId, and certificatePath");
        }
        logger.info("Initializing Certificate authentication");
        this.credential = new ClientCertificateCredential(this.config.tenantId, this.config.clientId, {
          certificatePath: this.config.certificatePath,
          certificatePassword: this.config.certificatePassword
        });
        break;
      case "interactive" /* Interactive */:
        const tenantId = this.config.tenantId || LokkaDefaultTenantId;
        const clientId = this.config.clientId || LokkaClientId;
        logger.info(`Initializing Interactive authentication with tenant ID: ${tenantId}, client ID: ${clientId}`);
        try {
          this.credential = new InteractiveBrowserCredential({
            tenantId,
            clientId,
            redirectUri: this.config.redirectUri || LokkaDefaultRedirectUri
          });
        } catch (error) {
          logger.info("Interactive browser failed, falling back to device code flow");
          this.credential = new DeviceCodeCredential({
            tenantId,
            clientId,
            userPromptCallback: (info) => {
              console.log(`
\u{1F510} Authentication Required:`);
              console.log(`Please visit: ${info.verificationUri}`);
              console.log(`And enter code: ${info.userCode}
`);
              return Promise.resolve();
            }
          });
        }
        break;
      default:
        throw new Error(`Unsupported authentication mode: ${this.config.mode}`);
    }
    await this.testCredential();
  }
  updateAccessToken(accessToken, expiresOn) {
    if (this.config.mode === "client_provided_token" /* ClientProvidedToken */ && this.credential instanceof ClientProvidedTokenCredential) {
      this.credential.updateToken(accessToken, expiresOn);
    } else {
      throw new Error("Token update only supported in client provided token mode");
    }
  }
  async testCredential() {
    if (!this.credential) {
      throw new Error("Credential not initialized");
    }
    if (this.config.mode === "client_provided_token" /* ClientProvidedToken */ && !this.config.accessToken) {
      logger.info("Skipping initial credential test as no token was provided at startup.");
      return;
    }
    try {
      const token = await this.credential.getToken("https://graph.microsoft.com/.default");
      if (!token) {
        throw new Error("Failed to acquire token");
      }
      logger.info("Authentication successful");
    } catch (error) {
      logger.error("Authentication test failed", error);
      throw error;
    }
  }
  getGraphAuthProvider() {
    if (!this.credential) {
      throw new Error("Authentication not initialized");
    }
    return new TokenCredentialAuthProvider(this.credential);
  }
  getAzureCredential() {
    if (!this.credential) {
      throw new Error("Authentication not initialized");
    }
    return this.credential;
  }
  getAuthMode() {
    return this.config.mode;
  }
  isClientCredentials() {
    return this.config.mode === "client_credentials" /* ClientCredentials */;
  }
  isClientProvidedToken() {
    return this.config.mode === "client_provided_token" /* ClientProvidedToken */;
  }
  isInteractive() {
    return this.config.mode === "interactive" /* Interactive */;
  }
  async getTokenStatus() {
    if (this.credential instanceof ClientProvidedTokenCredential) {
      const tokenStatus = {
        isExpired: this.credential.isExpired(),
        expiresOn: this.credential.getExpirationTime()
      };
      if (!tokenStatus.isExpired) {
        const accessToken = this.credential.getAccessToken();
        if (accessToken) {
          try {
            const scopes = parseJwtScopes(accessToken);
            return {
              ...tokenStatus,
              scopes
            };
          } catch (error) {
            logger.error("Error parsing token scopes in getTokenStatus", error);
            return tokenStatus;
          }
        }
      }
      return tokenStatus;
    } else if (this.credential) {
      try {
        const accessToken = await this.credential.getToken("https://graph.microsoft.com/.default");
        if (accessToken && accessToken.token) {
          const scopes = parseJwtScopes(accessToken.token);
          return {
            isExpired: false,
            expiresOn: new Date(accessToken.expiresOnTimestamp),
            scopes
          };
        }
      } catch (error) {
        logger.error("Error getting token for scope parsing", error);
      }
    }
    return { isExpired: false };
  }
}
export {
  AuthManager,
  AuthMode,
  ClientProvidedTokenCredential,
  TokenCredentialAuthProvider
};
