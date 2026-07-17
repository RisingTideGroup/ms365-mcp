import { ConfidentialClientApplication } from "@azure/msal-node";
import { readFileSync } from "fs";
import { createPrivateKey, createHash, X509Certificate } from "crypto";
import { logger } from "../logger.js";
const GRAPH_SCOPE = "https://graph.microsoft.com/.default";
const AZURE_RM_SCOPE = "https://management.azure.com/.default";
class OboBroker {
  cca;
  constructor(config) {
    const msalConfig = {
      auth: {
        clientId: config.clientId,
        authority: `https://login.microsoftonline.com/${config.tenantId}`
      }
    };
    if (config.clientSecret) {
      msalConfig.auth.clientSecret = config.clientSecret;
    } else if (config.certificatePath) {
      const pem = readFileSync(config.certificatePath, "utf8");
      const cert = new X509Certificate(pem);
      const thumbprint = createHash("sha256").update(cert.raw).digest("hex").toUpperCase();
      createPrivateKey(pem);
      msalConfig.auth.clientCertificate = {
        thumbprintSha256: thumbprint,
        privateKey: pem
      };
    } else {
      throw new Error("OboBroker requires CLIENT_SECRET or CERTIFICATE_PATH");
    }
    this.cca = new ConfidentialClientApplication(msalConfig);
  }
  /**
   * Exchange the user's inbound assertion for a downstream access token.
   */
  async getDownstreamToken(userAssertion, scope) {
    try {
      const result = await this.cca.acquireTokenOnBehalfOf({
        oboAssertion: userAssertion,
        scopes: [scope],
        skipCache: false
      });
      if (!result?.accessToken) {
        throw new Error("OBO exchange returned no access token");
      }
      return result.accessToken;
    } catch (e) {
      logger.error(`OBO exchange failed for scope ${scope}`, e?.message || e);
      const msg = e?.errorMessage || e?.message || String(e);
      if (msg.includes("AADSTS65001") || msg.includes("consent")) {
        throw new Error(
          `Graph consent required: an administrator must grant the app's delegated Graph permissions (admin consent) before On-Behalf-Of exchange can succeed. Original error: ${msg}`
        );
      }
      throw new Error(`On-Behalf-Of token exchange failed: ${msg}`);
    }
  }
}
export {
  AZURE_RM_SCOPE,
  GRAPH_SCOPE,
  OboBroker
};
