/**
 * On-Behalf-Of (OBO) token broker.
 *
 * The bearer token Claude presents has audience = this app. To call Microsoft
 * Graph / Azure RM we exchange it, per user, for a downstream token using the
 * OAuth 2.0 On-Behalf-Of flow. This preserves delegation end to end:
 *   - the downstream token carries the user's identity and only their consented
 *     delegated permissions
 *   - Conditional Access applies
 *   - Graph sign-in logs attribute activity to the user, with this app as client
 *
 * MSAL caches OBO results keyed by assertion, so repeated tool calls within a
 * session don't round-trip to Entra every time.
 */

import { ConfidentialClientApplication, Configuration } from "@azure/msal-node";
import { readFileSync } from "fs";
import { createPrivateKey, createHash, X509Certificate } from "crypto";
import { logger } from "../logger.js";
import type { HttpConfig } from "./config.js";

export const GRAPH_SCOPE = "https://graph.microsoft.com/.default";
export const AZURE_RM_SCOPE = "https://management.azure.com/.default";

export class OboBroker {
  private cca: ConfidentialClientApplication;

  constructor(config: HttpConfig) {
    const msalConfig: Configuration = {
      auth: {
        clientId: config.clientId,
        authority: `https://login.microsoftonline.com/${config.tenantId}`,
      },
    };

    if (config.clientSecret) {
      msalConfig.auth.clientSecret = config.clientSecret;
    } else if (config.certificatePath) {
      // PEM file containing both the certificate and its private key.
      const pem = readFileSync(config.certificatePath, "utf8");
      const cert = new X509Certificate(pem);
      const thumbprint = createHash("sha256")
        .update(cert.raw)
        .digest("hex")
        .toUpperCase();
      // Validate the private key parses; MSAL takes the PEM string.
      createPrivateKey(pem);
      msalConfig.auth.clientCertificate = {
        thumbprintSha256: thumbprint,
        privateKey: pem,
      };
    } else {
      throw new Error("OboBroker requires CLIENT_SECRET or CERTIFICATE_PATH");
    }

    this.cca = new ConfidentialClientApplication(msalConfig);
  }

  /**
   * Exchange the user's inbound assertion for a downstream access token.
   */
  async getDownstreamToken(userAssertion: string, scope: string): Promise<string> {
    try {
      const result = await this.cca.acquireTokenOnBehalfOf({
        oboAssertion: userAssertion,
        scopes: [scope],
        skipCache: false,
      });
      if (!result?.accessToken) {
        throw new Error("OBO exchange returned no access token");
      }
      return result.accessToken;
    } catch (e: any) {
      logger.error(`OBO exchange failed for scope ${scope}`, e?.message || e);
      // Surface consent errors usefully — the most common first-run failure.
      const msg: string = e?.errorMessage || e?.message || String(e);
      if (msg.includes("AADSTS65001") || msg.includes("consent")) {
        throw new Error(
          `Graph consent required: an administrator must grant the app's delegated Graph permissions ` +
          `(admin consent) before On-Behalf-Of exchange can succeed. Original error: ${msg}`
        );
      }
      throw new Error(`On-Behalf-Of token exchange failed: ${msg}`);
    }
  }
}
