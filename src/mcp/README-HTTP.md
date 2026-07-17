# Lokka HTTP — remote MCP with delegated Entra OAuth

Fork of [merill/lokka](https://github.com/merill/lokka) adding a native **Streamable HTTP**
transport protected by **delegated OAuth against your own Entra ID tenant**. Nobody can use
this server without signing in to your tenant; every Graph/Azure call runs **as the signed-in
user** via the On-Behalf-Of (OBO) flow.

Stdio mode (`build/main.js`) is untouched — this fork only adds `src/http/`.

## Security model

```
Claude (MCP client)
   │  bearer token, audience = YOUR app (api://<client-id>/access_as_user)
   ▼
Lokka HTTP  ──  validates JWT (JWKS, issuer, audience, tenant, scope, allowlists)
   │  On-Behalf-Of exchange (confidential client: secret or certificate)
   ▼
Microsoft Graph / Azure RM  ──  delegated token: user's own permissions,
                                Conditional Access enforced, sign-in logs
                                attribute activity to the human
```

- **App-only tokens are rejected by design** (`roles` without `scp`) — every call maps to a person.
- Sessions are bound to the user's `oid`; a stolen session id without that user's token is useless.
- Effective Graph permissions = intersection of (app's delegated permissions) ∩ (user's own privileges).
  A helpdesk tech connecting to this server cannot do Global-Admin things even if the app has broad scopes.

## 1. Entra app registration (one time)

1. **Entra admin center → App registrations → New registration**
   - Name: `Lokka HTTP`
   - Supported account types: *Accounts in this organizational directory only*
2. **Expose an API**
   - Set Application ID URI: `api://<client-id>`
   - Add a scope: name `access_as_user`, admins and users can consent,
     display/description e.g. "Access Lokka as the signed-in user".
3. **Manifest**: set `"requestedAccessTokenVersion": 2` (v2 tokens; the server accepts v1 too, but v2 is cleaner).
4. **API permissions → Microsoft Graph → Delegated** — start read-only, e.g.:
   `User.Read`, `Directory.Read.All`, `Policy.Read.All`,
   `DeviceManagementConfiguration.Read.All`, `DeviceManagementManagedDevices.Read.All`,
   `Reports.Read.All`, `AuditLog.Read.All`
   → **Grant admin consent**. (Add `https://management.azure.com/user_impersonation` under
   Azure Service Management if you want the `azure` apiType.)
5. **Authentication → Add a platform → Web** — redirect URI used by your MCP client's OAuth flow.
   For Claude custom connectors this is Anthropic's callback (check current value in
   Anthropic's connector docs, historically `https://claude.ai/api/mcp/auth_callback`).
6. **Certificates & secrets** — create a **client secret** (or better, upload a certificate and use
   `CERTIFICATE_PATH`). Required for the OBO exchange regardless of how clients authenticate.
7. Optional hardening:
   - **Enterprise application → Properties → Assignment required = Yes**, then assign only the
     users/group who may use Lokka. This blocks sign-in at Entra before the server even sees a token.
   - Emit the `groups` claim (Token configuration) if you want to use `ALLOWED_GROUP_ID`.
   - Conditional Access: require compliant device / MFA for this app specifically.

## 2. Deploy

```bash
docker build -t lokka-http .
docker run -d --name lokka-http -p 3000:3000 \
  -e TENANT_ID=<tenant-guid> \
  -e CLIENT_ID=<app-client-id> \
  -e CLIENT_SECRET=<secret> \            # or CERTIFICATE_PATH=/certs/lokka.pem (mount it)
  -e BASE_URL=https://lokka.example.com \
  -e ALLOWED_GROUP_ID=<optional-group-guid> \
  -e ALLOWED_USER_OIDS=<optional,csv,of,oids> \
  -e USE_GRAPH_BETA=true \
  lokka-http
```

Put TLS in front (Caddy/Traefik/nginx or a Cloudflare tunnel). **Never** expose plain HTTP.
`BASE_URL` must be the public HTTPS URL — it is baked into the OAuth resource metadata.

Endpoints:

| Path | Purpose |
|---|---|
| `POST/GET/DELETE /mcp` | MCP Streamable HTTP (bearer token required) |
| `/.well-known/oauth-protected-resource` | RFC 9728 metadata → points clients at Entra |
| `/healthz` | Liveness (unauthenticated) |

## 3. Connect from Claude

Settings → Connectors → **Add custom connector** → URL: `https://lokka.example.com/mcp`.

Entra does **not** support OAuth Dynamic Client Registration, so supply the app's
**client ID** (and secret) in the connector's advanced/OAuth settings rather than relying on DCR.
On connect, each user is sent through the normal Entra sign-in (MFA/CA included) and consents to
`access_as_user`. Tokens are refreshed by the client; the server refreshes its OBO assertion on
every request.

## 4. Environment reference

| Variable | Required | Notes |
|---|---|---|
| `TENANT_ID` | ✔ | Your Entra tenant GUID |
| `CLIENT_ID` | ✔ | App registration client id |
| `CLIENT_SECRET` / `CERTIFICATE_PATH` | one of | Confidential client credential for OBO (cert = PEM with key + cert) |
| `BASE_URL` | ✔ (prod) | Public HTTPS base URL |
| `PORT` | | Default 3000 |
| `SCOPE_NAME` | | Default `access_as_user` |
| `ALLOWED_USER_OIDS` | | CSV of user object ids permitted (empty = anyone who can sign in/consent) |
| `ALLOWED_GROUP_ID` | | Require this group id in the token's `groups` claim |
| `ALLOWED_ORIGINS` | | Browser-origin allowlist for CORS (server-side clients send no Origin) |
| `USE_GRAPH_BETA` | | `false` forces v1.0 (default beta allowed, matching upstream) |

## Known limitations

- **Sessions are in-memory** — run a single instance or use sticky routing. Fine for a small team.
- Built with esbuild transpile in the container; run `npx tsc --noEmit` locally for type-checking.
- The interactive-mode management tools from upstream (`add-graph-permission`, `set-access-token`, sign-in flows)
  are intentionally absent in HTTP mode: identity comes from the OAuth layer, and permission changes belong
  in the Entra portal with admin consent, not in-band.
- Token lifetime is ~60–90 min; the MCP client is responsible for refreshing and re-presenting tokens
  (the server re-reads the bearer on every request, so this is transparent when the client behaves).
