const LokkaClientId = "a9bac4c3-af0d-4292-9453-9da89e390140";
const LokkaDefaultTenantId = "common";
const LokkaDefaultRedirectUri = "http://localhost:3000";
const getDefaultGraphApiVersion = () => {
  return process.env.USE_GRAPH_BETA !== "false" ? "beta" : "v1.0";
};
export {
  LokkaClientId,
  LokkaDefaultRedirectUri,
  LokkaDefaultTenantId,
  getDefaultGraphApiVersion
};
