const { ConfidentialClientApplication } = require("@azure/msal-node");
const config = require("./config");

let msalClient = null;

function getMsalClient() {
  if (!msalClient) {
    msalClient = new ConfidentialClientApplication({
      auth: {
        clientId: config.MicrosoftAppId || "",
        clientSecret: config.MicrosoftAppPassword || "",
        authority: `https://login.microsoftonline.com/${config.MicrosoftAppTenantId || "common"}`,
      },
    });
  }
  return msalClient;
}

// user credentials (OBO)
async function exchangeTokenForDataverseToken(ssoToken) {
  const dataverseUrl = (config.dataverseUrl || "").replace(/\/+$/, "");
  if (!dataverseUrl) {
    throw new Error("DATAVERSE_URL is not configured.");
  }

  const scopes = [`${dataverseUrl}/.default`];

  const client = getMsalClient();
  const result = await client.acquireTokenOnBehalfOf({
    oboAssertion: ssoToken,
    scopes,
  });

  if (!result || !result.accessToken) {
    throw new Error("Failed to exchange SSO token for Dataverse token.");
  }

  return result.accessToken;
}

// client credentials
async function getDataverseAppToken() {
  const dataverseUrl = (config.dataverseUrl || "").replace(/\/+$/, "");
  if (!dataverseUrl) {
    throw new Error("DATAVERSE_URL is not configured.");
  }

  const scopes = [`${dataverseUrl}/.default`];
  const client = getMsalClient();

  const result = await client.acquireTokenByClientCredential({
    scopes,
  });

  if (!result || !result.accessToken) {
    throw new Error("Failed to acquire Dataverse app token via client credentials.");
  }

  return result.accessToken;
}

module.exports = {
  exchangeTokenForDataverseToken,
  getDataverseAppToken,
};