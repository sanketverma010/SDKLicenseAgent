const config = {
  azureOpenAIKey: process.env.AZURE_OPENAI_API_KEY,
  azureOpenAIEndpoint: process.env.AZURE_OPENAI_ENDPOINT,
  azureOpenAIDeploymentName: process.env.AZURE_OPENAI_DEPLOYMENT_NAME,
  dataverseUrl: process.env.DATAVERSE_URL,
  dataverseEntitySetTable1: process.env.DATAVERSE_TABLE_1,
  dataverseEntitySetTable2: process.env.DATAVERSE_TABLE_2,
  dataverseEntitySetTable3: process.env.DATAVERSE_TABLE_3,
  dataverseEntitySetTable4: process.env.DATAVERSE_TABLE_4 ,
  MicrosoftAppId: process.env.BOT_ID,
  MicrosoftAppPassword: process.env.BOT_PASSWORD,
  MicrosoftAppTenantId: process.env.MICROSOFT_APP_TENANT_ID,
};

module.exports = config;
