const { ActivityTypes } = require("@microsoft/agents-activity");
const { AgentApplication, MemoryStorage } = require("@microsoft/agents-hosting");
const { AzureOpenAI } = require("openai");

const config = require("./config");
const { getDataverseAppToken } = require("./tokenService");
const {
  getDataverseTableByKey,
  getDataverseTableDefinitions,
  getTableRecords,
} = require("./dataverseClient");

const client = new AzureOpenAI({
  apiVersion: "2024-12-01-preview",
  apiKey: config.azureOpenAIKey,
  endpoint: config.azureOpenAIEndpoint,
  deployment: config.azureOpenAIDeploymentName,
});
const BASE_SYSTEM_PROMPT =
  "You are an AI agent that answers user questions clearly and concisely.";
const TABLE_DEFINITIONS = getDataverseTableDefinitions();

function parseJsonObject(text) {
  try {
    return JSON.parse(text);
  } catch {
    return null;
  }
}

async function planDataverseQueries(userQuery) {
  const plannerSystemPrompt = [
    "You choose which Dataverse tables to query for a user question.",
    "Output must be valid JSON only.",
    'Schema: {"queries":[{"tableKey":"table1|table2|table3|table4","filter":"optional odata $filter","select":"optional comma list","orderBy":"optional order by"}]}',
    "Rules:",
    "- Include only relevant table(s).",
    "- Do not include top. Always fetch the full matching table/query result and derive the answer from that full dataset.",
    "- Use empty queries array if no Dataverse lookup is needed.",
    `- Today's date is ${new Date().toISOString().split("T")[0]}. Use ISO 8601 format (e.g. 2026-04-01) for date values in OData $filter expressions. Never leave a date placeholder empty.`,
    "- OData $filter does NOT support arithmetic operators (/, *, +, -) or ratio/percentage expressions. Never write expressions like 'col1 / col2 gt 0.8'. For ratio-based conditions, omit the ratio from the filter entirely and rely on post-processing. Only use simple comparisons (eq, ne, gt, ge, lt, le) and logical operators (and, or, not) in $filter.",
    "- OData $filter requires the right-hand side of every comparison to be a constant literal (e.g. a number, string, or date). Never compare two columns against each other (e.g. 'col1 ge col2' is invalid). If the condition requires comparing two fields, omit it from the filter and handle it in post-processing.",
  ].join("\n");

  const tableContext = TABLE_DEFINITIONS.map(
    (t) => `tableKey="${t.key}" → entitySet="${t.entitySet}" | ${t.description}`
  ).join("\n");

  const result = await client.chat.completions.create({
    messages: [
      { role: "system", content: plannerSystemPrompt },
      {
        role: "user",
        content: `Available tables:\n${tableContext}\n\nUser query:\n${userQuery}`,
      },
    ],
    response_format: { type: "json_object" },
    model: "",
  });

  const raw = result.choices?.[0]?.message?.content || "{}";
  const parsed = parseJsonObject(raw) || {};
  const queries = Array.isArray(parsed.queries) ? parsed.queries : [];

  return queries
    .map((q) => ({
      tableKey: q?.tableKey,
      filter: q?.filter,
      top: null,
      select: q?.select,
      orderBy: q?.orderBy,
    }))
    .filter((q) => getDataverseTableByKey(q.tableKey));
}

async function fetchDataverseContext(userQuery) {
  const plannedQueries = await planDataverseQueries(userQuery);
  if (!plannedQueries.length) {
    return { plannedQueries, tableData: [] };
  }

  const accessToken = await getDataverseAppToken();
  const tableData = [];

  for (const query of plannedQueries) {
    const table = getDataverseTableByKey(query.tableKey);
    let rows = await getTableRecords(
      accessToken,
      table.entitySet,
      query.filter,
      query.top,
      query.select,
      query.orderBy
    );

    tableData.push({
      tableKey: table.key,
      entitySet: table.entitySet,
      rowCount: rows.length,
      rows,
    });
  }

  return { plannedQueries, tableData };
}

async function generateAnswerFromData(userQuery, dataverseContext) {
  const answerPrompt = [
    "Use Dataverse results as primary source.",
    "If data is empty, say that clearly and suggest next useful question.",
    "Do not invent unavailable fields or records.",
  ].join("\n");

  const result = await client.chat.completions.create({
    messages: [
      { role: "system", content: `${BASE_SYSTEM_PROMPT}\n${answerPrompt}` },
      {
        role: "user",
        content: `User query:\n${userQuery}\n\nDataverse context (JSON):\n${JSON.stringify(dataverseContext)}`,
      },
    ],
    model: "",
  });

  return result.choices?.map((c) => c.message?.content || "").join("").trim();
}

// Define storage and application
const storage = new MemoryStorage();
const agentApp = new AgentApplication({
  storage,
});

agentApp.onConversationUpdate("membersAdded", async (context) => {
  await context.sendActivity(`Hi there! I'm an agent to chat with you.`);
});

// Listen for ANY message to be received. MUST BE AFTER ANY OTHER MESSAGE HANDLERS
agentApp.onActivity(ActivityTypes.Message, async (context) => {
  const userQuery = context.activity.text || "";

  await context.sendActivity({ type: ActivityTypes.Typing });

  try {
    const dataverseContext = await fetchDataverseContext(userQuery);
    const answer = await generateAnswerFromData(userQuery, dataverseContext);

    if (!answer) {
      await context.sendActivity("I could not generate an answer right now. Please try again.");
      return;
    }

    await context.sendActivity(answer);
  } catch (error) {
    console.error("Agent handling error:", error?.message || error);
    await context.sendActivity(
      "I ran into an issue while querying Dataverse. Please verify Dataverse/auth settings and try again."
    );
  }
});

module.exports = {
  agentApp,
};