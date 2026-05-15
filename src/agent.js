const { ActivityTypes } = require("@microsoft/agents-activity");
const { AgentApplication, MemoryStorage } = require("@microsoft/agents-hosting");
const { AzureOpenAI } = require("openai");

const config = require("./config");
const { getDataverseAppToken } = require("./tokenService");
const {
  getDataverseTableByKey,
  getDataverseTableDefinitions,
  getTableRecords,
  queryDataverseTableAggregate,
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
    '- Schema: {"queries":[{"tableKey":"table1|table2|table3|table4","filter":"optional odata $filter","select":"optional comma list","orderBy":"optional order by","apply":"optional odata $apply aggregation expression"}]}',
    "Rules:",
    "- Include only relevant table(s).",
    "- Do not include top. Always fetch the full matching table/query result and derive the answer from that full dataset.",
    "- If the question is a pure aggregate (e.g. sum, count, average, groupby), set 'apply' to a valid OData $apply expression (e.g. 'aggregate(gl_overage with sum as totalOverage)' or 'groupby((gl_name),aggregate(gl_overage with sum as totalOverage))'). When 'apply' is set, omit 'select' and 'orderBy'.",
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
      apply: q?.apply || null,
    }))
    .filter((q) => getDataverseTableByKey(q.tableKey));
}

const MAX_ROWS_TO_LLM = 300;

function stripAnnotations(record) {
  const clean = {};
  for (const key of Object.keys(record)) {
    if (!key.includes("@")) {
      clean[key] = record[key];
    }
  }
  return clean;
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

    // --- Strategy 1: OData $apply server-side aggregation ---
    if (query.apply) {
      const aggregateRows = await queryDataverseTableAggregate(
        accessToken,
        table.entitySet,
        query.apply,
        query.filter
      );
      tableData.push({
        tableKey: table.key,
        entitySet: table.entitySet,
        aggregated: true,
        rowCount: aggregateRows.length,
        rows: aggregateRows.map(stripAnnotations),
      });
      continue;
    }

    // --- Strategy 2: Full fetch with map-reduce if over limit ---
    let rows = await getTableRecords(
      accessToken,
      table.entitySet,
      query.filter,
      query.top,
      query.select,
      query.orderBy
    );

    const totalCount = rows.length;
    rows = rows.map(stripAnnotations);

    if (rows.length > MAX_ROWS_TO_LLM) {
      // Map-reduce: summarize chunks, then push the combined summary
      const summary = await summarizeInChunks(userQuery, table.key, rows);
      tableData.push({
        tableKey: table.key,
        entitySet: table.entitySet,
        totalCount,
        mapReduceSummary: true,
        summary,
      });
    } else {
      tableData.push({
        tableKey: table.key,
        entitySet: table.entitySet,
        totalCount,
        rowCount: rows.length,
        rows,
      });
    }
  }

  return { plannedQueries, tableData };
}

const CHUNK_SIZE = 250;

async function summarizeInChunks(userQuery, tableKey, rows) {
  const chunks = [];
  for (let i = 0; i < rows.length; i += CHUNK_SIZE) {
    chunks.push(rows.slice(i, i + CHUNK_SIZE));
  }

  const partialSummaries = [];
  for (let i = 0; i < chunks.length; i++) {
    const chunk = chunks[i];
    const result = await client.chat.completions.create({
      messages: [
        {
          role: "system",
          content:
            "You are a data analyst. Extract and summarize only facts directly relevant to the user question from the provided records. Be concise — output key numbers, names, and findings. Do not explain the data structure.",
        },
        {
          role: "user",
          content: `User question: ${userQuery}\n\nTable: ${tableKey} (chunk ${i + 1}/${chunks.length}, ${chunk.length} records):\n${JSON.stringify(chunk)}`,
        },
      ],
      model: "",
    });
    partialSummaries.push(result.choices?.[0]?.message?.content || "");
  }

  // Reduce: synthesize partial summaries into one
  if (partialSummaries.length === 1) return partialSummaries[0];

  const reduceResult = await client.chat.completions.create({
    messages: [
      {
        role: "system",
        content:
          "You are a data analyst. Combine the following partial summaries into a single cohesive summary relevant to the user question. Resolve duplicates and aggregate numbers where possible.",
      },
      {
        role: "user",
        content: `User question: ${userQuery}\n\nPartial summaries:\n${partialSummaries.map((s, i) => `[${i + 1}] ${s}`).join("\n\n")}`,
      },
    ],
    model: "",
  });

  return reduceResult.choices?.[0]?.message?.content || partialSummaries.join(" | ");
}

async function generateAnswerFromData(userQuery, dataverseContext) {
  const answerPrompt = [
    "Use Dataverse results as primary source.",
    "If data is empty, say that clearly and suggest next useful question.",
    "Do not invent unavailable fields or records.",
    "If a table has mapReduceSummary=true, the 'summary' field contains a pre-processed condensed summary of all records — use it as the source of truth.",
    "If a table has aggregated=true, the 'rows' field contains OData aggregate results — use them directly.",
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