const config = require("./config");

const API_VERSION = "v9.2";

// Replace these placeholders with your real Dataverse entity set names.
const TABLE_DEFINITIONS = [
  {
    key: "table1",
    entitySet: config.dataverseEntitySetTable1,
    description: [
      "Products table (gl_products).",
      "Key fields: gl_productid, gl_name (product name), gl_productcode, gl_description, statecode (0=Active), createdon.",
      "Use for: listing products, filtering by product name or code.",
    ].join(" "),
  },
  {
    key: "table2",
    entitySet: config.dataverseEntitySetTable2,
    description: [
      "SKUs table (gl_skus).",
      "Key fields: gl_skuid, gl_skuname, gl_skucode, gl_unitofmeasure, gl_status, _gl_productid_value (linked product), statecode (0=Active), createdon.",
      "Use for: listing SKUs, filtering by SKU name/code, looking up SKUs for a product.",
    ].join(" "),
  },
  {
    key: "table3",
    entitySet: config.dataverseEntitySetTable3,
    description: [
      "Rate Cards table (gl_ratecards).",
      "Key fields: gl_ratecardid, gl_ratecard1 (name), gl_unitprice, gl_unitprice_base, gl_currency, gl_effectivestartdate, gl_effectiveenddate, _gl_skucode_value (linked SKU), statecode (0=Active).",
      "Use for: pricing queries, looking up unit price for a SKU, effective date ranges.",
    ].join(" "),
  },
  {
    key: "table4",
    entitySet: config.dataverseEntitySetTable4,
    description: [
      "License Consumption Snapshots table (gl_licenseconsumptionsnapshots).",
      "Key fields: gl_licenseconsumptionsnapshotid, gl_name, gl_snapshotdate, gl_assignedlicenses, gl_consumedlicenses, gl_overage (overage count), gl_overageamount (overage cost/value), gl_rate, _gl_account_value (customer/account), _gl_sku_value (linked SKU), _gl_product_value (linked product), _gl_ratecard_value (linked rate card), statecode (0=Active).",
      "COMPUTED columns (cannot use in $filter — use postFilter instead): gl_overageamount, gl_overage, gl_overageamount_base, gl_rate_base.",
      "Filterable columns (safe in $filter): statecode, gl_snapshotdate, gl_consumedlicenses, gl_assignedlicenses, _gl_account_value, _gl_sku_value, _gl_product_value.",
      "Use for: overage queries, license consumption, top customers by overage, purchase amount thresholds. Sort by gl_overageamount desc for highest overage.",
    ].join(" "),
  },
];

const TABLE_BY_KEY = Object.fromEntries(TABLE_DEFINITIONS.map((t) => [t.key, t]));

function logDataverseQuery(functionName, comment, success, error) {
  const timestamp = new Date().toISOString();
  if (success) {
    console.log(`[${timestamp}] ✅ DATAVERSE SUCCESS | Function: ${functionName} | Comment: ${comment}`);
  } else {
    console.error(`[${timestamp}] ❌ DATAVERSE FAIL | Function: ${functionName} | Comment: ${comment} | Error: ${error}`);
  }
}

function getApiBaseUrl() {
  const base = config.dataverseUrl.replace(/\/+$/, "");
  return `${base}/api/data/${API_VERSION}`;
}


async function fetchWithRetry(url, options, maxRetries = 3) {
  /** @type {Error|null} */
  let lastError = null;
  
  for (let attempt = 1; attempt <= maxRetries; attempt++) {
    try {
      const response = await fetch(url, options);
      return response;
    } catch (error) {
      lastError = error;
      const isRetryable =
        error?.code === "ECONNRESET" ||
        error?.code === "ETIMEDOUT" ||
        error?.code === "ENOTFOUND" ||
        error?.message?.includes("fetch failed") ||
        error?.cause?.code === "ECONNRESET";
      
      if (!isRetryable || attempt === maxRetries) {
        throw error;
      }
      
      // Exponential backoff: 500ms, 1000ms, 2000ms...
      const delayMs = Math.min(500 * Math.pow(2, attempt - 1), 5000);
      console.warn(
        `[Dataverse] Retry ${attempt}/${maxRetries} after ${delayMs}ms due to: ${error?.code || error?.message}`
      );
      await new Promise((resolve) => setTimeout(resolve, delayMs));
    }
  }

  throw lastError || new Error("fetchWithRetry: unexpected state");
}

/*Generic helper to execute a single-page GET request against the Dataverse Web API.*/
async function queryDataverseTable(accessToken, entitySet, queryOptions) {
  const baseUrl = getApiBaseUrl();
  const url = queryOptions
    ? `${baseUrl}/${entitySet}?${queryOptions}`
    : `${baseUrl}/${entitySet}`;

  const comment = `Query entity set '${entitySet}'`;

  const headers = {
    Authorization: `Bearer ${accessToken}`,
    "OData-MaxVersion": "4.0",
    "OData-Version": "4.0",
    Accept: "application/json",
    "Content-Type": "application/json; charset=utf-8",
    Prefer: 'odata.include-annotations="*"',
  };

  try {
    const response = await fetchWithRetry(url, { method: "GET", headers });

    if (!response.ok) {
      const errorBody = await response.text();
      const errorMsg = `Dataverse API error (${response.status}): ${errorBody}`;
      logDataverseQuery("queryDataverseTable", comment, false, errorMsg);
      throw new Error(errorMsg);
    }

    const data = await response.json();
    logDataverseQuery("queryDataverseTable", comment, true);
    return data;
  } catch (error) {
    logDataverseQuery("queryDataverseTable", comment, false, error?.message);
    throw error;
  }
}

/*
 * Fetch ALL records from a Dataverse entity set by following @odata.nextLink
 * pages automatically until no more pages remain.
 */
async function queryDataverseTableAll(accessToken, entitySet, queryOptions) {
  const baseUrl = getApiBaseUrl();
  let nextUrl = queryOptions
    ? `${baseUrl}/${entitySet}?${queryOptions}`
    : `${baseUrl}/${entitySet}`;

  const headers = {
    Authorization: `Bearer ${accessToken}`,
    "OData-MaxVersion": "4.0",
    "OData-Version": "4.0",
    Accept: "application/json",
    "Content-Type": "application/json; charset=utf-8",
    Prefer: 'odata.include-annotations="*",odata.maxpagesize=5000',
  };

  const allRecords = [];
  let page = 0;

  while (nextUrl) {
    page++;
    const response = await fetchWithRetry(nextUrl, { method: "GET", headers });

    if (!response.ok) {
      const errorBody = await response.text();
      const errorMsg = `Dataverse API error (${response.status}): ${errorBody}`;
      logDataverseQuery("queryDataverseTableAll", `Page ${page} of '${entitySet}'`, false, errorMsg);
      throw new Error(errorMsg);
    }

    const data = await response.json();
    const records = data.value || [];
    allRecords.push(...records);

    logDataverseQuery(
      "queryDataverseTableAll",
      `Page ${page} of '${entitySet}' — fetched ${records.length} (total: ${allRecords.length})`,
      true
    );

    nextUrl = data["@odata.nextLink"] || null;
  }
  console.log(allRecords);
  return allRecords;
}

/*
 * Fetch records from any Dataverse table.
 * Pass top=null (or omit) to fetch ALL records via automatic OData pagination.
 * Pass a number to limit to that many records (single page).
 */
async function getTableRecords(accessToken, tableName, filter, top, select, orderBy) {
  const comment = `Fetch records from '${tableName}' (filter=${filter || "none"}, top=${top ?? "all"})`;

  try {
    let queryOptions = "";

    if (top != null) {
      queryOptions += `$top=${top}`;
    }

    if (select) {
      queryOptions += `${queryOptions ? "&" : ""}$select=${select}`;
    }

    if (orderBy) {
      queryOptions += `${queryOptions ? "&" : ""}$orderby=${encodeURIComponent(orderBy)}`;
    }

    if (filter) {
      queryOptions += `${queryOptions ? "&" : ""}$filter=${encodeURIComponent(filter)}`;
    }

    let records;
    if (top != null) {
      // single-page bounded fetch
      const data = await queryDataverseTable(accessToken, tableName, queryOptions);
      records = data.value || [];
    } else {
      // unbounded paginated fetch — retrieves every record
      records = await queryDataverseTableAll(accessToken, tableName, queryOptions);
    }

    logDataverseQuery("getTableRecords", comment, true);
    return records;
  } catch (error) {
    logDataverseQuery("getTableRecords", comment, false, error?.message);
    throw error;
  }
}

function getDataverseTableDefinitions() {
  return TABLE_DEFINITIONS;
}

function getDataverseTableByKey(tableKey) {
  return TABLE_BY_KEY[tableKey];
}

module.exports = {
  fetchWithRetry,
  getApiBaseUrl,
  getDataverseTableByKey,
  getDataverseTableDefinitions,
  logDataverseQuery,
  getTableRecords,
  queryDataverseTable,
  queryDataverseTableAll,
};