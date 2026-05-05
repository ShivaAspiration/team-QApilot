const axios = require("axios");
const XLSX = require("xlsx");
const fs = require("fs");
const path = require("path");

const DEFAULTS = {
  url: "https://api.qapilot.io/patapi/reports/report/v1/download-alerts",
  startdate: "2026-04-27",
  enddate: "2026-04-30",
  alerttype: "2",
  output: "alerts.xlsx",
};

const COLUMN_ORDER = [
  "Alert Type",
  "Tenant ID",
  "Message",
  "Tenant Name",
  "Project Na",
  "OS",
  "Service",
  "Device",
  "Step Title",
  "Testcase Title",
  "Testcase ID",
  "Step ID",
  "Report Trans ID",
  "Test Plan",
  "Report Status",
  "Total Testcount",
  "Success Count",
  "Failed Cases count",
  "Skipped Cases count",
  "QA Comments",
];

const COLUMN_ALIASES = {
  alerttype: "Alert Type",
  alerttypeid: "Alert Type",
  tenantid: "Tenant ID",
  tenantname: "Tenant Name",
  projectname: "Project Na",
  projectna: "Project Na",
  os: "OS",
  service: "Service",
  device: "Device",
  steptitle: "Step Title",
  testcase_t: "Testcase Title",
  testcaseid: "Testcase ID",
  testcasei: "Testcase ID",
  stepid: "Step ID",
  reporttra: "Report Trans ID",
  reporttransid: "Report Trans ID",
  testplan: "Test Plan",
  reportsta: "Report Status",
  totaltestc: "Total Testcount",
  totaltestcount: "Total Testcount",
  successc: "Success Count",
  successcount: "Success Count",
  failedcas: "Failed Cases count",
  failedcases: "Failed Cases count",
  skippedca: "Skipped Cases count",
  skippedcases: "Skipped Cases count",
  pass: "QA Comments",
  "qa comments": "QA Comments",
  qacomments: "QA Comments",
};

function parseArgs() {
  const args = process.argv.slice(2);
  const parsed = { ...DEFAULTS };

  args.forEach((arg) => {
    const [key, value] = arg.split("=");
    if (!value) return;
    const normalized = key.replace(/^--/, "");
    parsed[normalized] = value;
  });

  parsed.token = parsed.token || process.env.QAPILOT_AUTH_TOKEN;
  parsed.validateKey = parsed["validate-key"] || process.env.QAPILOT_VALIDATE_KEY;
  return parsed;
}

function findArray(data) {
  if (Array.isArray(data)) return data;
  if (data && typeof data === "object") {
    for (const key of Object.keys(data)) {
      if (Array.isArray(data[key])) return data[key];
    }
    return [data];
  }
  return [];
}

function normalizeKey(key) {
  return key
    .trim()
    .toLowerCase()
    .replace(/[_\s]/g, "")
    .replace(/[^a-z0-9]/g, "");
}

function renameRowColumns(row) {
  const renamed = {};
  Object.entries(row).forEach(([key, value]) => {
    const normalizedKey = normalizeKey(key);
    const alias = COLUMN_ALIASES[normalizedKey] || COLUMN_ALIASES[key.toLowerCase()] || key;
    if (renamed.hasOwnProperty(alias)) {
      if (!renamed[alias] && value) {
        renamed[alias] = value;
      }
    } else {
      renamed[alias] = value;
    }
  });
  return renamed;
}

function reorderRow(row) {
  const ordered = {};
  const seen = new Set();

  COLUMN_ORDER.forEach((column) => {
    ordered[column] = column in row ? row[column] : "";
    seen.add(normalizeKey(column));
  });

  Object.keys(row).forEach((key) => {
    const normalized = normalizeKey(key);
    if (!seen.has(normalized)) {
      ordered[key] = row[key];
      seen.add(normalized);
    }
  });

  return ordered;
}

function normalizeResponse(data) {
  if (typeof data === "string") {
    const trimmed = data.trim();
    if (trimmed.startsWith("{") || trimmed.startsWith("[")) {
      try {
        data = JSON.parse(trimmed);
      } catch (error) {
        console.warn("Warning: failed to parse JSON response; falling back to raw text.", error.message);
      }
    } else if (trimmed.includes("\n") && trimmed.includes(",")) {
      const workbook = XLSX.read(data, { type: "string" });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const parsed = XLSX.utils.sheet_to_json(sheet, { defval: "" });
      return parsed.map((row) => reorderRow(renameRowColumns(row)));
    }
  }

  const list = findArray(data);
  return list.map((row) => reorderRow(renameRowColumns(row)));
}

async function downloadAlerts(config) {
  const headers = {
    "x-validate": config.validateKey,
    Authorization: `Bearer ${config.token}`,
    Accept: "application/json, text/csv, */*",
  };

  const params = {
    startdate: config.startdate,
    enddate: config.enddate,
    alerttype: config.alerttype,
  };

  const response = await axios.get(config.url, {
    headers,
    params,
    timeout: 30000,
    responseType: "text",
  });
  return {
    data: response.data,
    contentType: response.headers["content-type"] || "",
  };
}

function saveToExcel(rows, outputPath) {
  const workbook = XLSX.utils.book_new();
  const worksheet = XLSX.utils.json_to_sheet(rows, { header: COLUMN_ORDER });
  XLSX.utils.book_append_sheet(workbook, worksheet, "Alerts");
  const fullPath = path.resolve(outputPath);
  XLSX.writeFile(workbook, fullPath);
  return fullPath;
}

async function main() {
  try {
    const config = parseArgs();
    if (!config.token || !config.validateKey) {
      console.error("Error: Authorization token and x-validate key are required.");
      console.error("Use --token=<token> and --validate-key=<value>, or set QAPILOT_AUTH_TOKEN and QAPILOT_VALIDATE_KEY.");
      process.exit(1);
    }

    const response = await downloadAlerts(config);
    console.log("Response Content-Type:", response.contentType);
    const rows = normalizeResponse(response.data);
    if (!rows.length) {
      console.log("No alerts returned from API.");
      process.exit(0);
    }

    console.log(`Parsed ${rows.length} rows.`);
    console.log("First row keys:", Object.keys(rows[0]).slice(0, 20).join(", "));

    const savedPath = saveToExcel(rows, config.output);
    console.log(`Saved ${rows.length} records to ${savedPath}`);
  } catch (error) {
    if (error.response) {
      console.error("API error:", error.response.status, error.response.statusText);
      console.error(error.response.data);
    } else {
      console.error("Request failed:", error.message);
    }
    process.exit(1);
  }
}

module.exports = {
  downloadAlerts,
  normalizeResponse,
  saveToExcel,
};

if (require.main === module) {
  main();
}
