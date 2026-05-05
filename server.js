const express = require('express');
const fs = require('fs');
const path = require('path');
const { downloadAlerts, normalizeResponse, saveToExcel } = require('./download_alerts_to_excel');

const app = express();
const PORT = process.env.PORT || 3001;
const downloadsDir = path.join(__dirname, 'public', 'downloads');

if (!fs.existsSync(downloadsDir)) {
  fs.mkdirSync(downloadsDir, { recursive: true });
}

app.use(express.urlencoded({ extended: true }));
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

/* ================= HELPERS ================= */

function getTenantName(row) {
  const keys = ['Tenant Name', 'tenantname', 'Tenant', 'tenant'];

  for (const key of keys) {
    if (row[key] && String(row[key]).trim()) {
      return String(row[key]).trim();
    }
  }

  return 'Unknown Tenant';
}

function getAlertType(row) {
  return row['Message'] || row['message'] || 'Unknown Alert';
}

/* ================= SUMMARY LOGIC ================= */

function buildSummaryTable(rows) {
  const summary = {};
  const tenants = new Set();

  rows.forEach((row) => {
    const alertType = getAlertType(row);
    const tenant = getTenantName(row);

    tenants.add(tenant);

    if (!summary[alertType]) {
      summary[alertType] = {};
    }

    if (!summary[alertType][tenant]) {
      summary[alertType][tenant] = 0;
    }

    summary[alertType][tenant]++;
  });

  const messages = Object.keys(summary);

  return {
    messages,
    tenants: Array.from(tenants).sort(),
    data: summary,
  };
}

/* ================= API ================= */

app.post('/preview', async (req, res) => {
  const { startdate, enddate, alerttype, token, validateKey } = req.body;

  if (!token || !validateKey) {
    return res.status(400).json({ error: 'Token & validateKey required' });
  }

  try {
    const response = await downloadAlerts({
      url: 'https://api.qapilot.io/patapi/reports/report/v1/download-alerts',
      startdate,
      enddate,
      alerttype,
      token,
      validateKey,
    });

    const rows = normalizeResponse(response.data);

    if (!rows.length) {
      return res.json({
        summary: { messages: [], tenants: [], data: {} },
        rawData: [],
        columns: []   // ✅ FIX
      });
    }

    const summary = buildSummaryTable(rows);
    const columns = Object.keys(rows[0]); // ✅ FIX

    res.json({
      summary,
      rawData: rows,
      columns   // ✅ FIX (IMPORTANT)
    });

  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

/* ================= DOWNLOAD ================= */

app.post('/download', async (req, res) => {
  const { startdate, enddate, alerttype, token, validateKey } = req.body;

  try {
    const response = await downloadAlerts({
      url: 'https://api.qapilot.io/patapi/reports/report/v1/download-alerts',
      startdate,
      enddate,
      alerttype,
      token,
      validateKey,
    });

    const rows = normalizeResponse(response.data);

    const file = `alerts_${Date.now()}.xlsx`;
    const filePath = path.join(downloadsDir, file);

    saveToExcel(rows, filePath);

    res.json({
      downloadUrl: `/downloads/${file}`,
    });

  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.use('/downloads', express.static(downloadsDir));

app.listen(PORT, () => {
  console.log(`Server running at http://localhost:${PORT}`);
});