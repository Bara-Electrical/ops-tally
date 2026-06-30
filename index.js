import express from "express";
import Airtable from "airtable";
import cron from "node-cron";
import crypto from "crypto";

const app = express();
app.use(express.json());

// ================================================================
// ENV
// ================================================================
const REQUIRED_ENV = [
  "GRAPH_TENANT_ID", "GRAPH_CLIENT_ID", "GRAPH_CLIENT_SECRET",
  "AIRTABLE_API_KEY", "AIRTABLE_BASE_ID",
  "AROFLO_ORG", "AROFLO_USER", "AROFLO_PASS",
  "SECRET_KEY",
];
for (const key of REQUIRED_ENV) {
  if (!process.env[key]) { console.error(`Missing env var: ${key}`); process.exit(1); }
}

const TEAMS_TEAM_ID    = "09353200-8046-4356-ae2f-2af74eb5a378";
const TEAMS_CHANNEL_ID = "19:NOuOBjEljSlXLjVe18oSlH7Z26QLFYNq3PICZrB3Fi01@thread.tacv2";

const airtableBase = new Airtable({ apiKey: process.env.AIRTABLE_API_KEY }).base(process.env.AIRTABLE_BASE_ID);

// ================================================================
// GRAPH API
// ================================================================
let graphToken = null;
let graphTokenExpiry = 0;

async function getGraphToken() {
  if (graphToken && Date.now() < graphTokenExpiry - 60000) return graphToken;
  const res = await fetch(
    `https://login.microsoftonline.com/${process.env.GRAPH_TENANT_ID}/oauth2/v2.0/token`,
    {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: new URLSearchParams({
        client_id:     process.env.GRAPH_CLIENT_ID,
        client_secret: process.env.GRAPH_CLIENT_SECRET,
        scope:         "https://graph.microsoft.com/.default",
        grant_type:    "client_credentials",
      }),
    }
  );
  const data = await res.json();
  if (!res.ok) throw new Error(`Graph auth failed: ${JSON.stringify(data)}`);
  graphToken       = data.access_token;
  graphTokenExpiry = Date.now() + data.expires_in * 1000;
  return graphToken;
}

async function graphFetch(path, options = {}) {
  const token = await getGraphToken();
  return fetch(`https://graph.microsoft.com/v1.0${path}`, {
    ...options,
    headers: {
      Authorization:  `Bearer ${token}`,
      "Content-Type": "application/json",
      ...options.headers,
    },
  });
}

// ================================================================
// AROFLO API
// ================================================================
function arofloAuthHeader() {
  const ts  = new Date().toUTCString();
  const sig = crypto.createHmac("sha1",
    Buffer.from(process.env.AROFLO_PASS, "base64")
  ).update(process.env.AROFLO_USER + ts).digest("base64");
  return {
    Authentication: `HMAC ${sig}`,
    afdatetimeutc:  ts,
    "Content-Type": "application/x-www-form-urlencoded",
  };
}

async function arofloGet(query) {
  const res  = await fetch(`https://api.aroflo.com/?${query}&org=${process.env.AROFLO_ORG}`, {
    headers: arofloAuthHeader(),
  });
  const data = await res.json();
  if (!res.ok) throw new Error(`Aroflo API error ${res.status}: ${JSON.stringify(data)}`);
  return data;
}

// ================================================================
// WEBHOOK — validate shared secret
// ================================================================
function validateSecret(req, res) {
  if (req.headers["x-secret-key"] !== process.env.SECRET_KEY) {
    res.status(401).json({ error: "Unauthorized" });
    return false;
  }
  return true;
}

// ================================================================
// WEBHOOK — invoice approved
// ================================================================
app.post("/webhook/invoice", async (req, res) => {
  if (!validateSecret(req, res)) return;
  res.sendStatus(200);
  try {
    const p             = req.body;
    const person        = p.assignedusername || p.createdbyusername || p.assigneduser || "Unknown";
    const amount        = parseFloat(p.totalamount ?? p.invoicetotal ?? p.amount ?? 0);
    const invoiceNumber = String(p.invoicenumber ?? p.id ?? "");
    const date          = (p.invoicedate ?? p.approveddate ?? new Date().toISOString()).slice(0, 10);

    await airtableBase("Invoice Log").create([{
      fields: {
        "Date":           date,
        "Person":         person,
        "Amount":         amount,
        "Invoice Number": invoiceNumber,
      },
    }]);
    console.log(`Invoice logged: ${person} $${amount} #${invoiceNumber} on ${date}`);
  } catch (err) {
    console.error("Invoice webhook error:", err.message);
  }
});

// ================================================================
// WEBHOOK — quote approved
// ================================================================
app.post("/webhook/quote", async (req, res) => {
  if (!validateSecret(req, res)) return;
  res.sendStatus(200);
  try {
    const p           = req.body;
    const person      = p.assignedusername || p.createdbyusername || p.assigneduser || "Unknown";
    const quoteNumber = String(p.quotenumber ?? p.id ?? "");
    const date        = (p.quotedate ?? p.approveddate ?? new Date().toISOString()).slice(0, 10);

    await airtableBase("Quote Log").create([{
      fields: {
        "Date":         date,
        "Person":       person,
        "Quote Number": quoteNumber,
      },
    }]);
    console.log(`Quote logged: ${person} #${quoteNumber} on ${date}`);
  } catch (err) {
    console.error("Quote webhook error:", err.message);
  }
});

// ================================================================
// SUMMARY LOGIC
// ================================================================
function getDateRange() {
  // Run at 7am Perth — calculate yesterday's range.
  // Monday: cover Fri + Sat + Sun (3 days back).
  const now        = new Date(new Date().toLocaleString("en-US", { timeZone: "Australia/Perth" }));
  const dayOfWeek  = now.getDay(); // 0=Sun 1=Mon
  const daysBack   = dayOfWeek === 1 ? 3 : 1;

  const to   = new Date(now); to.setDate(to.getDate() - 1);
  const from = new Date(now); from.setDate(from.getDate() - daysBack);

  const fmt = d => d.toLocaleDateString("en-CA"); // YYYY-MM-DD
  return { from: fmt(from), to: fmt(to) };
}

async function getInvoiceSummary(from, to) {
  const records = await airtableBase("Invoice Log").select({
    filterByFormula: `AND(Date >= '${from}', Date <= '${to}')`,
  }).all();

  const byPerson = {};
  let totalAmount = 0;
  for (const r of records) {
    const person = r.fields["Person"] || "Unknown";
    const amount = r.fields["Amount"] || 0;
    if (!byPerson[person]) byPerson[person] = { count: 0, amount: 0 };
    byPerson[person].count++;
    byPerson[person].amount += amount;
    totalAmount += amount;
  }
  return { byPerson, totalAmount, totalCount: records.length };
}

async function getQuoteSummary(from, to) {
  const records = await airtableBase("Quote Log").select({
    filterByFormula: `AND(Date >= '${from}', Date <= '${to}')`,
  }).all();

  const byPerson = {};
  for (const r of records) {
    const person = r.fields["Person"] || "Unknown";
    byPerson[person] = (byPerson[person] || 0) + 1;
  }
  return { byPerson, totalCount: records.length };
}

async function getClosedJobCount(from, to) {
  try {
    const fromEnc = encodeURIComponent(`and|completeddate|>=|${from}`);
    const toEnc   = encodeURIComponent(`and|completeddate|<=|${to}`);
    const statusEnc = encodeURIComponent(`and|statusname|=|Complete`);
    let count = 0, page = 1;
    while (true) {
      const zone = await arofloGet(
        `zone=tasks&where=${statusEnc}&where=${fromEnc}&where=${toEnc}&page=${page}`
      );
      const raw = zone?.tasks;
      if (!raw) break;
      const arr = Array.isArray(raw) ? raw : [raw];
      count += arr.length;
      if (arr.length < parseInt(zone.maxpageresults ?? 500)) break;
      page++;
    }
    return count;
  } catch (err) {
    console.error("Closed jobs query failed:", err.message);
    return null;
  }
}

function formatCurrency(amount) {
  return amount.toLocaleString("en-AU", { style: "currency", currency: "AUD" });
}

function formatDateLabel(from, to) {
  const opts = { weekday: "long", day: "numeric", month: "long" };
  const shortOpts = { weekday: "short", day: "numeric", month: "short" };
  const f = new Date(from + "T12:00:00");
  const t = new Date(to   + "T12:00:00");
  return from === to
    ? f.toLocaleDateString("en-AU", opts)
    : `${f.toLocaleDateString("en-AU", shortOpts)} – ${t.toLocaleDateString("en-AU", shortOpts)}`;
}

async function postDailySummary() {
  const { from, to } = getDateRange();
  console.log(`Posting summary for ${from} → ${to}`);

  const [invoices, quotes, closedJobs] = await Promise.all([
    getInvoiceSummary(from, to),
    getQuoteSummary(from, to),
    getClosedJobCount(from, to),
  ]);

  const invoiceLines = Object.entries(invoices.byPerson)
    .sort(([a], [b]) => a.localeCompare(b))
    .map(([p, { count }]) => `${p} - ${count}`)
    .join("<br>");

  const quoteLines = Object.entries(quotes.byPerson)
    .sort(([a], [b]) => a.localeCompare(b))
    .map(([p, count]) => `${p} - ${count}`)
    .join("<br>");

  const closedLine = closedJobs !== null ? String(closedJobs) : "unavailable";

  const html = [
    `<p><strong>📊 Daily Ops Summary — ${formatDateLabel(from, to)}</strong></p>`,
    `<p><strong>Invoices</strong><br>${invoiceLines || "None"}</p>`,
    `<p><strong>Total invoices - ${invoices.totalCount}</strong><br>${formatCurrency(invoices.totalAmount)}</p>`,
    `<p><strong>Quotes</strong><br>${quoteLines || "None"}</p>`,
    `<p><strong>Total quotes - ${quotes.totalCount}</strong></p>`,
    `<p><strong>Total closed jobs - ${closedLine}</strong></p>`,
  ].join("\n");

  const res = await graphFetch(
    `/teams/${TEAMS_TEAM_ID}/channels/${encodeURIComponent(TEAMS_CHANNEL_ID)}/messages`,
    { method: "POST", body: JSON.stringify({ body: { contentType: "html", content: html } }) }
  );

  if (!res.ok) {
    const err = await res.json();
    throw new Error(`Teams post failed: ${JSON.stringify(err)}`);
  }
  console.log("Summary posted to Teams");
}

// ================================================================
// CRON — 7am weekdays Perth time
// ================================================================
cron.schedule("0 7 * * 1-5", () => {
  postDailySummary().catch(err => console.error("Scheduled summary failed:", err.message));
}, { timezone: "Australia/Perth" });

// ================================================================
// MANUAL TRIGGER
// ================================================================
app.get("/run-summary", async (req, res) => {
  if (!validateSecret(req, res)) return;
  res.json({ status: "triggered" });
  postDailySummary().catch(err => console.error("Manual summary failed:", err.message));
});

// Debug: log raw webhook payload so we can map Aroflo field names
app.post("/webhook/debug", (req, res) => {
  if (!validateSecret(req, res)) return;
  console.log("WEBHOOK DEBUG:", JSON.stringify(req.body, null, 2));
  res.sendStatus(200);
});

// ================================================================
// START
// ================================================================
app.listen(process.env.PORT || 3000, () => {
  console.log(`Ops Tally running — ${new Date().toISOString()}`);
});
