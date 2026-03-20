const SPREADSHEET_ID = "1R-DKI16mwZT6nIGWdo858NEWCh3XOwZRHfAmIs142o8";

// ── JWT Auth ─────────────────────────────────────────────────────────────────
async function getAccessToken() {
  const credentials = JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT);
  const now = Math.floor(Date.now() / 1000);

  const header  = { alg: "RS256", typ: "JWT" };
  const payload = {
    iss:   credentials.client_email,
    scope: "https://www.googleapis.com/auth/spreadsheets.readonly",
    aud:   "https://oauth2.googleapis.com/token",
    exp:   now + 3600,
    iat:   now,
  };

  const encode = (obj) => Buffer.from(JSON.stringify(obj)).toString("base64url");
  const signingInput = `${encode(header)}.${encode(payload)}`;

  const crypto = await import("crypto");
  const sign   = crypto.createSign("RSA-SHA256");
  sign.update(signingInput);
  const signature = sign.sign(credentials.private_key, "base64url");
  const jwt = `${signingInput}.${signature}`;

  const tokenRes = await fetch("https://oauth2.googleapis.com/token", {
    method:  "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body:    new URLSearchParams({
      grant_type: "urn:ietf:params:oauth:grant-type:jwt-bearer",
      assertion:  jwt,
    }),
  });

  const tokenData = await tokenRes.json();
  if (!tokenData.access_token) throw new Error("Token error: " + JSON.stringify(tokenData));
  return tokenData.access_token;
}

// ── Fetch one sheet tab ──────────────────────────────────────────────────────
async function fetchTab(accessToken, tabName) {
  const range = encodeURIComponent(`'${tabName}'!A1:G1000`);
  const url   = `https://sheets.googleapis.com/v4/spreadsheets/${SPREADSHEET_ID}/values/${range}`;
  const res   = await fetch(url, {
    headers: { Authorization: `Bearer ${accessToken}` },
  });
  const data = await res.json();
  if (!data.values) throw new Error(`No values for tab "${tabName}": ` + JSON.stringify(data));
  return data.values;
}

// ── Parse rows into deals ────────────────────────────────────────────────────
function parseRows(rows) {
  if (rows.length < 3) return [];

  // Row 0 = team title, Row 1 = headers, Row 2+ = data
  const headers = (rows[1] || []).map(h => (h || "").trim().toLowerCase());

  const ci = (...pats) => {
    for (const p of pats) {
      const i = headers.findIndex(h => p.test(h));
      if (i >= 0) return i;
    }
    return -1;
  };

  const iSales = ci(/sales.?name/i, /commercial/i);
  const iLead  = ci(/lead.?name/i,  /client/i);
  const iAct   = ci(/last.?activity/i, /activit/i);
  const iNext  = ci(/next.?step/i,  /prochaine/i);
  const iFore  = ci(/forecast/i,    /clos/i);
  const iCom   = ci(/comment/i,     /note/i);
  const iMRR   = ci(/mrr/i);

  const deals = [];

  for (let r = 2; r < rows.length; r++) {
    const row  = rows[r] || [];
    const g    = (i) => i >= 0 ? String(row[i] || "").trim() : "";
    const lead = g(iLead);
    if (!lead) continue;

    deals.push({
      id:            r,
      salesName:     g(iSales),
      leadName:      lead,
      lastActivity:  toISO(g(iAct)),
      nextStep:      g(iNext) || "—",
      forecastClose: toISO(g(iFore)),
      comment:       g(iCom),
      mrr:           toMRR(g(iMRR)),
    });
  }

  return deals;
}

function toISO(val) {
  if (!val) return "";
  if (/^\d{4}-\d{2}-\d{2}/.test(val)) return val.slice(0, 10);
  const m = val.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/);
  if (m) return `${m[3]}-${m[2].padStart(2,"0")}-${m[1].padStart(2,"0")}`;
  const n = parseFloat(val);
  if (!isNaN(n) && n > 40000) {
    return new Date(Math.round((n - 25569) * 86400000)).toISOString().slice(0, 10);
  }
  return val;
}

function toMRR(val) {
  if (!val) return null;
  const n = parseFloat(val.replace(/[^\d.,]/g, "").replace(",", "."));
  return isNaN(n) ? null : n;
}

// ── Handler ──────────────────────────────────────────────────────────────────
export default async function handler(req, res) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "GET");
  res.setHeader("Cache-Control", "no-store, max-age=0");

  try {
    const token = await getAccessToken();

    const [hospiRows, rsslRows] = await Promise.all([
      fetchTab(token, "Hospi"),
      fetchTab(token, "RSSL"),
    ]);

    res.status(200).json({
      Hospi:     parseRows(hospiRows),
      RSSL:      parseRows(rsslRows),
      updatedAt: new Date().toISOString(),
    });
  } catch (err) {
    console.error("deals.js error:", err);
    res.status(500).json({ error: err.message });
  }
}
