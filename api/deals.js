const SPREADSHEET_ID = "1R-DKI16mwZT6nIGWdo858NEWCh3XOwZRHfAmIs142o8";

// JWT Auth for Google API
async function getAccessToken() {
  const credentials = JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT);
  const now = Math.floor(Date.now() / 1000);

  const header = { alg: "RS256", typ: "JWT" };
  const payload = {
    iss: credentials.client_email,
    scope: "https://www.googleapis.com/auth/spreadsheets.readonly",
    aud: "https://oauth2.googleapis.com/token",
    exp: now + 3600,
    iat: now,
  };

  const encode = (obj) =>
    Buffer.from(JSON.stringify(obj)).toString("base64url");

  const signingInput = `${encode(header)}.${encode(payload)}`;

  // Import crypto for RS256 signing
  const crypto = await import("crypto");
  const privateKey = credentials.private_key;
  const sign = crypto.createSign("RSA-SHA256");
  sign.update(signingInput);
  const signature = sign.sign(privateKey, "base64url");

  const jwt = `${signingInput}.${signature}`;

  // Exchange JWT for access token
  const tokenRes = await fetch("https://oauth2.googleapis.com/token", {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({
      grant_type: "urn:ietf:params:oauth:grant-type:jwt-bearer",
      assertion: jwt,
    }),
  });

  const tokenData = await tokenRes.json();
  return tokenData.access_token;
}

// Fetch a sheet tab by name
async function fetchSheet(accessToken, sheetName) {
  const range = encodeURIComponent(`${sheetName}!A1:Z1000`);
  const url = `https://sheets.googleapis.com/v4/spreadsheets/${SPREADSHEET_ID}/values/${range}`;
  const res = await fetch(url, {
    headers: { Authorization: `Bearer ${accessToken}` },
  });
  const data = await res.json();
  return data.values || [];
}

// Parse rows into deal objects
function parseRows(rows) {
  if (rows.length < 2) return [];
  const headers = rows[0].map((h) => h.trim().toLowerCase());

  const idx = (patterns) => {
    for (const p of patterns) {
      const i = headers.findIndex((h) => p.test(h));
      if (i >= 0) return i;
    }
    return -1;
  };

  const iSales = idx([/sales.?name/i, /commercial/i, /vendeur/i]);
  const iLead  = idx([/lead.?name/i, /client/i, /compte/i, /société/i]);
  const iAct   = idx([/last.?activity/i, /derni.re.activit/i, /date.call/i]);
  const iNext  = idx([/next.?step/i, /prochaine/i, /action/i]);
  const iFore  = idx([/forecast/i, /closing/i, /cl.ture/i, /prévisionnel/i]);
  const iCom   = idx([/comment/i, /remarque/i, /note/i]);
  const iMRR   = idx([/mrr/i, /revenu/i]);

  return rows.slice(1).map((row, i) => {
    const g = (ix) => (ix >= 0 && row[ix]) ? row[ix].trim() : "";
    const lead = g(iLead);
    if (!lead) return null;
    return {
      id: i + 1,
      salesName:     g(iSales),
      leadName:      lead,
      lastActivity:  g(iAct),
      nextStep:      g(iNext) || "—",
      forecastClose: g(iFore),
      comment:       g(iCom),
      mrr:           parseMRR(g(iMRR)),
    };
  }).filter(Boolean);
}

function parseMRR(s) {
  if (!s) return null;
  const n = parseFloat(s.replace(/[^\d.,]/g, "").replace(",", "."));
  return isNaN(n) ? null : n;
}

// Main handler
export default async function handler(req, res) {
  // CORS
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "GET");
  res.setHeader("Cache-Control", "no-store");

  try {
    const accessToken = await getAccessToken();

    // Fetch both sheets in parallel
    const [rsslRows, hospiRows] = await Promise.all([
      fetchSheet(accessToken, "RSSL"),
      fetchSheet(accessToken, "Hospi"),
    ]);

    const data = {
      RSSL:  parseRows(rsslRows),
      Hospi: parseRows(hospiRows),
      updatedAt: new Date().toISOString(),
    };

    res.status(200).json(data);
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
}
