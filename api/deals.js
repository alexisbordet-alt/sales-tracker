const SPREADSHEET_ID = "1R-DKI16mwZT6nIGWdo858NEWCh3XOwZRHfAmIs142o8";

async function getAccessToken() {
  const credentials = JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT);
  const now = Math.floor(Date.now() / 1000);
  const header  = { alg: "RS256", typ: "JWT" };
  const payload = {
    iss: credentials.client_email,
    scope: "https://www.googleapis.com/auth/spreadsheets.readonly",
    aud: "https://oauth2.googleapis.com/token",
    exp: now + 3600,
    iat: now,
  };
  const encode = (obj) => Buffer.from(JSON.stringify(obj)).toString("base64url");
  const si = `${encode(header)}.${encode(payload)}`;
  const crypto = await import("crypto");
  const sign = crypto.createSign("RSA-SHA256");
  sign.update(si);
  const sig = sign.sign(credentials.private_key, "base64url");
  const jwt = `${si}.${sig}`;
  const r = await fetch("https://oauth2.googleapis.com/token", {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({ grant_type: "urn:ietf:params:oauth:grant-type:jwt-bearer", assertion: jwt }),
  });
  const t = await r.json();
  if (!t.access_token) throw new Error("Token: " + JSON.stringify(t));
  return t.access_token;
}

async function fetchTab(token, tab) {
  // Use A1:Z1000 and valueRenderOption=UNFORMATTED_VALUE to get raw numbers
  const range = encodeURIComponent(`'${tab}'!A1:Z1000`);
  const url = `https://sheets.googleapis.com/v4/spreadsheets/${SPREADSHEET_ID}/values/${range}?valueRenderOption=UNFORMATTED_VALUE&dateTimeRenderOption=FORMATTED_STRING`;
  const r = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
  const d = await r.json();
  if (!d.values) throw new Error(`No values "${tab}": ` + JSON.stringify(d));
  return d.values;
}

function parseTab(rows) {
  if (rows.length < 2) return { deals: [], objectif: null, mrrClosed: 0, mrrPrev: 0 };

  // Row 0 = title, Row 1 = headers
  const headers = (rows[1] || []).map(h => String(h || "").trim().toLowerCase());

  // Find column indexes by header name
  const ci = (...pats) => {
    for (const p of pats) {
      const i = headers.findIndex(h => p.test(h));
      if (i >= 0) return i;
    }
    return -1;
  };

  const iSales = ci(/sales.?name/i);
  const iLead  = ci(/lead.?name/i, /client/i);
  const iAct   = ci(/last.?activity/i);
  const iNext  = ci(/^next.?step$/i);
  const iNDate = ci(/next.?step.?date/i);
  const iFore  = ci(/forecast/i, /closing/i);
  const iCom   = ci(/comment/i);
  const iMRR   = ci(/^mrr$/i);
  const iState = ci(/^state$/i, /statut/i);
  const iObj   = ci(/objectif/i, /target/i, /goal/i);

  // Get objectif from first data row (row index 2)
  let objectif = null;
  if (iObj >= 0) {
    for (let r = 2; r < rows.length; r++) {
      const v = (rows[r] || [])[iObj];
      const n = parseFloat(String(v || "").replace(/[^0-9.,]/g, "").replace(",", "."));
      if (!isNaN(n) && n > 0) { objectif = n; break; }
    }
  }

  let mrrClosed = 0, mrrPrev = 0;

  const deals = rows.slice(2).map((row, i) => {
    const g = idx => idx >= 0 ? String(row[idx] ?? "").trim() : "";

    const lead = g(iLead);
    if (!lead) return null;

    // MRR: parse as float directly from cell value (already unformatted)
    let mrr = null;
    if (iMRR >= 0 && row[iMRR] !== undefined && row[iMRR] !== "") {
      const v = parseFloat(String(row[iMRR]).replace(/[^0-9.,]/g, "").replace(",", "."));
      if (!isNaN(v) && v > 0) mrr = v;
    }

    const stateRaw = g(iState).toLowerCase();
    const closed = stateRaw === "closed" || stateRaw === "gagné" || stateRaw === "won";

    if (mrr) { if (closed) mrrClosed += mrr; else mrrPrev += mrr; }

    return {
      id: i + 3,
      salesName:     g(iSales),
      leadName:      lead,
      lastActivity:  toISO(g(iAct)),
      nextStep:      g(iNext) || "—",
      nextStepDate:  toISO(g(iNDate)),
      forecastClose: toISO(g(iFore)),
      comment:       g(iCom),
      mrr,
      state: closed ? "closed" : "open",
    };
  }).filter(Boolean);

  return { deals, objectif, mrrClosed, mrrPrev };
}

function toISO(val) {
  if (!val) return "";
  // Already ISO
  if (/^\d{4}-\d{2}-\d{2}/.test(val)) return val.slice(0, 10);
  // DD/MM/YYYY
  const m = val.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/);
  if (m) return `${m[3]}-${m[2].padStart(2,"0")}-${m[1].padStart(2,"0")}`;
  // Google serial number
  const n = parseFloat(val);
  if (!isNaN(n) && n > 40000)
    return new Date(Math.round((n - 25569) * 86400000)).toISOString().slice(0, 10);
  return val;
}

export default async function handler(req, res) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Cache-Control", "no-store");
  try {
    const token = await getAccessToken();
    const [hr, rr] = await Promise.all([fetchTab(token, "Hospi"), fetchTab(token, "RSSL")]);
    const h = parseTab(hr);
    const r = parseTab(rr);

    res.status(200).json({
      Hospi:                h.deals,
      RSSL:                 r.deals,
      objectifHospi:        h.objectif,
      objectifRSSL:         r.objectif,
      mrrClosedHospi:       h.mrrClosed,
      mrrPrevisionnelHospi: h.mrrPrev,
      mrrClosedRSSL:        r.mrrClosed,
      mrrPrevisionnelRSSL:  r.mrrPrev,
      // Debug info
      _debug: {
        hospiHeaders: hr[1] || [],
        rsslHeaders:  rr[1] || [],
        hospiRow3:    hr[2] || [],
        rsslRow3:     rr[2] || [],
      },
      updatedAt: new Date().toISOString(),
    });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
}
