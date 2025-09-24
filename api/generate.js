// api/generate.js
const fs = require("fs");
const path = require("path");
const createReport = require("docx-templates").default;

/* ---------------- CORS: stabil & permissiv ---------------- */
function setCors(req, res) {
  const origin = req.headers.origin || "*";
  res.setHeader("Access-Control-Allow-Origin", origin);
  res.setHeader("Vary", "Origin");
  res.setHeader("Access-Control-Allow-Methods", "POST,OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");
  res.setHeader("Access-Control-Max-Age", "86400");
  res.setHeader("Cache-Control", "no-store");
}

/* ---------------- Utils (defensiv & minimal-invasiv) ---------------- */

// Liste robust normalisieren
function toArrayFlexible(v) {
  if (v == null) return [];
  if (Array.isArray(v)) return v.map(x => String(x ?? "").trim()).filter(Boolean);
  if (typeof v !== "string") return [];        // WICHTIG: Nicht-Strings ignorieren
  const s = v.trim();
  if (!s) return [];
  if (s.includes("|")) return s.split("|").map(t => t.trim()).filter(Boolean);
  if (s.includes("\n")) return s.split("\n").map(t => t.trim()).filter(Boolean);
  return s.split(",").map(t => t.trim()).filter(Boolean);
}

// Pipes → Tabs (für echte 2-Spalten in Word)
function pipesToTabs(text) {
  if (text == null) return "";
  return String(text).replace(/\s*\|\s*/g, "\t");
}

// Host → pretty label
function hostToPretty(url) {
  try {
    const u = new URL(String(url));
    return u.hostname.replace(/^www\./, "");
  } catch {
    return String(url || "source");
  }
}

/* ---------------- Sanfte Payload-Normalisierung ---------------- */

function normalizeSourceMeta(payload) {
  if (!payload.source_link_pretty && payload.source_link) {
    payload.source_link_pretty = hostToPretty(payload.source_link);
  }
}

function ensureContentRich(payload) {
  // Für alle *_content (außer Abitur) → *_content_rich, wenn fehlt
  Object.keys(payload).forEach(k => {
    if (/_content$/i.test(k) && !/_word_box_content$/i.test(k)) {
      const base = k.replace(/_content$/i, "");
      if (/^abitur_/i.test(base)) return;
      const richKey = `${base}_content_rich`;
      if (!payload[richKey]) {
        payload[richKey] = String(payload[k] ?? ""); // String-Coercion
      }
    }
  });
}

function normalizeTwoColumnsAndTabs(payload) {
  // *_content_col1/_col2 → Tab-Zeilen zusammenführen
  Object.keys(payload).forEach(k => {
    if (/_content_col1$/i.test(k)) {
      const base = k.replace(/_content_col1$/i, "");
      const col1 = String(payload[k] ?? "");
      const col2 = String(payload[`${base}_content_col2`] ?? "");
      const left = col1 ? col1.split(/\n/) : [];
      const right = col2 ? col2.split(/\n/) : [];
      const n = Math.max(left.length, right.length);
      const rows = [];
      for (let i = 0; i < n; i++) {
        const l = (left[i] ?? "").trimEnd();
        const r = (right[i] ?? "").trimEnd();
        rows.push(`${l}\t${r}`);
      }
      const combined = rows.join("\n");
      if (!payload[`${base}_content`]) payload[`${base}_content`] = combined;
      if (!payload[`${base}_content_rich`]) payload[`${base}_content_rich`] = combined;
    }
  });

  // Zusätzlich: überall Pipes in *_content_rich in Tabs wandeln
  Object.keys(payload).forEach(k => {
    if (/_content_rich$/i.test(k)) {
      payload[k] = pipesToTabs(payload[k]);
    }
  });
}

function normalizeWordBoxes(payload) {
  Object.keys(payload).forEach(k => {
    if (/_word_box_content$/i.test(k)) {
      const base = k.replace(/_word_box_content$/i, "");
      const arr = toArrayFlexible(payload[k]);       // sicher zu Array normalisieren
      if (arr.length && !payload[`${base}_word_box_content_line`]) {
        payload[`${base}_word_box_content_line`] = arr.join("   |   ");
      }
      // immer als String zurückschreiben (kein Array/Objekt im Payload belassen)
      payload[k] = arr.join(" | ");
    }
  });
}

function deriveArticleParagraphRich(payload) {
  // Falls article_text_paragraphX als HTML kam, aber kein _rich existiert → minimal in Text
  for (let i = 1; i <= 16; i++) {
    const key = `article_text_paragraph${i}`;
    if (key in payload && !payload[`${key}_rich`]) {
      let s = String(payload[key] ?? "");
      s = s.replace(/<\/p>\s*/gi, "\n\n");   // Absätze
      s = s.replace(/<br\s*\/?>/gi, "\n");   // Zeilen
      s = s.replace(/<[^>]+>/g, "");         // Tags raus
      s = s.replace(/\r\n?/g, "\n").replace(/\n{3,}/g, "\n\n").trim();
      payload[`${key}_rich`] = s;
    }
  }
}

/* ---------------- HTTP Handler ---------------- */

module.exports = async (req, res) => {
  setCors(req, res);

  if (req.method === "OPTIONS") {
    res.status(204).end();
    return;
  }
  if (req.method !== "POST") {
    res.status(405).json({ error: "Method Not Allowed" });
    return;
  }

  try {
    let payload = req.body;
    if (typeof payload === "string") {
      try { payload = JSON.parse(payload); } catch { payload = {}; }
    }
    if (!payload || typeof payload !== "object") payload = {};

    // Headline Alias (harmlos)
    if (payload.headline_article && !payload.headline_artikel) payload.headline_artikel = payload.headline_article;
    if (payload.headline_artikel && !payload.headline_article) payload.headline_article = payload.headline_artikel;

    // 1) Meta absichern
    normalizeSourceMeta(payload);

    // 2) Artikel _rich (nur falls fehlt)
    deriveArticleParagraphRich(payload);

    // 3) Word-Boxes (Liste → Line), alles als Strings
    normalizeWordBoxes(payload);

    // 4) *_content ⇒ *_content_rich (außer Abitur)
    ensureContentRich(payload);

    // 5) 2-Spalten/Pipes → Tabs
    normalizeTwoColumnsAndTabs(payload);

    // Template laden
    const templateBuffer = fs.readFileSync(path.join(__dirname, "template.docx"));

    // DOCX generieren
    const docBuffer = await createReport({
      template: templateBuffer,
      data: payload,
      cmdDelimiter: ["{", "}"],
      processLineBreaksAsNewText: true,
      rejectNullish: false,
      errorHandler: (err) => {
        console.log("Template Error:", err);
        return "";
      },
    });

    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
    res.setHeader("Content-Disposition", 'attachment; filename="generated.docx"');
    res.status(200).send(Buffer.from(docBuffer));
  } catch (err) {
    console.error("Generate Error:", err);
    res.status(500).json({ error: err?.message || String(err) });
  }
};
