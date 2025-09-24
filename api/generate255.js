// api/generate.js
const fs = require("fs");
const path = require("path");
const createReport = require("docx-templates").default;

/* ---------------- CORS: stabil & permissiv ---------------- */
function setCors(req, res) {
  const origin = req.headers.origin || "*";
  res.setHeader("Access-Control-Allow-Origin", origin);
  res.setHeader("Vary", "Origin");
  // Wichtig: Keine Credentials in Kombi mit "*"
  // res.setHeader("Access-Control-Allow-Credentials", "true"); // NICHT setzen
  res.setHeader("Access-Control-Allow-Methods", "POST,OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");
  res.setHeader("Access-Control-Max-Age", "86400");
  res.setHeader("Cache-Control", "no-store");
}

/* ---------------- Kleine Utils (non-invasiv) ---------------- */

// Normalisiert eine Liste (Array/String) zu Array<string>
function toArrayFlexible(v) {
  if (v == null) return [];
  if (Array.isArray(v)) return v.map(x => String(x).trim()).filter(Boolean);
  const s = String(v);
  if (s.includes("|")) return s.split("|").map(t => t.trim()).filter(Boolean);
  if (s.includes("\n")) return s.split("\n").map(t => t.trim()).filter(Boolean);
  return s.split(",").map(t => t.trim()).filter(Boolean);
}

// Konvertiere "left   |   right" in Tab-getrennte Zeile für echte Spalten in Word
function pipesToTabs(text) {
  if (text == null) return "";
  return String(text).replace(/\s*\|\s*/g, "\t");
}

// Kleines Hilfsmittel: Hostname extrahieren → pretty label
function hostToPretty(url) {
  try {
    const u = new URL(String(url));
    return u.hostname.replace(/^www\./, "");
  } catch {
    return String(url || "source");
  }
}

/* ---------------- Payload-Normalisierung (minimal-invasiv) ---------------- */

function normalizeSourceMeta(payload) {
  // Falls Frontend nichts gesetzt hat, nichts kaputt machen
  if (typeof payload.source_link_pretty === "undefined" && payload.source_link) {
    payload.source_link_pretty = hostToPretty(payload.source_link);
  }
}

function ensureContentRich(payload) {
  // Für alle *_content (außer Abitur) ein *_content_rich spiegeln, falls fehlt
  Object.keys(payload).forEach(k => {
    if (/_content$/i.test(k) && !/_word_box_content$/i.test(k)) {
      const base = k.replace(/_content$/i, "");
      const isAbitur = /^abitur_/i.test(base);
      if (isAbitur) return;
      const richKey = `${base}_content_rich`;
      if (!payload[richKey]) {
        payload[richKey] = payload[k];
      }
    }
  });
}

function normalizeTwoColumnsAndTabs(payload) {
  // Wenn *_content_col1/_col2 existieren, daraus *_content_rich als Tab-Zeilen erzeugen
  Object.keys(payload).forEach(k => {
    if (/_content_col1$/i.test(k)) {
      const base = k.replace(/_content_col1$/i, "");
      const col1 = String(payload[k] || "");
      const col2 = String(payload[`${base}_content_col2`] || "");
      const left = col1.split(/\n/);
      const right = col2.split(/\n/);
      const n = Math.max(left.length, right.length);
      const rows = [];
      for (let i = 0; i < n; i++) {
        const l = (left[i] || "").trimEnd();
        const r = (right[i] || "").trimEnd();
        // Tab als echte Spaltentrennung in Word
        rows.push(`${l}\t${r}`);
      }
      const combined = rows.join("\n");
      // Setze *_content und *_content_rich, falls nicht schon vorhanden
      if (!payload[`${base}_content`]) payload[`${base}_content`] = combined;
      if (!payload[`${base}_content_rich`]) payload[`${base}_content_rich`] = combined;
    }
  });

  // Zusätzlich: überall dort, wo Pipe-Zeilen im Content stehen, in Tabs umwandeln
  Object.keys(payload).forEach(k => {
    if (/_content_rich$/i.test(k)) {
      payload[k] = pipesToTabs(payload[k]);
    }
  });
}

function normalizeWordBoxes(payload) {
  // Wenn nur *_word_box_content (Liste) vorhanden, *_word_box_content_line hinzufügen
  Object.keys(payload).forEach(k => {
    if (/_word_box_content$/i.test(k)) {
      const base = k.replace(/_word_box_content$/i, "");
      const arr = toArrayFlexible(payload[k]);
      if (arr.length && !payload[`${base}_word_box_content_line`]) {
        payload[`${base}_word_box_content_line`] = arr.join("   |   ");
      }
    }
  });
}

function deriveArticleParagraphRich(payload) {
  // Falls Frontend article_text_paragraphX als HTML geschickt hat, aber kein _rich vorhanden ist:
  for (let i = 1; i <= 16; i++) {
    const key = `article_text_paragraph${i}`;
    if (key in payload && !payload[`${key}_rich`]) {
      // Minimal: HTML-Tags grob entfernen und Zeilenumbrüche vereinheitlichen – nicht aggressiv
      let s = String(payload[key] || "");
      // Absatzenden → Doppel-Umbruch
      s = s.replace(/<\/p>\s*/gi, "\n\n");
      // Zeilenumbrüche
      s = s.replace(/<br\s*\/?>/gi, "\n");
      // Rest-Tags strippen
      s = s.replace(/<[^>]+>/g, "");
      // Whitespace normalisieren
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

    // Sanfte Defaults – nichts überschreiben, was vom Frontend kommt
    if (payload.headline_article && !payload.headline_artikel) payload.headline_artikel = payload.headline_article;
    if (payload.headline_artikel && !payload.headline_article) payload.headline_article = payload.headline_artikel;

    // 1) Meta absichern
    normalizeSourceMeta(payload);

    // 2) Artikel (optional) _rich ableiten, falls Frontend nur HTML geschickt hat
    deriveArticleParagraphRich(payload);

    // 3) Word-Boxes hübsch machen (Line bauen, falls nur Liste kam)
    normalizeWordBoxes(payload);

    // 4) Für alle außer Abitur: *_content ⇒ *_content_rich spiegeln (falls fehlt)
    ensureContentRich(payload);

    // 5) 2-Spalten / Pipes → Tabs (für echtes Nebeneinander)
    normalizeTwoColumnsAndTabs(payload);

    // Template laden
    const templateBuffer = fs.readFileSync(path.join(__dirname, "template.docx"));

    // DOCX generieren
    const docBuffer = await createReport({
      template: templateBuffer,
      data: payload,
      cmdDelimiter: ["{", "}"],
      processLineBreaksAsNewText: true,
      // Wichtig: nullish nicht rejecten → optionale Blöcke im Template greifen sauber
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
