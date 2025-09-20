// api/generate.js
// Vercel Node.js Serverless Function (CommonJS) + docx-templates
// Ziele:
// - CORS + OPTIONS
// - Robust: Body-Parsing, Defaults für optionale Variablen, Normalisierung von Arrays/Objekten zu String
// - DOCX-Binary zurückgeben (korrekter MIME-Type), kein JSON-Wrapper
// - Fehler -> JSON

const fs = require("fs");
const path = require("path");
// WICHTIG: Default-Export der Library korrekt holen
const createReport = require("docx-templates").default;

module.exports = async (req, res) => {
  // --- CORS / Cache ---
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "POST,OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type,Authorization");
  res.setHeader("Cache-Control", "no-store");

  // Preflight
  if (req.method === "OPTIONS") { res.status(200).end(); return; }

  if (req.method !== "POST") {
    res.status(405).json({ error: "Method Not Allowed" });
    return;
  }

  try {
    // --- Body robust einlesen ---
    let payload = req.body;
    if (typeof payload === "string") {
      try { payload = JSON.parse(payload); } catch { payload = {}; }
    }
    if (!payload || typeof payload !== "object") payload = {};

    // --- Hilfsfunktionen ---
    const ensure = (k, v = "") => { if (typeof payload[k] === "undefined") payload[k] = v; };

    // Arrays/Objekte in Strings verwandeln, damit keine "[object Object]" im DOCX landen
    const toText = (v) => {
      if (v == null) return "";
      if (Array.isArray(v)) {
        return v.map(x => {
          if (x == null) return "";
          if (typeof x === "object") {
            // häufige Felder, sonst JSON für Debug
            const t = x.text ?? x.sentence ?? x.value ?? x.title ?? null;
            return t != null ? String(t) : JSON.stringify(x);
          }
          return String(x);
        }).join("\n");
      }
      if (typeof v === "object") {
        const t = v.text ?? v.sentence ?? v.value ?? v.title ?? null;
        return t != null ? String(t) : JSON.stringify(v);
      }
      return String(v);
    };

    // --- Defaults für kritische, optionale Variablen (verhindert ReferenceError bei INS) ---
    ensure("midjourney_article_logo", "");
    ensure("teacher_cloud_logo", "");
    // Doppelte Headline-Varianten im Template absichern
    // (falls nur eine Seite liefert, bleibt die andere nicht undefined)
    if (typeof payload.headline_article === "string" && !payload.headline_artikel) {
      payload.headline_artikel = payload.headline_article;
    }
    if (typeof payload.headline_artikel === "string" && !payload.headline_article) {
      payload.headline_article = payload.headline_artikel;
    }

    // --- Bekannte Textfelder in Strings normalisieren ---
    // Alles zu Text mappen, was typischerweise Rich-Content sein kann:
    const normalizeKeys = (obj) => {
      for (const k of Object.keys(obj)) {
        if (
          /_content$/.test(k) ||
          /_word_box_content$/.test(k) ||
          /^help_link_/.test(k) ||
          k === "article_text" ||
          k === "article_vocab" ||
          k === "themenbereich" ||
          k === "unterthema_des_themenbereichs" ||
          k === "headline_article" ||
          k === "headline_artikel" ||
          k === "source_link"
        ) {
          obj[k] = toText(obj[k]);
        }
      }
    };
    normalizeKeys(payload);

    // --- Template laden (muss via vercel.json includeFiles gebundled sein) ---
    const templatePath = path.join(__dirname, "template.docx");
    const templateBuffer = fs.readFileSync(templatePath);

    // --- DOCX generieren ---
    const docBuffer = await createReport({
      template: templateBuffer,
      data: payload,
      cmdDelimiter: ["{", "}"],
      // Defensive Options:
      rejectNullish: false,  // null/undefined führt nicht zum Abbruch
      errorHandler: (e) => {
        // Fallback: leere Einfügung statt Abbruch – Logik kann bei Bedarf verschärft werden
        return "";
      }
    });

    // --- Korrekte Binary-Antwort ---
    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    );
    res.setHeader("Content-Disposition", 'attachment; filename="generated.docx"');
    res.status(200).send(Buffer.from(docBuffer));
  } catch (err) {
    res.status(500).json({ error: err && err.message ? err.message : String(err) });
  }
};
