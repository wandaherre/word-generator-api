// api/generate.js
const fs = require("fs");
const path = require("path");
const createReport = require("docx-templates").default;

module.exports = async (req, res) => {
  // --- CORS ---
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "POST,OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type,Authorization");
  res.setHeader("Cache-Control", "no-store");

  // Preflight
  if (req.method === "OPTIONS") {
    res.status(200).end();
    return;
  }

  // Nur POST zulassen
  if (req.method !== "POST") {
    res.status(405).json({ error: "Method Not Allowed" });
    return;
  }

  try {
    // Body robust lesen (AI Studio sendet application/json)
    let payload = req.body;
    if (typeof payload === "string") {
      try { payload = JSON.parse(payload); } catch { payload = {}; }
    }
    if (!payload || typeof payload !== "object") payload = {};

    // ---- Defaults für optionale Variablen (verhindert ReferenceError) ----
    const ensure = (k, v = "") => { if (typeof payload[k] === "undefined") payload[k] = v; };
    ensure("midjourney_article_logo", "");
    ensure("teacher_cloud_logo", "");

    // Headline-Sprachvarianten angleichen
    if (typeof payload.headline_article === "string" && !payload.headline_artikel) {
      payload.headline_artikel = payload.headline_article;
    }
    if (typeof payload.headline_artikel === "string" && !payload.headline_article) {
      payload.headline_article = payload.headline_artikel;
    }

    // ---- Arrays/Objekte zu String normalisieren (verhindert "[object Object]") ----
    const toText = (v) => {
      if (v == null) return "";
      if (Array.isArray(v)) {
        return v.map(x => {
          if (x == null) return "";
          if (typeof x === "object") {
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

    // ---- Template laden (liegt dank includeFiles garantiert daneben) ----
    const templatePath = path.join(__dirname, "template.docx");
    const templateBuffer = fs.readFileSync(templatePath);

    // ---- DOCX generieren (docx-templates nutzt SINGLE BRACES {key}) ----
    const docBuffer = await createReport({
      template: templateBuffer,
      data: payload,
      cmdDelimiter: ["{", "}"],
      rejectNullish: false, // null/undefined nicht als Fehler werten
      errorHandler: () => "" // im Zweifel leeren String einfügen statt Abbruch
    });

    // ---- Nur den DOCX-Binärstream senden (korrekter MIME) ----
    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    );
    res.setHeader("Content-Disposition", 'attachment; filename="generated.docx"');
    res.status(200).send(Buffer.from(docBuffer));
  } catch (err) {
    // Nur hier JSON – in allen Erfolgsfällen nie JSON verwenden
    res.status(500).json({ error: err && err.message ? err.message : String(err) });
  }
};
