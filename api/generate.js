// api/generate.js
// Vercel Serverless (CommonJS) + docx-templates: liefert ein echtes DOCX (Binary), nicht JSON.

const fs = require("fs");
const path = require("path");
// WICHTIG: Default-Export korrekt holen
const createReport = require("docx-templates").default;

function setCors(req, res) {
  const origin  = req.headers["origin"] || "*";
  const reqMeth = req.headers["access-control-request-method"] || "POST,OPTIONS";
  const reqHead = req.headers["access-control-request-headers"] || "Content-Type,Authorization";
  res.setHeader("Access-Control-Allow-Origin", origin);
  res.setHeader("Vary", "Origin");
  res.setHeader("Access-Control-Allow-Methods", reqMeth);
  res.setHeader("Access-Control-Allow-Headers", reqHead);
  res.setHeader("Access-Control-Max-Age", "86400");
  res.setHeader("Cache-Control", "no-store");
}

module.exports = async (req, res) => {
  setCors(req, res);

  // Preflight beantworten
  if (req.method === "OPTIONS") { res.status(200).end(); return; }

  if (req.method !== "POST") {
    res.status(405).json({ error: "Method Not Allowed" });
    return;
  }

  try {
    // Body robust lesen
    let payload = req.body;
    if (typeof payload === "string") {
      try { payload = JSON.parse(payload); } catch { payload = {}; }
    }
    if (!payload || typeof payload !== "object") payload = {};

    // Minimal-Defaults gegen Undefined-Fehler in Template-Ausdrücken
    const ensure = (k, v = "") => { if (typeof payload[k] === "undefined") payload[k] = v; };
    ensure("midjourney_article_logo", "");
    ensure("teacher_cloud_logo", "");
    // Headline-Varianten spiegeln, falls nur eine kommt
    if (payload.headline_article && !payload.headline_artikel) payload.headline_artikel = payload.headline_article;
    if (payload.headline_artikel && !payload.headline_article) payload.headline_article = payload.headline_artikel;

    // Template laden (liegt dank vercel.json includeFiles nebenan)
    const templatePath = path.join(__dirname, "template.docx");
    const templateBuffer = fs.readFileSync(templatePath);

    // DOCX erzeugen (docx-templates nutzt SINGLE BRACES {key})
    const docBuffer = await createReport({
      template: templateBuffer,
      data: payload,
      cmdDelimiter: ["{", "}"],
      rejectNullish: false,     // null/undefined → leer statt Abbruch
      errorHandler: () => ""    // falls ein Ausdruck crasht, leer einsetzen statt 500
    });
