// api/generate.js
// Vercel Serverless (CommonJS) + docx-templates, mit harter CORS-Preflight-Behandlung
// und Payload-Härtung gegen "item.options is undefined" u. "[object Object]" im Output.

const fs = require("fs");
const path = require("path");
const createReport = require("docx-templates").default;

// ---- Payload-Härtung: stellt sicher, dass überall, wo "options" erwartet werden könnte,
//      mindestens ein leeres Array vorhanden ist und komplexe Werte zu String werden. ----
function hardenPayloadDeep(input) {
  const seen = new WeakSet();

  const toText = (v) => {
    if (v == null) return "";
    if (Array.isArray(v)) {
      return v.map(x => {
        if (x == null) return "";
        if (typeof x === "object") {
          const t = x.text ?? x.sentence ?? x.value ?? x.title ?? x.label ?? null;
          return t != null ? String(t) : JSON.stringify(x);
        }
        return String(x);
      }).join("\n");
    }
    if (typeof v === "object") {
      const t = v.text ?? v.sentence ?? v.value ?? v.title ?? v.label ?? null;
      return t != null ? String(t) : JSON.stringify(v);
    }
    return String(v);
  };

  const ensureOptionsArray = (obj) => {
    // Falls ein Objekt "options" hat/erwartet → sicherstellen, dass es ein Array ist
    if (!("options" in obj) || obj.options == null) obj.options = [];
    if (!Array.isArray(obj.options)) obj.options = [obj.options].filter(Boolean);
  };

  const normalizeObject = (obj) => {
    if (!obj || typeof obj !== "object") return obj;
    if (seen.has(obj)) return obj;
    seen.add(obj);

    // Heuristik: Viele Strukturen heißen item/question/choice/… → dort "options" absichern
    const likelyItem = /item|question|exercise|choice|option|entry|row|line|block/i;

    // 1) Direkte Ebene prüfen
    if (Object.keys(obj).some(k => likelyItem.test(k))) {
      ensureOptionsArray(obj); // global auf obj
    }
    // Immer dann absichern, wenn ein Feld "options" explizit existiert oder existieren könnte
    if ("options" in obj) ensureOptionsArray(obj);

    // 2) Tiefen-Iteration
    for (const k of Object.keys(obj)) {
      const val = obj[k];

      // Felder, die typischerweise Text wollen → in String transformieren
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
        obj[k] = toText(val);
        continue;
      }

      // Arrays tief normalisieren
      if (Array.isArray(val)) {
        obj[k] = val.map(el => {
          if (el && typeof el === "object") {
            // Für jedes Array-Element "options" absichern
            ensureOptionsArray(el);
            return normalizeObject(el);
          }
          return el;
        });
        continue;
      }

      // Objekte tief normalisieren
      if (val && typeof val === "object") {
        // Wenn das Feld nach "item/choice/…" klingt → options absichern
        if (likelyItem.test(k)) ensureOptionsArray(val);
        obj[k] = normalizeObject(val);
      }
    }

    return obj;
  };

  // Start: sichere Kopie erzeugen
  const payload = JSON.parse(JSON.stringify(input ?? {}));

  // Kritische Defaults (verhindert ReferenceError bei INS)
  if (typeof payload.midjourney_article_logo === "undefined") payload.midjourney_article_logo = "";
  if (typeof payload.teacher_cloud_logo === "undefined") payload.teacher_cloud_logo = "";

  // Headline-Spiegelung (falls nur eine Variante kommt)
  if (typeof payload.headline_article === "string" && !payload.headline_artikel) {
    payload.headline_artikel = payload.headline_article;
  }
  if (typeof payload.headline_artikel === "string" && !payload.headline_article) {
    payload.headline_article = payload.headline_artikel;
  }

  return normalizeObject(payload);
}

// ---- CORS-Header setzen; bei OPTIONS die gewünschten Header/Methoden spiegeln ----
function setCors(req, res) {
  const reqOrigin  = req.headers["origin"] || "*";
  const reqMethods = req.headers["access-control-request-method"] || "POST,OPTIONS";
  const reqHeaders = req.headers["access-control-request-headers"] || "*";

  res.setHeader("Access-Control-Allow-Origin", reqOrigin);
  res.setHeader("Vary", "Origin");
  res.setHeader("Access-Control-Allow-Methods", reqMethods);
  res.setHeader("Access-Control-Allow-Headers", reqHeaders);
  res.setHeader("Access-Control-Max-Age", "86400");
  res.setHeader("Cache-Control", "no-store");
}

module.exports = async (req, res) => {
  // CORS immer setzen – auch bei Fehlern
  setCors(req, res);

  // Preflight
  if (req.method === "OPTIONS") {
    res.status(200).end();
    return;
  }

  if (req.method !== "POST") {
    res.status(405).json({ error: "Method Not Allowed" });
    return;
  }

  try {
    // Body robust parsen
    let payload = req.body;
    if (typeof payload === "string") {
      try { payload = JSON.parse(payload); } catch { payload = {}; }
    }
    if (!payload || typeof payload !== "object") payload = {};

    // <<< zentrale Härtung (u. a. options: []) >>>
    payload = hardenPayloadDeep(payload);

    // Template laden (liegt dank includeFiles nebenan)
    const templatePath = path.join(__dirname, "template.docx");
    const templateBuffer = fs.readFileSync(templatePath);

    // DOCX generieren (Single-Braces {key}); Fehler in Ausdrücken abfedern
    const docBuffer = await createReport({
      template: templateBuffer,
      data: payload,
      cmdDelimiter: ["{", "}"],
      rejectNullish: false,
      errorHandler: () => "" // letztes Netz: statt 500 leere Einfügung
    });

    // Binary senden
    res.setHeader(
      "Content-Type",
      "applic
