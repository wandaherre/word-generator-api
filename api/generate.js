// api/generate.js
// Kompatibel mit docx-templates@4.14.1
const fs = require("fs");
const path = require("path");
const createReport = require("docx-templates").default;

/* ---------------- Runtime-Version (für F12/Debug) ---------------- */
let DOCX_TEMPLATES_VERSION = "unknown";
try {
  DOCX_TEMPLATES_VERSION = require("docx-templates/package.json").version;
} catch (e) {
  DOCX_TEMPLATES_VERSION = "unknown:" + (e && e.message ? e.message : String(e));
}

/* ---------------- CORS (stabil & permissiv) ---------------- */
function setCors(req, res) {
  const origin = req.headers.origin || "*";
  res.setHeader("Access-Control-Allow-Origin", origin);
  res.setHeader("Vary", "Origin");
  res.setHeader("Access-Control-Allow-Methods", "POST,OPTIONS,GET");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");
  res.setHeader("Access-Control-Max-Age", "86400");
  res.setHeader("Cache-Control", "no-store");
  // Debug-Header für F12/Network
  res.setHeader("x-docx-templates-version", DOCX_TEMPLATES_VERSION);
}

/* ---------------- Helpers (defensiv & minimal-invasiv) ---------------- */
function safeString(v, fallback = "") {
  if (v == null) return fallback;
  if (typeof v === "string") return v;
  if (typeof v === "number" || typeof v === "boolean") return String(v);
  if (Array.isArray(v)) return v.map(x => String(x ?? "")).join("\n");
  return fallback;
}
function safeLines(v) {
  const s = safeString(v, "");
  return s ? s.split(/\r?\n/) : [];
}
function hostToPretty(url) {
  try {
    const u = new URL(String(url));
    return u.hostname.replace(/^www\./, "");
  } catch {
    return safeString(url, "source");
  }
}

/**
 * Baut die Link-Objekte, damit das Template keine JS-Objektliterale mehr auswerten muss.
 * Verwendung im Template:
 *   {LINK source_link_obj}
 *   {?help_link_b1_1a_obj}{LINK help_link_b1_1a_obj}{/help_link_b1_1a_obj}
 *   (B2 analog)
 */
function buildLinkObjects(payload) {
  // Source
  const srcUrl = safeString(payload.source_link, "");
  const srcLabel = safeString(
    payload.source_link_pretty || (srcUrl ? hostToPretty(srcUrl) : ""),
    srcUrl || "source"
  );
  if (srcUrl) {
    payload.source_link_obj = { url: srcUrl, label: srcLabel };
  }

  // Help-Links (B1/B2 … a..d)
  const prefixes = ["help_link_b1_1", "help_link_b2_1"];
  const suffixes = ["a", "b", "c", "d"];
  for (const pref of prefixes) {
    for (const suf of suffixes) {
      const key = `${pref}${suf}`;            // z.B. help_link_b1_1a
      const prettyKey = `${key}_pretty`;      // z.B. help_link_b1_1a_pretty
      const url = safeString(payload[key], "");
      if (!url) continue;
      const label = safeString(payload[prettyKey], "help");
      payload[`${key}_obj`] = { url, label };
    }
  }
}

/**
 * *_content_col1/_col2 → Tab-getrennte Zeilen zusammenführen
 * (nicht aggressiv; nur wenn beide Felder existieren)
 */
function mergeTwoColumns(payload) {
  Object.keys(payload).forEach(k => {
    if (/_content_col1$/i.test(k)) {
      const base = k.replace(/_content_col1$/i, "");
      const col1 = safeLines(payload[k]);
      const col2 = safeLines(payload[`${base}_content_col2`]);
      if (!col1.length && !col2.length) return;

      const n = Math.max(col1.length, col2.length);
      const rows = [];
      for (let i = 0; i < n; i++) {
        const left = safeString(col1[i], "");
        const right = safeString(col2[i], "");
        rows.push(`${left}\t${right}`); // Tabs → echte Spalten in Word
      }
      const combined = rows.join("\n");
      if (!payload[`${base}_content`]) payload[`${base}_content`] = combined;
      if (!payload[`${base}_content_rich`]) payload[`${base}_content_rich`] = combined;
    }
  });
}

/**
 * *_content → *_content_rich spiegeln (nur wenn _rich fehlt; Abitur ausnehmen, falls dort gesondert befüllt wird)
 * minimal-invasiv, um bestehendes Verhalten nicht zu ändern.
 */
function mirrorContentRich(payload) {
  Object.keys(payload).forEach(k => {
    if (/_content$/i.test(k) && !/_word_box_content$/i.test(k)) {
      const base = k.replace(/_content$/i, "");
      if (/^abitur_/i.test(base)) return;
      const richKey = `${base}_content_rich`;
      if (!payload[richKey]) payload[richKey] = safeString(payload[k], "");
    }
  });
}

/**
 * Artikel-Absätze: falls _rich fehlt, aus HTML minimal in Plaintext umwandeln.
 * (nur Ersatz-Fallback; ändert vorhandene Werte nicht)
 */
function ensureArticleParagraphsRich(payload) {
  for (let i = 1; i <= 16; i++) {
    const key = `article_text_paragraph${i}`;
    if (key in payload && !payload[`${key}_rich`]) {
      let s = safeString(payload[key], "");
      s = s.replace(/<\/p>\s*/gi, "\n\n")
           .replace(/<br\s*\/?>/gi, "\n")
           .replace(/<[^>]+>/g, "")
           .replace(/\r\n?/g, "\n")
           .replace(/\n{3,}/g, "\n\n")
           .trim();
      payload[`${key}_rich`] = s;
    }
  }
}

/* ---------------- HTTP Handler ---------------- */
module.exports = async (req, res) => {
  // Debug: GET ?debug=version → zeigt Runtime-Version (Header & Body)
  if (req.method === "GET" && (req.query?.debug === "version" || req.query?.debug === "1")) {
    setCors(req, res);
    res.setHeader("Content-Type", "application/json");
    res.status(200).send(JSON.stringify({ docxTemplatesVersion: DOCX_TEMPLATES_VERSION }));
    return;
  }

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
    // Payload defensiv einlesen
    let payload = req.body;
    if (typeof payload === "string") {
      try { payload = JSON.parse(payload); } catch { payload = {}; }
    }
    if (!payload || typeof payload !== "object") payload = {};

    // Headline-Alias (harmlos, wie gehabt)
    if (payload.headline_article && !payload.headline_artikel) payload.headline_artikel = payload.headline_article;
    if (payload.headline_artikel && !payload.headline_article) payload.headline_article = payload.headline_artikel;

    // Metadaten normalisieren
    if (payload.source_link && !payload.source_link_pretty) {
      payload.source_link_pretty = hostToPretty(payload.source_link);
    }

    // Link-Objekte bereitstellen → Template nutzt {LINK source_link_obj} / {LINK help_link_*_obj}
    buildLinkObjects(payload);

    // Sanfte Fallbacks (keine aggressiven Änderungen)
    ensureArticleParagraphsRich(payload);
    mirrorContentRich(payload);
    mergeTwoColumns(payload);

    // Template laden (robust: zwei Pfade)
    const tryPaths = [
      path.join(__dirname, "template.docx"),
      path.join(process.cwd(), "api", "template.docx"),
    ];
    let templateBuffer = null;
    for (const p of tryPaths) {
      try { templateBuffer = fs.readFileSync(p); break; } catch {}
    }
    if (!templateBuffer) {
      console.error("Template not found. Tried:", tryPaths);
      throw new Error("TEMPLATE_MISSING: Place template.docx at /api/template.docx");
    }

    // DOCX generieren
    const docBuffer = await createReport({
      template: templateBuffer,
      data: payload,
      // Achtung: 4.14.1 + Objektliterale im Template kollidieren mit { }
      // Daher: KEINE Objektliterale im Template verwenden, sondern *_obj-Variablen (s. buildLinkObjects).
      cmdDelimiter: ["{", "}"],
      processLineBreaksAsNewText: true,
      rejectNullish: false,
      fixSmartQuotes: true,
      errorHandler: (err) => { console.log("Template Error:", err); return ""; },
    });

    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
    res.setHeader("Content-Disposition", 'attachment; filename="generated.docx"');
    res.status(200).send(Buffer.from(docBuffer));
  } catch (err) {
    console.error("Generate Error:", err);
    res.status(500).json({
      error: err?.message || String(err),
      docxTemplatesVersion: DOCX_TEMPLATES_VERSION
    });
  }
};
