// api/generate.js
const fs = require("fs");
const path = require("path");
const createReport = require("docx-templates").default;

/* ---------------- CORS: stabil & permissiv ---------------- */
function setCors(req, res) {
  const origin = req.headers.origin || "*";
  res.setHeader("Access-Control-Allow-Origin", origin);
  res.setHeader("Vary", "Origin");
  // KEINE Credentials in Kombi mit "*"
  res.setHeader("Access-Control-Allow-Methods", "POST,OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");
  res.setHeader("Access-Control-Max-Age", "86400");
  res.setHeader("Cache-Control", "no-store");
}

/* ---------------- Safe Helpers ---------------- */
function safeString(v, fallback = "") {
  if (v == null) return fallback;
  if (typeof v === "string") return v;
  if (Array.isArray(v)) return v.map(x => String(x ?? "")).join("\n");
  if (typeof v === "number" || typeof v === "boolean") return String(v);
  return fallback; // Objekte ignorieren
}
function safeSplitLines(v) {
  const s = safeString(v, "");
  return s ? s.split(/\r?\n/) : [];
}
function toArrayFlexible(v) {
  const s = safeString(v, "").trim();
  if (!s) return [];
  if (s.includes("|")) return s.split("|").map(t => t.trim()).filter(Boolean);
  if (s.includes("\n")) return s.split("\n").map(t => t.trim()).filter(Boolean);
  return s.split(",").map(t => t.trim()).filter(Boolean);
}
function pipesToTabs(text) {
  const s = safeString(text, "");
  return s.replace(/\s*\|\s*/g, "\t");
}
function hostToPretty(url) {
  try {
    const u = new URL(String(url));
    return u.hostname.replace(/^www\./, "");
  } catch {
    return String(url || "source");
  }
}

/* ---------------- Link-HTML für Template {HTML ...} ---------------- */
function escHtml(s) {
  return String(s ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}
function mkAnchor(label, url) {
  const L = escHtml(label || url || "link");
  const U = escHtml(url || "");
  if (!U) return L; // falls URL fehlt, nur Label ausgeben
  return `<a href="${U}">${L}</a>`;
}

/* ---------------- Payload-Normalisierung ---------------- */
function normalizeSourceMeta(payload) {
  if (!payload.source_link_pretty && payload.source_link) {
    payload.source_link_pretty = hostToPretty(payload.source_link);
  }
}

// Hat der Content echte Gaps?
function contentHasGaps(payload, base) {
  const c = (payload[`${base}_content`] ?? payload[`${base}_content_rich`] ?? "");
  const s = safeString(c, "");
  return /_{3,}/.test(s); // mind. drei Unterstriche
}

function ensureContentRich(payload) {
  Object.keys(payload).forEach(k => {
    if (/_content$/i.test(k) && !/_word_box_content$/i.test(k)) {
      const base = k.replace(/_content$/i, "");
      if (/^abitur_/i.test(base)) return;
      const richKey = `${base}_content_rich`;
      if (!payload[richKey]) payload[richKey] = safeString(payload[k], "");
    }
  });
}

function normalizeTwoColumnsAndTabs(payload) {
  // *_content_col1/_col2 → Tab-Zeilen
  Object.keys(payload).forEach(k => {
    if (/_content_col1$/i.test(k)) {
      const base = k.replace(/_content_col1$/i, "");
      const left = safeSplitLines(payload[k]).map(s => s.trimEnd());
      const right = safeSplitLines(payload[`${base}_content_col2`]).map(s => s.trimEnd());
      const n = Math.max(left.length, right.length);
      const rows = [];
      for (let i = 0; i < n; i++) rows.push(`${left[i] ?? ""}\t${right[i] ?? ""}`);
      const combined = rows.join("\n");
      if (!payload[`${base}_content`]) payload[`${base}_content`] = combined;
      if (!payload[`${base}_content_rich`]) payload[`${base}_content_rich`] = combined;
    }
  });
  // Pipes in *_content_rich → Tabs
  Object.keys(payload).forEach(k => {
    if (/_content_rich$/i.test(k)) {
      payload[k] = pipesToTabs(payload[k]);
    }
  });
}

/**
 * Word-Boxen:
 * - Nur akzeptieren, wenn Liste plausibel (<=24 Elemente, jedes <50 Zeichen)
 * - UND der zugehörige Content Gaps enthält
 * - Sonst verwerfen (kein _word_box_content, kein _line)
 */
function normalizeWordBoxes(payload) {
  Object.keys(payload).forEach(k => {
    if (/_word_box_content$/i.test(k)) {
      const base = k.replace(/_word_box_content$/i, "");
      const arr = toArrayFlexible(payload[k]);
      const looksLikeWordBox = arr.length > 0 && arr.length <= 24 && arr.every(x => x.length > 0 && x.length < 50);
      const gaps = contentHasGaps(payload, base);

      if (looksLikeWordBox && gaps) {
        if (!payload[`${base}_word_box_content_line`]) {
          payload[`${base}_word_box_content_line`] = arr.join("   |   ");
        }
        payload[k] = arr.join(" | "); // immer String speichern
      } else {
        delete payload[k];
        delete payload[`${base}_word_box_content_line`];
      }
    }
  });
}

function deriveArticleParagraphRich(payload) {
  // article_text_paragraphX (HTML) → minimaler Plaintext für _rich, falls fehlt
  for (let i = 1; i <= 16; i++) {
    const key = `article_text_paragraph${i}`;
    if (key in payload && !payload[`${key}_rich`]) {
      let s = safeString(payload[key], "");
      s = s.replace(/<\/p>\s*/gi, "\n\n");
      s = s.replace(/<br\s*\/?>/gi, "\n");
      s = s.replace(/<[^>]+>/g, "");
      s = s.replace(/\r\n?/g, "\n").replace(/\n{3,}/g, "\n\n").trim();
      payload[`${key}_rich`] = s;
    }
  }
}

/* ----------- HTML-Link-Felder bereitstellen (für {HTML ...} im Template) ----------- */
function buildHtmlLinkFields(payload) {
  // Source
  if (payload.source_link) {
    const label = payload.source_link_pretty || payload.source_link;
    payload.source_link_html = mkAnchor(label, payload.source_link);
  }
  // Help-Links (B1/B2 – a..d)
  const prefixes = ["help_link_b1_1", "help_link_b2_1"];
  const suffixes = ["a", "b", "c", "d"];
  for (const pref of prefixes) {
    for (const suf of suffixes) {
      const key = `${pref}${suf}`;
      if (payload[key]) {
        const lbl = payload[`${key}_pretty`] || "help";
        payload[`${key}_html`] = mkAnchor(lbl, payload[key]);
      }
    }
  }
}

/* ---------------- HTTP Handler ---------------- */
module.exports = async (req, res) => {
  setCors(req, res);

  if (req.method === "OPTIONS") { res.status(204).end(); return; }
  if (req.method !== "POST") { res.status(405).json({ error: "Method Not Allowed" }); return; }

  try {
    let payload = req.body;
    if (typeof payload === "string") { try { payload = JSON.parse(payload); } catch { payload = {}; } }
    if (!payload || typeof payload !== "object") payload = {};

    // Headline Alias (harmlos)
    if (payload.headline_article && !payload.headline_artikel) payload.headline_artikel = payload.headline_article;
    if (payload.headline_artikel && !payload.headline_article) payload.headline_article = payload.headline_artikel;

    // 1) Meta
    normalizeSourceMeta(payload);

    // 2) Artikel _rich (nur falls fehlt)
    deriveArticleParagraphRich(payload);

    // 3) Word-Boxes NUR wenn plausibel und Content Gaps hat
    normalizeWordBoxes(payload);

    // 4) *_content ⇒ *_content_rich (außer Abitur)
    ensureContentRich(payload);

    // 5) 2-Spalten / Pipes → Tabs
    normalizeTwoColumnsAndTabs(payload);

    // 6) Link-HTML Felder bereitstellen (für Template {HTML ...})
    buildHtmlLinkFields(payload);

    // Template laden
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
      cmdDelimiter: ["{", "}"],
      processLineBreaksAsNewText: true,
      rejectNullish: false,
      errorHandler: (err) => { console.log("Template Error:", err); return ""; },
    });

    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
    res.setHeader("Content-Disposition", 'attachment; filename="generated.docx"');
    res.status(200).send(Buffer.from(docBuffer));
  } catch (err) {
    console.error("Generate Error:", err);
    res.status(500).json({ error: err?.message || String(err) });
  }
};
