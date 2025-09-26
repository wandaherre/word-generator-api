// api/generate.js
const fs = require("fs");
const path = require("path");
const createReport = require("docx-templates").default;

// --- Versions-Check einfügen ---
try {
  const pkg = require('docx-templates/package.json');
  console.log('[docx-templates] version at runtime:', pkg.version);
} catch (e) {
  console.log('[docx-templates] version unknown:', e?.message || e);
}
// --------------------------------

// deine bisherigen Helper-Funktionen hier...

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

/* ---------------- Payload-Normalisierung ---------------- */
function normalizeSourceMeta(payload) {
  if (!payload.source_link_pretty && payload.source_link) {
    payload.source_link_pretty = hostToPretty(payload.source_link);
  }
}

// Entscheidungs-Hilfe: Hat die Übung Gaps (Unterstriche)?
function contentHasGaps(payload, base) {
  const c = (payload[`${base}_content`] ?? payload[`${base}_content_rich`] ?? "");
  const s = safeString(c, "");
  return /_{3,}/.test(s); // mind. drei Unterstriche (___ oder lange Linie)
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
  // *_content_col1/_col2 → Tab-Zeilen zusammenführen
  Object.keys(payload).forEach(k => {
    if (/_content_col1$/i.test(k)) {
      const base = k.replace(/_content_col1$/i, "");
      const left = safeSplitLines(payload[k]).map(s => s.trimEnd());
      const right = safeSplitLines(payload[`${base}_content_col2`]).map(s => s.trimEnd());
      const n = Math.max(left.length, right.length);
      const rows = [];
      for (let i = 0; i < n; i++) {
        rows.push(`${left[i] ?? ""}\t${right[i] ?? ""}`);
      }
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
 * Word-Box-Normalisierung (strenger):
 * - Nur akzeptieren, wenn:
 *   a) die Liste "wie eine Word-Box" aussieht (<=24 entries, jedes < 50 chars)
 *   b) der zugehörige Content Gaps hat (___ / Unterstrichlinien)
 * - Sonst: vollständig verwerfen (kein _word_box_content, kein _line)
 */
function normalizeWordBoxes(payload) {
  Object.keys(payload).forEach(k => {
    if (/_word_box_content$/i.test(k)) {
      const base = k.replace(/_word_box_content$/i, "");
      const arr = toArrayFlexible(payload[k]);

      const looksLikeWordBox = arr.length > 0 && arr.length <= 24 && arr.every(x => x.length > 0 && x.length < 50);
      const gaps = contentHasGaps(payload, base);

      if (looksLikeWordBox && gaps) {
        // hübsche Line bauen, falls nicht vorhanden
        if (!payload[`${base}_word_box_content_line`]) {
          payload[`${base}_word_box_content_line`] = arr.join("   |   ");
        }
        // und *_word_box_content sicher als String halten
        payload[k] = arr.join(" | ");
      } else {
        // Box verwerfen, wenn sie nicht plausibel ist oder Content keine Gaps hat
        delete payload[k];
        delete payload[`${base}_word_box_content_line`];
      }
    }
  });
}

function deriveArticleParagraphRich(payload) {
  // Falls article_text_paragraphX als HTML kam, aber kein _rich existiert → minimal in Text
  for (let i = 1; i <= 16; i++) {
    const key = `article_text_paragraph${i}`;
    if (key in payload && !payload[`${key}_rich`]) {
      let s = safeString(payload[key], "");
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

    console.log('Received payload:', JSON.stringify(payload, null, 2)); // Debug logging

    // Headline Alias (harmlos)
    if (payload.headline_article && !payload.headline_artikel) payload.headline_artikel = payload.headline_article;
    if (payload.headline_artikel && !payload.headline_article) payload.headline_article = payload.headline_artikel;

    // 1) Meta absichern
    if (payload.source_link && !payload.source_link_pretty) {
      payload.source_link_pretty = hostToPretty(payload.source_link);
    }

    // 2) Artikel _rich (nur falls fehlt)
    deriveArticleParagraphRich(payload);

    // 3) Word-Boxes NUR wenn plausibel und Content echte Gaps hat
    normalizeWordBoxes(payload);

    // 4) *_content ⇒ *_content_rich (außer Abitur)
    ensureContentRich(payload);

// 5) 2-Spalten/Pipes → Tabs
normalizeTwoColumnsAndTabs(payload);

// Clean undefined values to prevent split errors
Object.keys(payload).forEach(key => {
  if (payload[key] === undefined) {
    payload[key] = '';
  }
});

// Template laden
const templateBuffer = fs.readFileSync(path.join(__dirname, "template.docx"));

    // DOCX generieren
    const docBuffer = await createReport({
      template: templateBuffer,
      data: payload,
      cmdDelimiter: ["{", "}"],
      processLineBreaksAsNewText: true,
      rejectNullish: false,
      fixSmartQuotes: true,
      additionalJsContext: {
        LINK: (obj) => {
          return { text: obj.label || obj.url, hyperlink: obj.url };
        }
      },
      errorHandler: (err, cmd) => {
        console.error("Template Error at command:", cmd, "Error:", err);
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
