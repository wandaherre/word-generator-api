// api/generate.js
// Ziel: Übungen mit HTML im Payload robust verarbeiten.
// - HTML in *_content wird serverseitig nach "leichtem Markdown" normalisiert:
//     * Absätze/Zeilen: \n
//     * <ol>/<ul>/<li> → Bullet-Zeilen ("- " wird später optisch "• ")
//     * <strong>/<b> → **bold**, <em>/<i> → *italic*
//     * alle übrigen Tags raus
// - Danach entstehen *_content_rich (Word-Runs + <w:br/> + Hanging Indent)
// - Originalfelder bleiben unberührt (keine Überschreibung)
// - Wortboxen: *_word_box_content_line (mit exakt "   |   ") + *_word_box_content_rich
// - Vokabel je Absatz: *_line + *_rich
// - Help-Links: *_pretty = "help (URL)"
// - CORS: dynamisch, Preflight sauber

const fs = require("fs");
const path = require("path");
const createReport = require("docx-templates").default;

/* ---------------- CORS ---------------- */
function setCors(req, res) {
  const origin = req.headers.origin || "*";
  const reqMethods = req.headers["access-control-request-method"] || "POST,OPTIONS";
  const reqHeaders = req.headers["access-control-request-headers"] || "Content-Type,Authorization";
  res.setHeader("Access-Control-Allow-Origin", origin);
  res.setHeader("Vary", "Origin");
  res.setHeader("Access-Control-Allow-Credentials", "true");
  res.setHeader("Access-Control-Allow-Methods", reqMethods);
  res.setHeader("Access-Control-Allow-Headers", reqHeaders);
  res.setHeader("Access-Control-Max-Age", "86400");
  res.setHeader("Cache-Control", "no-store");
}

/* ---------- HTML → "leichtes Markdown" (Text) ---------- */
function htmlToLightMarkdown(input) {
  if (input == null) return "";
  let s = String(input);

  // Normalisiere häufige HTML-Listen/Textblöcke
  s = s.replace(/<br\s*\/?>/gi, "\n");
  s = s.replace(/<\/p>\s*/gi, "\n\n").replace(/<p[^>]*>/gi, "");

  // Ordered/Unordered Lists -> Zeilen mit "- "
  // zuerst <li>, damit verschachtelte Listen robust werden
  s = s.replace(/<li[^>]*>\s*/gi, "- ").replace(/<\/li>/gi, "\n");
  s = s.replace(/<\/?(ul|ol)[^>]*>/gi, "");

  // Fett/Kursiv in Markdown
  s = s.replace(/<\/?strong>/gi, "**").replace(/<\/?b>/gi, "**");
  s = s.replace(/<\/?em>/gi, "*").replace(/<\/?i>/gi, "*");

  // Restliche Tags komplett entfernen (inkl. Klassen/Spans/Divs/Headers etc.)
  s = s.replace(/<[^>]+>/g, "");

  // Windows-Zeilenenden → \n
  s = s.replace(/\r\n?/g, "\n");

  // Mehrfache leere Zeilen reduzieren
  s = s.replace(/\n{3,}/g, "\n\n").trim();

  return s;
}

/* ---------- Markdown (Zeilen) → Word-Runs als Literal-XML ---------- */
// Escaping für <w:t>
function escText(t) {
  return String(t).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

// Split in Runs (**bold**, *italic*)
function splitRuns(mdLine) {
  const out = [];
  let rest = String(mdLine);
  const re = /(\*\*[^*]+\*\*|\*[^*]+\*)/;
  while (rest.length) {
    const m = rest.match(re);
    if (!m) { out.push({ t: rest }); break; }
    const [match] = m; const i = m.index;
    if (i > 0) out.push({ t: rest.slice(0, i) });
    if (match.startsWith("**")) out.push({ t: match.slice(2, -2), b: true });
    else out.push({ t: match.slice(1, -1), i: true });
    rest = rest.slice(i + match.length);
  }
  return out;
}

function runXml({ t, b, i }) {
  const pr = (b || i) ? `<w:rPr>${b ? "<w:b/>" : ""}${i ? "<w:i/>" : ""}</w:rPr>` : "";
  return `<w:r>${pr}<w:t xml:space="preserve">${escText(t)}</w:t></w:r>`;
}

// Eine Zeile → Runs; wenn sie mit "- " oder "• " beginnt, später als Bullet behandeln
function lineToRunsXml(line) {
  const L = String(line || "");
  return splitRuns(L).map(runXml).join("");
}

// Mehrere Zeilen → Runs + <w:br/> (alles im selben Absatz)
function linesToRunsXml(lines) {
  const parts = [];
  lines.forEach((ln, idx) => {
    parts.push(lineToRunsXml(ln));
    if (idx < lines.length - 1) parts.push("<w:br/>");
  });
  return parts.join("");
}

// Literal-XML für docx-templates (Default-Delimiter "||")
function toLiteralXmlString(runsXml) {
  return `||${runsXml}||`;
}

/* ---------- Ableitungen ohne Originale zu überschreiben ---------- */
const MAX_P = 16;

function deriveArticleAndVocab(payload) {
  for (let i = 1; i <= MAX_P; i++) {
    const kTxt = `article_text_paragraph${i}`;
    if (kTxt in payload) {
      const md = htmlToLightMarkdown(payload[kTxt]);
      const lines = md.split(/\n/);
      payload[`${kTxt}_rich`] = toLiteralXmlString(linesToRunsXml(lines));
    }

    const w1 = (payload[`article_vocab_p${i}_1`] ?? "").toString().trim();
    const w2 = (payload[`article_vocab_p${i}_2`] ?? "").toString().trim();
    const w3 = (payload[`article_vocab_p${i}_3`] ?? "").toString().trim();
    const words = [w1, w2, w3].filter(Boolean);
    if (words.length) {
      payload[`article_vocab_p${i}_line`] = words.join("   |   "); // exakt 3 Spaces
      payload[`article_vocab_p${i}_rich`] = toLiteralXmlString(linesToRunsXml(words));
    }
  }
}

function deriveExercises(payload) {
  for (const k of Object.keys(payload)) {
    // Wortboxen → *_line + *_rich
    if (/^.*_word_box_content$/i.test(k)) {
      const raw = (payload[k] ?? "").toString().trim();
      if (raw) {
        let items;
        if (raw.includes("|")) items = raw.split("|");
        else if (raw.includes("\n")) items = raw.split("\n");
        else items = raw.split(",");
        items = items.map(s => s.trim()).filter(Boolean);
        payload[`${k}_line`] = items.join("   |   ");      // exakt 3 Spaces rund um |
        payload[`${k}_rich`] = toLiteralXmlString(linesToRunsXml(items));
      }
    }

    // *_content → HTML raus + Markdown → Runs + Hanging Indent (über Absatzformat im Template)
    if (/_content$/i.test(k)) {
      const md = htmlToLightMarkdown(payload[k]);
      const lines = md.split(/\n/);

      // Bullet-Heuristik (wir machen optische Bullets "• " aus "- "/ "• ")
      const bulletized = lines.map(line =>
        /^\s*[-•]\s+/.test(line) ? line.replace(/^\s*[-•]\s+/, "• ") : line
      );

      payload[`${k}_rich`] = toLiteralXmlString(linesToRunsXml(bulletized));
    }

    // Help-Links: pretty
    if (/^help_link_/i.test(k)) {
      const url = (payload[k] ?? "").toString().trim();
      payload[`${k}_pretty`] = url ? `help (${url})` : "";
    }
  }
}

/* -------------------- Handler -------------------- */
module.exports = async (req, res) => {
  setCors(req, res);
  if (req.method === "OPTIONS") { res.status(204).end(); return; }
  if (req.method !== "POST") { res.status(405).json({ error: "Method Not Allowed" }); return; }

  try {
    let payload = req.body;
    if (typeof payload === "string") { try { payload = JSON.parse(payload); } catch { payload = {}; } }
    if (!payload || typeof payload !== "object") payload = {};

    // harmlose Defaults
    if (typeof payload.midjourney_article_logo === "undefined") payload.midjourney_article_logo = "";
    if (typeof payload.teacher_cloud_logo === "undefined") payload.teacher_cloud_logo = "";
    if (payload.headline_article && !payload.headline_artikel) payload.headline_artikel = payload.headline_article;
    if (payload.headline_artikel && !payload.headline_article) payload.headline_article = payload.headline_artikel;

    deriveArticleAndVocab(payload);
    deriveExercises(payload);

    const templateBuffer = fs.readFileSync(path.join(__dirname, "template.docx"));
    const docBuffer = await createReport({
      template: templateBuffer,
      data: payload,
      cmdDelimiter: ["{", "}"],
      processLineBreaksAsNewText: true,
      rejectNullish: false,
      errorHandler: () => ""
    });

    setCors(req, res);
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
    res.setHeader("Content-Disposition", 'attachment; filename="generated.docx"');
    res.status(200).send(Buffer.from(docBuffer));
  } catch (err) {
    setCors(req, res);
    res.status(500).json({ error: err?.message || String(err) });
  }
};
