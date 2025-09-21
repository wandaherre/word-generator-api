// api/generate.js
// Library: docx-templates (guigrpa) – Raw XML via literalXmlDelimiter (||)
// - liefert *_rich als STRING mit "||<w:r>…</w:r>||" (kein Objekt!)
// - *_content_rich / article_text_paragraphX_rich: Runs + <w:br/> für neue Zeilen
// - Wortboxen: *_word_box_content_line mit exakt "   |   "
// - Help-Links: *_pretty = "help (URL)"
// - CORS dynamisch, Preflight sauber

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

/* ------------- Helpers: Markdown → Word <w:r> ------------- */
// Escape für <w:t> (mit Whitespace-Preservierung)
function escText(t) {
  return String(t)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}
// Split in Runs mit **bold** und *italic*
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
// Erzeuge einen <w:r> mit optional <w:b/>, <w:i/>
function runXml({ t, b, i }) {
  const pr = (b || i)
    ? `<w:rPr>${b ? "<w:b/>" : ""}${i ? "<w:i/>" : ""}</w:rPr>`
    : "";
  // xml:space="preserve" hält führende Leerzeichen
  return `<w:r>${pr}<w:t xml:space="preserve">${escText(t)}</w:t></w:r>`;
}
// Baue eine Zeile (optional mit "• " vorn), gebe nur <w:r>… zurück
function lineToRunsXml(line, bullet) {
  const L = String(line || "");
  const text = bullet ? (L.replace(/^\s*[-•]\s+/, "• ")) : L;
  return splitRuns(text).map(runXml).join("");
}
// Mehrere Zeilen → Runs + <w:br/> zwischen den Zeilen (alles im selben Absatz)
function linesToParagraphRunsXml(lines, bulletize) {
  const parts = [];
  lines.forEach((ln, idx) => {
    parts.push(lineToRunsXml(ln, bulletize));
    if (idx < lines.length - 1) parts.push("<w:br/>");
  });
  return parts.join("");
}
// Erzeuge Literal-XML-STRING für docx-templates (mit || … ||)
function toLiteralXmlString(runsXml) {
  // Wichtig: kein Objekt zurückgeben – docx-templates erwartet Text!
  return `||${runsXml}||`;
}

/* --------- Ableitungen ohne Originale zu überschreiben --------- */
const MAX_P = 16;

function deriveArticleAndVocab(payload) {
  for (let i = 1; i <= MAX_P; i++) {
    const kTxt = `article_text_paragraph${i}`;
    if (kTxt in payload) {
      const raw = String(payload[kTxt] ?? "");
      const lines = raw
        .replace(/<br\s*\/?>/gi, "\n")
        .replace(/<\/p>\s*/gi, "\n\n")
        .replace(/<p[^>]*>/gi, "")
        .replace(/<\/?(ul|ol)[^>]*>/gi, "")
        .replace(/<li[^>]*>\s*/gi, "• ").replace(/<\/li>/gi, "\n")
        .replace(/<[^>]+>/g, "")
        .replace(/\r\n?/g, "\n")
        .split(/\n{2,}/)
        .flatMap(p => p.split("\n"));
      payload[`${kTxt}_rich`] = toLiteralXmlString(
        linesToParagraphRunsXml(lines, false)
      );
    }
    // Vokabel: pX_1..3 → *_line plus *_rich (Zeilen)
    const w1 = (payload[`article_vocab_p${i}_1`] ?? "").toString().trim();
    const w2 = (payload[`article_vocab_p${i}_2`] ?? "").toString().trim();
    const w3 = (payload[`article_vocab_p${i}_3`] ?? "").toString().trim();
    const words = [w1, w2, w3].filter(Boolean);
    if (words.length) {
      payload[`article_vocab_p${i}_line`] = words.join("   |   "); // exakt 3 Spaces
      payload[`article_vocab_p${i}_rich`] = toLiteralXmlString(
        linesToParagraphRunsXml(words, false)
      );
    }
  }
}

function deriveExercises(payload) {
  for (const k of Object.keys(payload)) {
    // Wortboxen → *_line + *_rich (Zeilenliste)
    if (/^.*_word_box_content$/i.test(k)) {
      const raw = (payload[k] ?? "").toString().trim();
      if (raw) {
        let items;
        if (raw.includes("|")) items = raw.split("|");
        else if (raw.includes("\n")) items = raw.split("\n");
        else items = raw.split(",");
        items = items.map(s => s.trim()).filter(Boolean);
        payload[`${k}_line`] = items.join("   |   "); // exakt 3 Spaces
        payload[`${k}_rich`] = toLiteralXmlString(
          linesToParagraphRunsXml(items, false)
        );
      }
    }
    // Übungsinhalte → *_content_rich (Runs + <w:br/>)
    if (/_content$/i.test(k)) {
      const val = String(payload[k] ?? "");
      // Bullet-Heuristik: Wenn die Zeile mit "- " oder "• " beginnt, Bullet anzeigen
      const lines = val.replace(/\r\n?/g, "\n").split("\n");
      const bulletize = lines.some(s => /^\s*[-•]\s+/.test(s));
      payload[`${k}_rich`] = toLiteralXmlString(
        linesToParagraphRunsXml(lines, bulletize)
      );
    }
    // Help-Links → *_pretty
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

    // harmlose Defaults (keine Überschreibung echter Werte)
    if (typeof payload.midjourney_article_logo === "undefined") payload.midjourney_article_logo = "";
    if (typeof payload.teacher_cloud_logo === "undefined") payload.teacher_cloud_logo = "";
    if (payload.headline_article && !payload.headline_artikel) payload.headline_artikel = payload.headline_article;
    if (payload.headline_artikel && !payload.headline_article) payload.headline_article = payload.headline_artikel;

    // Ableitungen erzeugen
    deriveArticleAndVocab(payload);
    deriveExercises(payload);

    // Template rendern
    const templateBuffer = fs.readFileSync(path.join(__dirname, "template.docx"));
    const docBuffer = await createReport({
      template: templateBuffer,
      data: payload,
      // wichtig: wir nutzen hier Standard '{','}' – deine .docx hat diese bereits
      cmdDelimiter: ["{", "}"],
      // Zeilenumbrüche robuster als <w:br/> zu behandeln (optional)
      processLineBreaksAsNewText: true,
      rejectNullish: false,
      errorHandler: () => "" // lieber leer als 500
    });

    // DOCX zurück
    setCors(req, res);
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
    res.setHeader("Content-Disposition", 'attachment; filename="generated.docx"');
    res.status(200).send(Buffer.from(docBuffer));
  } catch (err) {
    setCors(req, res);
    res.status(500).json({ error: err?.message || String(err) });
  }
};
