// api/generate.js
// Vercel Serverless (CommonJS) + docx-templates
// - arbeitet NUR mit Absatz-/Vokabel-Slots (article_text_paragraph1..16, article_vocab_pX_1..3)
// - gibt zusätzliche *_rich Felder (RAW WordprocessingML) aus:
//     * echte Zeilenumbrüche/Absätze
//     * **fett** / *kursiv*
//     * Aufzählungen mit einheitlicher Einrückung (auch Zeilenumbruchzeilen)
// - baut "help (URL)"-Strings, damit keine nackten URLs mehr stehen
// - Vokabel-Linien mit exakt "   |   " (3 Spaces) als Trenner
// - CORS stabil, DOCX wird als Binary mit korrektem MIME gesendet

const fs = require("fs");
const path = require("path");
const createReport = require("docx-templates").default;

/* ---------------- CORS ---------------- */
function setCors(req, res) {
  const origin = req.headers["origin"] || "*";
  res.setHeader("Access-Control-Allow-Origin", origin);
  res.setHeader("Vary", "Origin");
  res.setHeader("Access-Control-Allow-Credentials", "true");
  const reqMethods = req.headers["access-control-request-method"] || "POST,OPTIONS";
  const reqHeaders = req.headers["access-control-request-headers"] || "Content-Type,Authorization";
  res.setHeader("Access-Control-Allow-Methods", reqMethods);
  res.setHeader("Access-Control-Allow-Headers", reqHeaders);
  res.setHeader("Access-Control-Max-Age", "86400");
  res.setHeader("Cache-Control", "no-store");
}

/* -------- Helpers: Plain & Rich -------- */
// 1) HTML/Plain nach Text (Fallback – keine sichtbaren Tags)
function stripHtmlToText(input) {
  if (input == null) return "";
  let s = String(input);
  s = s.replace(/<br\s*\/?>/gi, "\n")
       .replace(/<\/p>\s*/gi, "\n\n")
       .replace(/<p[^>]*>/gi, "");
  s = s.replace(/<li[^>]*>\s*/gi, "- ").replace(/<\/li>/gi, "\n");
  s = s.replace(/<\/?(ul|ol)[^>]*>/gi, "");
  s = s.replace(/<[^>]+>/g, "");
  s = s.replace(/\r\n?/g, "\n").replace(/\n{3,}/g, "\n\n").trim();
  return s;
}

// 2) Mini-Markdown → RAW WordprocessingML
//    – **bold**, *italic*
//    – Absätze (doppelte \n) & Zeilenumbrüche
//    – Aufzählungen ("- " oder "• ") mit Hängeeinzug für alle Zeilen
function mdToRawXml(input) {
  if (input == null) return { _type: "rawXml", xml: "" };
  let s = String(input);

  // HTML → rudimentäres Markdown
  s = s.replace(/<br\s*\/?>/gi, "\n")
       .replace(/<\/p>\s*/gi, "\n\n")
       .replace(/<p[^>]*>/gi, "")
       .replace(/<\/?(ul|ol)[^>]*>/gi, "")
       .replace(/<li[^>]*>\s*/gi, "• ").replace(/<\/li>/gi, "\n");
  s = s.replace(/<\/?strong>/gi, "**").replace(/<\/?b>/gi, "**");
  s = s.replace(/<\/?em>/gi, "*").replace(/<\/?i>/gi, "*");
  s = s.replace(/<[^>]+>/g, "");
  s = s.replace(/\r\n?/g, "\n");

  const paragraphs = s.split(/\n{2,}/).map(p => p.replace(/[ \t]+\n/g, "\n"));
  const esc = (t) => String(t).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;");
  const re = /(\*\*[^*]+\*\*|\*[^*]+\*)/;

  const w = (tag, inner) => `<w:${tag}>${inner}</w:${tag}>`;
  const rPr = (r) => {
    const bits = [];
    if (r.b) bits.push("<w:b/>");
    if (r.i) bits.push("<w:i/>");
    return bits.length ? w("rPr", bits.join("")) : "";
  };

  function runsFor(line) {
    const out = [];
    let rest = line;
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

  // Paragraph (normal)
  const pNormal = (line) => {
    const runs = runsFor(line);
    const rXml = runs.map(r => w("r", rPr(r) + w("t", `<w:noProof/>${esc(r.t)}`))).join("");
    return w("p", rXml);
  };

  // Paragraph (Bullet mit einheitlicher Einrückung – auch bei Umbruchzeilen)
  const pBullet = (line) => {
    const bulletText = line.replace(/^\s*[-•]\s+/, "• ");
    const runs = runsFor(bulletText);
    const rXml = runs.map(r => w("r", rPr(r) + w("t", `<w:noProof/>${esc(r.t)}`))).join("");
    // 720 twips = 0,5", hanging 360 twips (~0,25")
    const pPr = w("pPr", `<w:ind w:left="720" w:hanging="360"/>`);
    return w("p", pPr + rXml);
  };

  let xml = "";
  for (const p of paragraphs) {
    const lines = p.split("\n");
    for (let line of lines) {
      const isBullet = /^\s*[-•]\s+/.test(line);
      xml += isBullet ? pBullet(line) : pNormal(line);
    }
  }
  return { _type: "rawXml", xml };
}

function buildPlainAndRich(val) {
  return { plain: stripHtmlToText(val), rich: mdToRawXml(val) };
}

/* ---------- Artikel-/Vokabel-Slots ---------- */
const MAX_P = 16;
function prepareArticleParagraphs(payload) {
  for (let i = 1; i <= MAX_P; i++) {
    const kTxt = `article_text_paragraph${i}`;
    const txt = payload[kTxt] ?? "";
    const pr = buildPlainAndRich(txt);
    payload[kTxt] = pr.plain;                  // Plain für {article_text_paragraphX}
    payload[`${kTxt}_rich`] = pr.rich;         // RAW für {RAW article_text_paragraphX_rich}

    // Vokabel: drei Slots pX_1..3 → zusätzliche Formate
    const w1 = (payload[`article_vocab_p${i}_1`] ?? "").toString().trim();
    const w2 = (payload[`article_vocab_p${i}_2`] ?? "").toString().trim();
    const w3 = (payload[`article_vocab_p${i}_3`] ?? "").toString().trim();
    const words = [w1, w2, w3].filter(Boolean);

    // exakt "   |   " (3 Spaces) als Trenner
    payload[`article_vocab_p${i}_line`] = words.join("   |   ");

    // RAW-Variante als Zeilenliste (jede Zeile ein Absatz mit Einrückung 0)
    const vocText = words.join("\n");
    payload[`article_vocab_p${i}_rich`] = mdToRawXml(vocText);
  }
}

/* ---------- Übungen & help-Links ---------- */
function normalizeExerciseContentAndHelp(payload) {
  for (const k of Object.keys(payload)) {
    // *_content → Plain + *_content_rich (für RAW in Template)
    if (/_content$/.test(k)) {
      const pr = buildPlainAndRich(payload[k]);
      payload[k] = pr.plain;
      payload[`${k}_rich`] = pr.rich;
    }
    // help_link_* → pretty: "help (URL)"
    if (/^help_link_/i.test(k)) {
      const url = (payload[k] ?? "").toString().trim();
      payload[`${k}_pretty`] = url ? `help (${url})` : "";
      // Optional: kurze sichtbare Variante "help" nur als Text (ohne aktiven Link)
      payload[`${k}_label`] = url ? "help" : "";
    }
  }
}

/* --------------- Handler --------------- */
module.exports = async (req, res) => {
  setCors(req, res);
  if (req.method === "OPTIONS") { res.status(204).end(); return; }
  if (req.method !== "POST") { res.status(405).json({ error: "Method Not Allowed" }); return; }

  try {
    let payload = req.body;
    if (typeof payload === "string") { try { payload = JSON.parse(payload); } catch { payload = {}; } }
    if (!payload || typeof payload !== "object") payload = {};

    // Durchreichen von Meta (z. B. themenbereich im Header) – NICHT verändern
    if (typeof payload.midjourney_article_logo === "undefined") payload.midjourney_article_logo = "";
    if (typeof payload.teacher_cloud_logo === "undefined") payload.teacher_cloud_logo = "";
    // Headline-Spiegelung nur falls nötig
    if (payload.headline_article && !payload.headline_artikel) payload.headline_artikel = payload.headline_article;
    if (payload.headline_artikel && !payload.headline_article) payload.headline_article = payload.headline_artikel;

    // Artikel/Vokabel-Slots vorbereiten
    prepareArticleParagraphs(payload);

    // Übungen & help-Links normalisieren
    normalizeExerciseContentAndHelp(payload);

    // Template laden & rendern
    const templateBuffer = fs.readFileSync(path.join(__dirname, "template.docx"));
    const docBuffer = await createReport({
      template: templateBuffer,
      data: payload,
      cmdDelimiter: ["{", "}"],
      rejectNullish: false,
      errorHandler: () => "" // lieber leer als 500
    });

    // DOCX Binary zurück
    setCors(req, res);
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
    res.setHeader("Content-Disposition", 'attachment; filename="generated.docx"');
    res.status(200).send(Buffer.from(docBuffer));
  } catch (err) {
    setCors(req, res);
    res.status(500).json({ error: err?.message || String(err) });
  }
};
