// api/generate.js
// DOCX via docx-templates. Fügt *_rich als Literal-XML (||...||) ein.
// Fixes: Links, Active-Absätze/De-Duplicate, Unterstrich-Normalisierung, Table->Text (Idioms),
// Label-Stripping vor MC a) b) c), keine extra BR am Blockende, Blank-Page-Fix (max. 1 Leerzeile).

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

/* ---------- HTML (inkl. Tabellen) -> „leichtes Markdown“ ---------- */
function normalizeTables(html) {
  if (html == null) return "";
  let s = String(html);
  // Tabellen grob in „Zeilen | Spalten“ überführen
  s = s.replace(/<tr[^>]*>/gi, "");
  s = s.replace(/<\/tr>/gi, "\n");
  s = s.replace(/<t[hd][^>]*>/gi, "");
  s = s.replace(/<\/t[hd]>/gi, " | ");
  // Mehrere Pipes zusammenfassen, führende/trailing Pipen später getrimmt
  s = s.replace(/\s*\|\s*(\|\s*)+/g, " | ");
  return s;
}

function htmlToLightMd(input, { forActive = false } = {}) {
  if (input == null) return "";
  // 1) Tabellen zuerst in Pipes
  let s = normalizeTables(input);

  // 2) Zeilen/Absätze
  s = s.replace(/<br\s*\/?>/gi, "\n");
  s = s.replace(/<\/p>\s*/gi, "\n\n").replace(/<p[^>]*>/gi, "");
  s = s.replace(/<\/div>\s*/gi, "\n\n").replace(/<div[^>]*>/gi, "");

  // 3) Headings als Absatz + fett
  s = s.replace(/<\/h[1-6]>\s*/gi, "\n\n")
       .replace(/<h[1-6][^>]*>/gi, "**")
       .replace(/<\/h[1-6]>/gi, "**");

  // 4) Listen
  s = s.replace(/<li[^>]*>\s*/gi, "- ").replace(/<\/li>/gi, "\n");
  s = s.replace(/<\/?(ul|ol)[^>]*>/gi, "");

  // 5) Bold/Kursiv
  s = s.replace(/<\/?strong>/gi, "**").replace(/<\/?b>/gi, "**");
  s = s.replace(/<\/?em>/gi, "*").replace(/<\/?i>/gi, "*");

  // 6) Restliche Tags raus
  s = s.replace(/<[^>]+>/g, "");

  // 7) Windows-Zeilenenden
  s = s.replace(/\r\n?/g, "\n");

  // 8) Active-Phasen deutlicher absetzen
  if (forActive) {
    s = s.replace(/(?:^|\n)\s*(Phase\s*\d+\s*:)/gi, (m, g1) => `\n\n**${g1.trim()}**\n`);
    // Falls alles in einer Zeile: splitte vor "1. ", "2. ", "• ", "- " usw.
    s = s.replace(/(\s)(\d+\.\s+)/g, "$1\n$2");
  }

  // 9) Pipe-Kosmetik (am Zeilenende)
  s = s.split("\n").map(line => line.replace(/\s*\|\s*$/,"").trimEnd()).join("\n");

  // 10) Mehrfache Leerzeilen grob auf max. 2 reduzieren (Feinbegrenzung kommt später)
  s = s.replace(/\n{3,}/g, "\n\n").trim();

  return s;
}

/* ---------- Markdown Runs -> Literal XML (||...||) ---------- */
function escText(t){return String(t).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;");}
function splitRuns(md){
  const out=[]; let rest=String(md);
  const re=/(\*\*[^*]+\*\*|\*[^*]+\*)/;
  while(rest.length){
    const m=rest.match(re);
    if(!m){out.push({t:rest}); break;}
    const [match]=m; const i=m.index;
    if(i>0) out.push({t:rest.slice(0,i)});
    if(match.startsWith("**")) out.push({t:match.slice(2,-2),b:true});
    else out.push({t:match.slice(1,-1),i:true});
    rest=rest.slice(i+match.length);
  }
  return out;
}
function runXml({t,b,i}) {
  const pr=(b||i)?`<w:rPr>${b?"<w:b/>":""}${i?"<w:i/>":""}</w:rPr>`:"";
  return `<w:r>${pr}<w:t xml:space="preserve">${escText(t)}</w:t></w:r>`;
}

/* ---------- Items & Rendering (kein Extra-BR am Blockende) ---------- */
// items = Array aus { kind:'text', line:'...' } oder { kind:'blank' }
function collapseBlankItems(items){
  const out=[]; let prevBlank=false;
  for (const it of items){
    if (it.kind === "blank"){
      if (!prevBlank) { out.push(it); prevBlank = true; }
    } else {
      out.push(it); prevBlank = false;
    }
  }
  while (out[0] && out[0].kind === "blank") out.shift();
  while (out[out.length-1] && out[out.length-1].kind === "blank") out.pop();
  return out;
}

function itemsToRunsXml(items){
  const collapsed = collapseBlankItems(items);
  const parts=[];
  for (let idx=0; idx<collapsed.length; idx++){
    const it = collapsed[idx];
    const next = collapsed[idx+1];
    if (it.kind === "blank") { parts.push("<w:br/>"); continue; }
    splitRuns(it.line).forEach(r=>parts.push(runXml(r)));
    if (next && next.kind === "text") parts.push("<w:br/>");
  }
  return parts.join("");
}
function toLiteral(runsXml){ return `||${runsXml}||`; }

/* ---------- Helpers: Unterstriche, Labels, Nummerierung ---------- */
function repeatChar(ch,n){return new Array(n+1).join(ch);}

// Normalisiert JEDE Sequenz von >=3 "_" auf gewünschte Länge
function normalizeUnderscores(line, desiredLen){
  return String(line).replace(/_{3,}/g, repeatChar("_", desiredLen));
}
function sentenceUnderline(line){ return String(line).replace(/___SENTENCE___/g, repeatChar("_", 80)); }

// Entfernt führende Labels/Nummern: "A.", "B)", "1.", "a)" etc.
function stripLeadingLabels(line){
  return String(line).replace(/^\s*([A-Za-z]\)|[A-Za-z]\.|[0-9]+\.)\s+/, "");
}

// Nummeriere NUR nicht-leere Zeilen; leere bleiben Leerzeilen
function enumeratePreserveBlanks(lines){
  const items=[]; let n=0;
  for (const raw of lines) {
    const s = String(raw);
    if (s.trim().length === 0) { items.push({ kind: "blank" }); continue; }
    if (/^\s*([0-9]+\.)|[-•]\s+/.test(s)) { items.push({ kind: "text", line: s }); continue; }
    n += 1; items.push({ kind: "text", line: `${n}. ${s}` });
  }
  return items;
}

// MC: a) b) c); entfernt vorhandene führende Label
function choicesABC(lines){
  const labels="abcdefghijklmnopqrstuvwxyz".split("");
  const items=[]; let i=0;
  for (const raw of lines) {
    let s=String(raw);
    if (s.trim().length===0) { items.push({ kind: "blank" }); continue; }
    s = stripLeadingLabels(s);
    const label = labels[i] || String.fromCharCode(97+i);
    items.push({ kind:"text", line: `${label}) ${s.replace(/^\s*[-•]\s+/, "")}` });
    i++;
  }
  return items;
}

/* ---------- Ableitungen ---------- */
const MAX_P=16;

function deriveArticle(payload){
  if (payload.source_link) payload.source_link_pretty = `source (${payload.source_link})`;

  for(let i=1;i<=MAX_P;i++){
    const k=`article_text_paragraph${i}`;
    if(k in payload){
      const md = htmlToLightMd(payload[k]);
      const lines = md.split(/\n/);
      const items = lines.map(s => s.trim().length ? {kind:"text", line:s} : {kind:"blank"});
      payload[`${k}_rich`] = toLiteral(itemsToRunsXml(items));
    }
    const w1=(payload[`article_vocab_p${i}_1`]||"").toString().trim();
    const w2=(payload[`article_vocab_p${i}_2`]||"").toString().trim();
    const w3=(payload[`article_vocab_p${i}_3`]||"").toString().trim();
    const words=[w1,w2,w3].filter(Boolean);
    if(words.length){
      payload[`article_vocab_p${i}_line`] = words.join("   |   ");
      const it = words.map(s => ({kind:"text", line:s}));
      payload[`article_vocab_p${i}_rich`] = toLiteral(itemsToRunsXml(it));
    }
  }
  const paras=[];
  for(let i=1;i<=MAX_P;i++){
    const k=`article_text_paragraph${i}`;
    if(payload[k]) paras.push(htmlToLightMd(payload[k]));
  }
  if(paras.length){
    const lines = paras.flatMap(p=>p.split(/\n{2,}/)).flatMap(p=>p.split("\n"));
    const items = lines.map(s => s.trim().length ? {kind:"text", line:s} : {kind:"blank"});
    payload.article_text_all_rich = toLiteral(itemsToRunsXml(items));
  }
}

function deriveExercises(payload){
  for(const k of Object.keys(payload)){
    // Wortboxen
    if(/_word_box_content$/i.test(k)){
      let raw=(payload[k]||"").toString().trim(); if(!raw) continue;
      let itemsArr;
      if(raw.includes("|")) itemsArr=raw.split("|");
      else if(raw.includes("\n")) itemsArr=raw.split("\n");
      else itemsArr=raw.split(",");
      itemsArr=itemsArr.map(s => s.trim()).filter(Boolean);
      payload[`${k}_line`] = itemsArr.join("   |   ");
      const it = itemsArr.map(s => ({kind:"text", line:s}));
      payload[`${k}_rich`] = toLiteral(itemsToRunsXml(it));
    }

    // Inhalte
    if(/_content$/i.test(k)){
      const isActive = /^active_/i.test(k);
      const isIdioms = /idioms/i.test(k);
      const isGrammar = /b1_|b2_|grammar/i.test(k) && !isIdioms;

      // HTML->Text (Active spezieller)
      let md = htmlToLightMd(payload[k], { forActive: isActive });

      // Heuristik: Explanation = Content? -> bei Active: Explanation nicht duplizieren
      if (isActive) {
        const exKey = k.replace(/_content$/i, "_explanation");
        if (payload[exKey] && String(payload[exKey]).trim() === String(payload[k]).trim()) {
          // Explanation bleibt kurz: nur erster Satz
          const firstSentence = String(payload[exKey]).split(/(?<=[.!?])\s+/)[0] || "";
          payload[exKey] = firstSentence;
        }
      }

      // Zeilen splitten
      let lines = md.split(/\n/).map(s => s.replace(/\s*\|\s*/g, "   |   ").trimEnd());

      // Unterstriche/Linien normalisieren
      lines = lines.map(s => sentenceUnderline(s));
      lines = lines.map(s => normalizeUnderscores(s, isIdioms ? 20 : (isGrammar ? 13 : 13)));

      // MC-Labels oder Nummerierung
      let items;
      if (isIdioms) {
        items = choicesABC(lines);
      } else if (isActive) {
        items = lines.map(s => s.trim().length ? {kind:"text", line:s} : {kind:"blank"});
      } else {
        items = enumeratePreserveBlanks(lines);
      }

      payload[`${k}_rich`] = toLiteral(itemsToRunsXml(items));
    }

    // Help-Links (sichtbarer Text)
    if(/^help_link_/i.test(k)){
      const url=(payload[k]||"").toString().trim();
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
    if (typeof payload === "string"){ try{ payload=JSON.parse(payload);}catch{ payload={}; } }
    if (!payload || typeof payload !== "object") payload = {};

    // Defaults
    if (typeof payload.midjourney_article_logo === "undefined") payload.midjourney_article_logo = "";
    if (typeof payload.teacher_cloud_logo === "undefined") payload.teacher_cloud_logo = "";
    if (payload.headline_article && !payload.headline_artikel) payload.headline_artikel = payload.headline_article;
    if (payload.headline_artikel && !payload.headline_article) payload.headline_article = payload.headline_artikel;
    if (payload.source_link && !payload.source_link_pretty) payload.source_link_pretty = `source (${payload.source_link})`;

    deriveArticle(payload);
    deriveExercises(payload);

    const templateBuffer = fs.readFileSync(path.join(__dirname, "template.docx"));
    const docBuffer = await createReport({
      template: templateBuffer,
      data: payload,
      cmdDelimiter: ["{","}"],
      processLineBreaksAsNewText: true,
      rejectNullish: false,
      errorHandler: () => ""
    });

    setCors(req, res);
    res.setHeader("Content-Type","application/vnd.openxmlformats-officedocument.wordprocessingml.document");
    res.setHeader("Content-Disposition",'attachment; filename="generated.docx"');
    res.status(200).send(Buffer.from(docBuffer));
  } catch (err) {
    setCors(req, res);
    res.status(500).json({ error: err?.message || String(err) });
  }
};
