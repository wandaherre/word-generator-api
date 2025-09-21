// api/generate.js
// DOCX via docx-templates. Liefert *_rich als Literal-XML (||...||) und erzeugt:
// - Artikel gesamt: article_text_all_rich (optional)
// - Konditionssichere *_content_rich (Absätze, Bullet/Nummern, Abstand)
// - Unterstriche: Grammar=13, Idioms=20, komplette Satzlinie breit
// - MC-Labels a) b) c) bei Idioms
// - Wortboxen: *_word_box_content_line (mit exakt "   |   ") + *_rich
// - Vokabel je Absatz: *_line + *_rich
// - Abitur wird NICHT gefiltert
// - CORS dynamisch, Preflight ok

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

/* ---------- Sanitizer: HTML -> leichtes Markdown ---------- */
function htmlToLightMd(input) {
  if (input == null) return "";
  let s = String(input);
  s = s.replace(/<br\s*\/?>/gi, "\n");
  s = s.replace(/<\/p>\s*/gi, "\n\n").replace(/<p[^>]*>/gi, "");
  s = s.replace(/<li[^>]*>\s*/gi, "- ").replace(/<\/li>/gi, "\n");
  s = s.replace(/<\/?(ul|ol)[^>]*>/gi, "");
  s = s.replace(/<\/?strong>/gi, "**").replace(/<\/?b>/gi, "**");
  s = s.replace(/<\/?em>/gi, "*").replace(/<\/?i>/gi, "*");
  s = s.replace(/<[^>]+>/g, "");
  s = s.replace(/\r\n?/g, "\n");
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
function runXml({t,b,i}){const pr=(b||i)?`<w:rPr>${b?"<w:b/>":""}${i?"<w:i/>":""}</w:rPr>`:"";return `<w:r>${pr}<w:t xml:space="preserve">${escText(t)}</w:t></w:r>`;}
function linesToRunsXml(lines){const parts=[];lines.forEach((ln,idx)=>{splitRuns(ln).forEach(r=>parts.push(runXml(r))); if(idx<lines.length-1) parts.push("<w:br/><w:br/>");}); return parts.join("");} // doppeltes <w:br/> = mehr Abstand
function toLiteral(runsXml){ return `||${runsXml}||`; }

/* ---------- Helpers: Unterstriche, Nummerierung, MC-Labels ---------- */
function repeatChar(ch,n){return new Array(n+1).join(ch);}
function widenGaps(line, underscoreLen){
  // Ersetze „___“ durch längere Linien, belasse vorhandene 5+ Unterstriche
  return line.replace(/_{3}(?=\b)/g, repeatChar("_", underscoreLen));
}
function sentenceUnderline(line){
  // Für „ganzer Satz hinschreiben“: Linie über (fast) ganze Zeile
  // Heuristik: Placeholder „___SENTENCE___“ → ersetze
  return line.replace(/___SENTENCE___/g, repeatChar("_", 80));
}
function autoEnumerate(lines){
  // Nummeriere 1.,2.,3. wenn keine Bullets/Nummern vorhanden
  const hasMarkers = lines.some(s => /^\s*([0-9]+\.)|[-•]\s+/.test(s));
  if (hasMarkers) return lines;
  return lines.map((s,i)=>`${i+1}. ${s}`);
}
function labelChoicesABC(lines){
  const labels="abcdefghijklmnopqrstuvwxyz".split("");
  return lines.map((s,i)=>`${labels[i] || String.fromCharCode(97+i)}) ${s.replace(/^\s*[-•]\s+/, "")}`);
}

/* ---------- Ableitungen ---------- */
const MAX_P=16;

function deriveArticle(payload){
  // Einzelabsätze (für bedingte Darstellung)
  for(let i=1;i<=MAX_P;i++){
    const k=`article_text_paragraph${i}`;
    if(k in payload){
      const md = htmlToLightMd(payload[k]);
      const lines = md.split(/\n/).filter(x=>x.length||md==="");
      payload[`${k}_rich`] = toLiteral(linesToRunsXml(lines));
    }
    const w1=(payload[`article_vocab_p${i}_1`]||"").toString().trim();
    const w2=(payload[`article_vocab_p${i}_2`]||"").toString().trim();
    const w3=(payload[`article_vocab_p${i}_3`]||"").toString().trim();
    const words=[w1,w2,w3].filter(Boolean);
    if(words.length){
      payload[`article_vocab_p${i}_line`] = words.join("   |   ");
      payload[`article_vocab_p${i}_rich`] = toLiteral(linesToRunsXml(words));
    }
  }
  // Gesamtartikel (optional)
  const paras=[];
  for(let i=1;i<=MAX_P;i++){
    const k=`article_text_paragraph${i}`;
    if(payload[k]) paras.push(htmlToLightMd(payload[k]));
  }
  if(paras.length){
    const lines = paras.flatMap(p=>p.split(/\n{2,}/)).flatMap(p=>p.split("\n"));
    payload.article_text_all_rich = toLiteral(linesToRunsXml(lines));
  }
}

function deriveExercises(payload){
  for(const k of Object.keys(payload)){
    // Wortboxen
    if(/_word_box_content$/i.test(k)){
      let raw=(payload[k]||"").toString().trim(); if(!raw) continue;
      let items;
      if(raw.includes("|")) items=raw.split("|");
      else if(raw.includes("\n")) items=raw.split("\n");
      else items=raw.split(",");
      items=items.map(s=>s.trim()).filter(Boolean);
      payload[`${k}_line`] = items.join("   |   ");
      payload[`${k}_rich`] = toLiteral(linesToRunsXml(items));
    }
    // Inhalte
    if(/_content$/i.test(k)){
      const val = htmlToLightMd(payload[k]);
      let lines = val.split(/\n/);

      // Unterstriche anpassen je Bereich
      const isIdioms = /idioms/i.test(k);
      const isGrammar = /b1_|b2_|grammar/i.test(k) && !isIdioms;
      lines = lines.map(s => sentenceUnderline(s));
      lines = lines.map(s => widenGaps(s, isIdioms ? 20 : (isGrammar ? 13 : 13)));

      // MC-Labels bei Idioms, wenn es wie Optionen aussieht
      if (isIdioms) {
        const looksLikeOptions = lines.length>=3 && lines.every(x=>x.trim().length>0);
        if (looksLikeOptions) lines = labelChoicesABC(lines);
      }

      // Nummerierung, wenn keine Marker vorhanden und nicht MC
      if (!isIdioms) lines = autoEnumerate(lines);

      payload[`${k}_rich`] = toLiteral(linesToRunsXml(lines));
    }
    // Help-Links
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

    // harmlose Defaults
    if (typeof payload.midjourney_article_logo === "undefined") payload.midjourney_article_logo = "";
    if (typeof payload.teacher_cloud_logo === "undefined") payload.teacher_cloud_logo = "";
    if (payload.headline_article && !payload.headline_artikel) payload.headline_artikel = payload.headline_article;
    if (payload.headline_artikel && !payload.headline_article) payload.headline_article = payload.headline_artikel;

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
