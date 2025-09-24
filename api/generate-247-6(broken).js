// api/generate.js
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

/* ---------------- Helpers ---------------- */
function escText(t){return String(t).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;");}
function stripInlineStars(s){ return String(s).replace(/\*\*(.+?)\*\*/g,"$1").replace(/\*([^*]+)\*/g,"$1"); }
function stripHtmlEntities(s){
  return String(s)
    .replace(/&nbsp;/g, " ")
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&#160;/g, " ");
}

// WICHTIG: Behalte Klammern für Modal-Übungen!
function sentenceUnderline(line, preserveBrackets = false){ 
  let result = String(line).replace(/___SENTENCE___/g, "_".repeat(80));
  if (!preserveBrackets) {
    return result;
  }
  // Preserve content in brackets for modal exercises
  return result;
}

function hostToLabel(url) {
  try {
    const u = new URL(String(url));
    const host = u.hostname.replace(/^www\./,"");
    const map = {
      "cfr.org":"CFR","forbes.com":"Forbes","theguardian.com":"The Guardian",
      "nytimes.com":"The New York Times","bbc.com":"BBC","economist.com":"The Economist",
      "ft.com":"Financial Times","wsj.com":"WSJ","reuters.com":"Reuters",
      "apnews.com":"AP","bloomberg.com":"Bloomberg"
    };
    if (map[host]) return map[host];
    const base = host.split(".")[0];
    return base ? base[0].toUpperCase()+base.slice(1) : "source";
  } catch { return "source"; }
}

/* ---------------- HTML → light-md ---------------- */
function normalizeTables(html){
  if (html == null) return "";
  let s = String(html);
  s = s.replace(/<tr[^>]*>/gi,"");
  s = s.replace(/<\/tr>/gi,"\n");
  s = s.replace(/<t[hd][^>]*>/gi,"");
  s = s.replace(/<\/t[hd]>/gi," | ");
  s = s.replace(/\s*\|\s*(\|\s*)+/g," | ");
  return s;
}

// WICHTIG: preserveBrackets für Modal-Übungen
function htmlToLightMd(input,{forActive=false, preserveBrackets=false}={}){
  if (input == null) return "";
  let s = normalizeTables(input);

  s = s.replace(/<h[1-6][^>]*>(.*?)<\/h[1-6]>/gi, (_,g1)=>`\n\n**${g1.trim()}**\n\n`);
  s = s.replace(/<br\s*\/?>/gi,"\n");
  s = s.replace(/<\/p>\s*/gi,"\n\n").replace(/<p[^>]*>/gi,"");
  s = s.replace(/<\/div>\s*/gi,"\n\n").replace(/<div[^>]*>/gi,"");
  s = s.replace(/<li[^>]*>\s*/gi,"- ").replace(/<\/li>/gi,"\n");
  s = s.replace(/<\/?(ul|ol)[^>]*>/gi,"");
  s = s.replace(/<\/?strong>/gi,"**").replace(/<\/?b>/gi,"**");
  s = s.replace(/<\/?em>/gi,"*").replace(/<\/?i>/gi,"*");
  s = s.replace(/<[^>]+>/g,"");
  s = stripHtmlEntities(s).replace(/\r\n?/g,"\n");

  if (forActive){
    s = s.replace(/(?:^|\n)\s*(Phase\s*\d+\s*:)/gi,(m,g1)=>`\n\n**${g1.trim()}**\n`);
    if (!/\n{2,}/.test(s)) s = s.replace(/(?<=[.!?])\s+(?=[A-ZÄÖÜ])/g,"\n\n");
    s = s.replace(/(\s)(\d+\.\s+)/g,"$1\n$2");
  }

  s = s.split("\n").map(l=>l.replace(/\s*\|\s*$/,"").trimRight()).join("\n");
  s = s.replace(/\n{3,}/g,"\n\n").trim();
  return s;
}

/* ---------------- Markdown → Runs (für *_rich) ---------------- */
function splitRuns(md){
  const out=[]; let rest=String(md);
  const re=/(\*\*[^*]+\*\*|\*[^*]+\*)/;
  while(rest.length){
    const m=rest.match(re);
    if(!m){ out.push({t:rest}); break; }
    const [match]=m; const i=m.index;
    if(i>0) out.push({t:rest.slice(0,i)});
    if(match.startsWith("**")) out.push({t:match.slice(2,-2),b:true});
    else out.push({t:match.slice(1,-1),i:true});
    rest = rest.slice(i+match.length);
  }
  return out;
}
function runXml({t,b,i}){
  const pr = (b||i)?`<w:rPr>${b?"<w:b/>":""}${i?"<w:i/>":""}</w:rPr>`:"";
  return `<w:r>${pr}<w:t xml:space="preserve">${escText(t)}</w:t></w:r>`;
}

/* ---------------- items → <w:r>… ---------------- */
function collapseBlanks(items){
  const out=[]; let prevBlank=false;
  for (const it of items){
    if (it.kind==="blank"){ if(!prevBlank){out.push(it); prevBlank=true;} }
    else { out.push(it); prevBlank=false; }
  }
  while (out[0]?.kind==="blank") out.shift();
  while (out[out.length-1]?.kind==="blank") out.pop();
  return out;
}
function itemsToRunsXml(items){
  const collapsed = collapseBlanks(items);
  const parts=[];
  for (let i=0;i<collapsed.length;i++){
    const it=collapsed[i], next=collapsed[i+1];
    if (it.kind==="blank"){ parts.push("<w:br/>"); continue; }
    splitRuns(it.line).forEach(r=>parts.push(runXml(r)));
    if (next && next.kind==="text") parts.push("<w:br/>");
  }
  return parts.join("");
}
const toLiteral = xml => `||${xml}||`;

/* ---------------- Listen/Parsing ---------------- */
function stripLeadingLabels(line){
  return String(line).replace(/^\s*((?:[A-Za-z][\)\.]|[0-9]+\.)\s+|[-•]\s+)/,"");
}

// WICHTIG: Bei Modal-Übungen Klammern erhalten
function enumeratePreserveBlanks(lines, preserveBrackets = false){
  const items=[]; let n=0;
  for (const raw of lines){
    const s = String(raw);
    if (!s.trim()){ items.push({kind:"blank"}); continue; }
    if (/^\s*([0-9]+\.)|[-•]\s+/.test(s)){ items.push({kind:"text", line:s}); continue; }
    n += 1; 
    const processedLine = preserveBrackets ? s : sentenceUnderline(s, true);
    items.push({kind:"text", line:`${n}. ${processedLine}`});
  }
  return items;
}

function toArrayFlexible(v){
  if (v == null) return [];
  if (Array.isArray(v)) return v.map(x=>String(x).trim()).filter(Boolean);
  const s = String(v);
  if (s.includes("|")) return s.split("|").map(t=>t.trim()).filter(Boolean);
  if (s.includes("\n")) return s.split("\n").map(t=>t.trim()).filter(Boolean);
  return s.split(",").map(t=>t.trim()).filter(Boolean);
}

/* ----- Idioms: blockweise MC ----- */
function buildIdiomsBlocks(md){
  const rawLines = md.split(/\n/).map(l=>l.trimRight());
  const blocks=[]; let cur=null; let optIdx=0; const labels="abcd".split("");

  for (const L of rawLines){
    const line = sentenceUnderline(L).trim();
    if (!line){ if(cur) cur.items.push({kind:"blank"}); continue; }
    if (/^\d+\./.test(line)){
      if (cur) blocks.push(cur);
      cur = { question: line, items: [] };
      optIdx = 0;
      continue;
    }
    const clean = stripLeadingLabels(line);
    const lab = labels[Math.min(optIdx, labels.length-1)];
    const opt = `${lab}) ${clean}`;
    optIdx++;
    if (!cur) blocks.push({ question:"", items:[{kind:"text", line:opt}] });
    else cur.items.push({kind:"text", line:opt});
  }
  if (cur) blocks.push(cur);

  const items=[];
  for (const b of blocks){
    if (b.question) items.push({kind:"text", line:b.question});
    b.items.forEach(it=>items.push(it));
    items.push({kind:"blank"});
  }
  return items;
}

/* ----- Matching ----- */
function buildMatchingLines(md){
  const lines = md.split(/\n/).map(s=>s.trim()).filter(Boolean);
  const left=[], right=[];
  for (const s of lines){
    if (/^\d+\./.test(s)) left.push(s.replace(/^\d+\.\s*/,""));
    else if (/^[A-Za-z]\./.test(s)) right.push(s.replace(/^[A-Za-z]\.\s*/,""));
    else if (s.includes("|")) {
      const [a,b] = s.split("|");
      left.push((a||"").trim()); right.push((b||"").trim());
    } else {
      return lines.map(t=>({kind:"text", line:t}));
    }
  }
  const n = Math.max(left.length,right.length);
  const out=[];
  for (let i=0;i<n;i++){
    out.push({kind:"text", line:`${left[i]||""}   |   ${right[i]||""}`});
  }
  return out;
}

/* ---------------- Wortbox Harvest ---------------- */
function harvestWordBox(payload, base){
  // Erweiterte Suche nach allen möglichen Word Box Varianten
  const candidates = [
    payload[`${base}_word_box_content`],
    payload[`${base}_word_box_content_line`],
    payload[`${base}_word_box`],
    payload[`${base}_wordbox`],
    payload[`${base}_options`],
    payload[`${base}_choices`],
    payload[`${base}_words`],
  ];
  
  for (const c of candidates){
    const a = toArrayFlexible(c);
    if (a.length) return a;
  }
  
  // Spezialfall für MC mit gemeinsamer Word Box
  if (base.includes('vocabulary') || base.includes('collocations')) {
    // Prüfe ob es eine globale Word Box für diese Übung gibt
    const mcWordBox = payload[`${base.split('_').slice(0,3).join('_')}_word_box`];
    if (mcWordBox) return toArrayFlexible(mcWordBox);
  }
  
  return [];
}

/* ---------------- Article & Vocab ---------------- */
const MAX_P=16;

function deriveArticle(payload){
  // Source Link mit korrektem Hyperlink
  if (payload.source_link) {
    payload.source_link_pretty = hostToLabel(payload.source_link);
    // Verwende rId1 für externe Links
    payload.source_link_hyperlink_raw = `<w:hyperlink r:id="rId1"><w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr><w:t>${escText(payload.source_link_pretty)}</w:t></w:r></w:hyperlink>`;
  }

  for (let i=1;i<=MAX_P;i++){
    const k = `article_text_paragraph${i}`;
    if (k in payload){
      const md = htmlToLightMd(payload[k]);
      const items = md.split(/\n/)
        .map(s=>sentenceUnderline(s))
        .map(s=>s.trim()?{kind:"text", line:s}:{kind:"blank"});
      payload[`${k}_rich`]  = toLiteral(itemsToRunsXml(items));
      payload[`${k}_plain`] = stripInlineStars(md);
    }
    
    const w1=(payload[`article_vocab_p${i}_1`]||"").toString().trim();
    const w2=(payload[`article_vocab_p${i}_2`]||"").toString().trim();
    const w3=(payload[`article_vocab_p${i}_3`]||"").toString().trim();
    const words=[w1,w2,w3].filter(Boolean);
    if (words.length){
      payload[`article_vocab_p${i}_line`] = words.join("   |   ");
    }
  }

  const paras=[];
  for (let i=1;i<=MAX_P;i++){
    const k=`article_text_paragraph${i}`;
    if (payload[k]) paras.push(htmlToLightMd(payload[k]));
  }
  if (paras.length){
    const lines = paras.flatMap(p=>p.split(/\n{2,}/)).flatMap(p=>p.split("\n"))
      .map(s=>sentenceUnderline(s));
    const items = lines.map(s=>s.trim()?{kind:"text", line:s}:{kind:"blank"});
    payload.article_text_all_rich  = toLiteral(itemsToRunsXml(items));
    payload.article_text_all_plain = stripInlineStars(paras.join("\n\n"));
  }
}

/* ---------------- Exercises ---------------- */
function deriveExercises(payload){
  // Help-Links
  for (const k of Object.keys(payload)){
    if (/^help_link_/i.test(k) && !k.endsWith('_pretty')){
      const url = (payload[k]||"").toString().trim();
      if (!url) continue;
      payload[`${k}_pretty`] = "help";
    }
  }

  // Wortboxen explizit verarbeiten BEVOR Content verarbeitet wird
  for (const k of Object.keys(payload)){
    if (/_word_box_content$/i.test(k)){
      const base = k.replace(/_word_box_content$/i,"");
      const items = toArrayFlexible(payload[k]);
      if (items.length) {
        payload[`${base}_word_box_content_line`] = items.join("   |   ");
      }
    }
  }

  // Content verarbeiten
  for (const k of Object.keys(payload)){
    if (/_content$/i.test(k) && !/_word_box_content$/i.test(k)){
      const base = k.replace(/_content$/i,"");
      const isActive   = /^active_/i.test(base);
      const isIdioms   = /idioms/i.test(base);
      const isMatching = /matching/i.test(base);
      const isModal    = /modal/i.test(base) || /modal/i.test(payload[`${base}_title`] || '');

      let md = htmlToLightMd(payload[k], { 
        forActive:isActive, 
        preserveBrackets:isModal 
      });
      
      let lines = md.split(/\n/).map(s=>s.replace(/\s*\|\s*/g,"   |   ").trimEnd());
      
      // Behalte Klammern bei Modal-Übungen
      if (!isModal) {
        lines = lines.map(sentenceUnderline);
      }

      let items;
      if (isIdioms){
        items = buildIdiomsBlocks(md);
        // Wortbox für Idioms
        if (!payload.hasOwnProperty(`${base}_word_box_content_line`)){
          const wb = harvestWordBox(payload, base);
          if (wb.length) payload[`${base}_word_box_content_line`] =
