// api/generate.js
// Vercel Serverless (CommonJS) + docx-templates, robust gegen fehlende Felder/Arrays

const fs = require("fs");
const path = require("path");
const createReport = require("docx-templates").default;

// -------- Utility: defensive payload hardening ----------
function hardenPayload(input) {
  const seen = new WeakSet();

  const toText = (v) => {
    if (v == null) return "";
    if (Array.isArray(v)) {
      return v.map(x => {
        if (x == null) return "";
        if (typeof x === "object") {
          const t = x.text ?? x.sentence ?? x.value ?? x.title ?? null;
          return t != null ? String(t) : JSON.stringify(x);
        }
        return String(x);
      }).join("\n");
    }
    if (typeof v === "object") {
      const t = v.text ?? v.sentence ?? v.value ?? v.title ?? null;
      return t != null ? String(t) : JSON.stringify(v);
    }
    return String(v);
  };

  const normalize = (obj) => {
    if (!obj || typeof obj !== "object") return obj;
    if (seen.has(obj)) return obj;
    seen.add(obj);

    // 1) Falls ein Objekt "options" erwartet, aber fehlt → setze leeres Array
    if (obj && typeof obj === "object" && !Array.isArray(obj)) {
      if (!("options" in obj) || obj.options == null) obj.options = [];
      if (!Array.isArray(obj.options)) obj.options = [obj.options].filter(Boolean);
    }

    // 2) Rekursiv über Arrays/Objekte
    for (const k of Object.keys(obj)) {
      const val = obj[k];
      if (Array.isArray(val)) {
        obj[k] = val.map(el => {
          if (el && typeof el === "object") {
            // Für Elemente, die selbst "options" nutzen könnten
            if (!("options" in el) || el.options == null) el.options = [];
            if (!Array.isArray(el.options)) el.options = [el.options].filter(Boolean);
            return normalize(el);
          }
