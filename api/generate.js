// api/generate.js
// Vercel Serverless Function (Node.js runtime, ESM)
// Zweck: .docx aus Template + Daten erzeugen, mit korrekter Interpretation von {LINK(...)} / {HTML ...}
// WICHTIG: In deinem Template müssen die Platzhalter mit { } stehen (nicht +++).
//          Für HTML-Blocks nutze im Template {HTML dein_feld}. Für Links {LINK({ url: ..., label: ... })}.

import createReport from "docx-templates";
import { readFile } from "node:fs/promises";
import { readFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { createRequire } from "node:module";

// ---- feste Node-Runtime auf Vercel erzwingen (nicht Edge) ----
export const config = { runtime: "nodejs18.x" };

// ---- lib-Version für Debug ermitteln ----
const require = createRequire(import.meta.url);
const docxPkgPath = require.resolve("docx-templates/package.json");
const { version: DOCX_TEMPLATES_VERSION } = JSON.parse(
  readFileSync(docxPkgPath, "utf8")
);

// ---- Hilfen: Pfade & Request-Parsing ----
const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// Wir erwarten das Template im Repo unter: api/template.docx
// Falls du es anders nennst/ablegst, Pfad unten anpassen.
const TEMPLATE_PATH = path.join(process.cwd(), "api", "template.docx");

/**
 * Liest JSON-Body aus Request (bei POST). Bei Fehler -> {}.
 */
async function readJsonBody(req) {
  try {
    const text = await req.text();
    if (!text) return {};
    return JSON.parse(text);
  } catch {
    return {};
  }
}

/**
 * Hilfsfunktion: Data aus Query (GET) lesen (nur für schnelle Tests)
 * /api/generate?source_link=https://…&source_link_pretty=NYTimes
 */
function dataFromQuery(urlObj) {
  const d = {};
  for (const [k, v] of urlObj.searchParams.entries()) d[k] = v;
  return d;
}

/**
 * Optionale Sanity-Fixes für Daten
 * - Falls source_link_pretty fehlt, aus URL host ableiten (nur kosmetisch)
 */
function normalizeData(data) {
  const out = { ...data };

  if (!out.source_link_pretty && typeof out.source_link === "string") {
    try {
      const u = new URL(out.source_link);
      out.source_link_pretty = u.hostname.replace(/^www\./, "");
    } catch {
      // lasse wie es ist
    }
  }

  return out;
}

/**
 * Liefert standardisierte Fehlerantwort (JSON)
 */
function jsonResponse(obj, status = 200, extraHeaders = {}) {
  return new Response(JSON.stringify(obj), {
    status,
    headers: {
      "content-type": "application/json",
      ...extraHeaders,
    },
  });
}

/**
 * GET-Handler:
 * - ?debug=version -> Versionsinfo
 * - Nutzung über GET für Tests erlaubt (Query-Params -> data)
 */
export async function GET(request) {
  const url = new URL(request.url);

  // ---- Debug-Endpunkt ----
  if (url.searchParams.get("debug") === "version") {
    return jsonResponse(
      {
        ok: true,
        docxTemplatesVersion: DOCX_TEMPLATES_VERSION,
        cmdDelimiter: "{ }",
        templatePath: "/api/template.docx",
      },
      200,
      { "x-docx-templates-version": DOCX_TEMPLATES_VERSION }
    );
  }

  // ---- Ad-hoc-Generierung mit Query-Daten (nur zu Testzwecken) ----
  const data = normalizeData(dataFromQuery(url));
  try {
    const template = await readFile(TEMPLATE_PATH);
    const buffer = await createReport({
      template,
      data,
      // ENTSCHEIDEND: {}-Delimiter, sonst werden {LINK}/{HTML} nicht erkannt!
      cmdDelimiter: ["{", "}"],
      // Sinnvolle Defaults:
      fixSmartQuotes: true,
      rejectNullish: false, // nicht hart abbrechen bei leeren Feldern
    });

    return new Response(buffer, {
      status: 200,
      headers: {
        "content-type":
          "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        "content-disposition": 'attachment; filename="generated.docx"',
        "x-docx-templates-version": DOCX_TEMPLATES_VERSION,
      },
    });
  } catch (err) {
    return jsonResponse(
      {
        ok: false,
        error: String(err && err.message ? err.message : err),
      },
      500,
      { "x-docx-templates-version": DOCX_TEMPLATES_VERSION }
    );
  }
}

/**
 * POST-Handler:
 * - Erwartet JSON-Body = { ...deine Felder... }
 * - Gibt .docx zurück
 *
 * WICHTIG: Für fette Formatierung in „Active & Cooperative“ müssen die
 * entsprechenden Template-Platzhalter als {HTML deine_felder} angelegt sein.
 * (Im Template selbst – nicht hier im Code.)
 */
export async function POST(request) {
  try {
    const bodyData = await readJsonBody(request);
    const data = normalizeData(bodyData);

    const template = await readFile(TEMPLATE_PATH);
    const buffer = await createReport({
      template,
      data,
      cmdDelimiter: ["{", "}"], // <<< Kernfix für {LINK}/{HTML}
      fixSmartQuotes: true,
      rejectNullish: false,
    });

    return new Response(buffer, {
      status: 200,
      headers: {
        "content-type":
          "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        "content-disposition": 'attachment; filename="generated.docx"',
        "x-docx-templates-version": DOCX_TEMPLATES_VERSION,
      },
    });
  } catch (err) {
    return jsonResponse(
      {
        ok: false,
        error: String(err && err.message ? err.message : err),
      },
      500,
      { "x-docx-templates-version": DOCX_TEMPLATES_VERSION }
    );
  }
}
