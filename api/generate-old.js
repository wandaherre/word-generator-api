// api/generate.js
const fs = require("fs");
const path = require("path");
// WICHTIG: Default-Export korrekt holen
const createReport = require("docx-templates").default;

module.exports = async (req, res) => {
  // CORS + no-store
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "POST,OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type,Authorization");
  res.setHeader("Cache-Control", "no-store");

  if (req.method === "OPTIONS") { res.status(200).end(); return; }
  if (req.method !== "POST") { res.status(405).json({ error: "Method Not Allowed" }); return; }

  try {
    // Body robust lesen (Vercel liefert meist schon geparst)
    let payload = req.body;
    if (typeof payload === "string") {
      try { payload = JSON.parse(payload); } catch { payload = {}; }
    }
    if (!payload || typeof payload !== "object") payload = {};

    // Template muss via includeFiles gebundled sein (siehe vercel.json)
    const templatePath = path.join(__dirname, "template.docx");
    const templateBuffer = fs.readFileSync(templatePath);

    // DOCX erzeugen – docx-templates nutzt SINGLE BRACES {key}
    const docBuffer = await createReport({
      template: templateBuffer,
      data: payload,
      cmdDelimiter: ["{", "}"]
    });

    // KORREKTER MIME-TYP + Download-Header, NUR Binary senden
    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    );
    res.setHeader("Content-Disposition", 'attachment; filename="generated.docx"');
    res.status(200).send(Buffer.from(docBuffer));
  } catch (err) {
    res.status(500).json({ error: err && err.message ? err.message : String(err) });
  }
};
