// api/generate.js
const fs = require("fs");
const path = require("path");
const createReport = require("docx-templates");

module.exports = async (req, res) => {
  // CORS + kein Cache
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "POST,OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type,Authorization");
  res.setHeader("Cache-Control", "no-store");

  if (req.method === "OPTIONS") {
    res.status(200).end();
    return;
  }
  if (req.method !== "POST") {
    res.status(405).json({ error: "Method Not Allowed" });
    return;
  }

  try {
    const payload = req.body && typeof req.body === "object" ? req.body : {};

    // Template muss per includeFiles gebundled sein
    const templatePath = path.join(__dirname, "template.docx");
    const templateBuffer = fs.readFileSync(templatePath);

    // DOCX erzeugen (docx-templates: SINGLE BRACES {key})
    const docBuffer = await createReport({
      template: templateBuffer,
      data: payload,
      cmdDelimiter: ["{", "}"]
    });

    // KORREKTE MIME-TYPE & Download-Header (keine JSON-Hülle!)
    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    );
    res.setHeader("Content-Disposition", 'attachment; filename="generated.docx"');

    res.status(200).send(Buffer.from(docBuffer));
  } catch (err) {
    // Nur im Fehlerfall JSON zurückgeben
    res.status(500).json({ error: err && err.message ? err.message : String(err) });
  }
};
