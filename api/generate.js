const fs = require('fs');
const path = require('path');
const { createReport } = require('docx-templates');

module.exports = async (req, res) => {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, PUT, DELETE, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'X-Requested-With, Content-Type, Accept, Authorization');

  if (req.method === 'OPTIONS') {
    res.status(200).end();
    return;
  }

  if (req.method !== 'POST') {
    res.status(405).json({ error: 'Method Not Allowed' });
    return;
  }

  try {
    // Debug: Template-Info
    const templatePath = path.resolve(__dirname, 'template.docx');
    console.log('Template path:', templatePath);
    console.log('Template exists:', fs.existsSync(templatePath));
    
    if (!fs.existsSync(templatePath)) {
      return res.status(404).json({ error: 'Template not found' });
    }

    const template = fs.readFileSync(templatePath);
    console.log('Template size:', template.length, 'bytes');

    // Debug: Input data
    console.log('Input data:', JSON.stringify(req.body, null, 2));

    // Teste ohne Template-Engine
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', 'attachment; filename=debug.docx');
    
    // Sende Original-Template zurück (ohne Variablen-Ersetzung)
    res.status(200).send(template);

  } catch (error) {
    console.error('Error:', error);
    res.status(500).json({ error: error.message, stack: error.stack });
  }
};
