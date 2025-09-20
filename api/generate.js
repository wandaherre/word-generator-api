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
    const templatePath = path.resolve(__dirname, 'template.docx');
    
    if (!fs.existsSync(templatePath)) {
      return res.status(404).json({ error: 'Template not found' });
    }

    const template = fs.readFileSync(templatePath);
    console.log('=== DOCX-TEMPLATES DEBUG ===');
    console.log('Template size:', template.length, 'bytes');
    console.log('Input data:', JSON.stringify(req.body, null, 2));

    // Minimale Testdaten falls Body leer
    const data = Object.keys(req.body || {}).length === 0 ? {
      title: 'Test Title',
      content: 'Test Content'
    } : req.body;

    console.log('Using data:', JSON.stringify(d
