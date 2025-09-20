const fs = require('fs');
const path = require('path');

module.exports = async (req, res) => {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') {
    return res.status(200).end();
  }

  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method not allowed' });
  }

  try {
    // Test docx-templates import
    const { createReport } = require('docx-templates');
    
    // Check template file
    const templatePath = path.resolve(__dirname, 'template.docx');
    const templateExists = fs.existsSync(templatePath);
    
    if (!templateExists) {
      return res.status(404).json({
        error: 'Template not found',
        path: templatePath
      });
    }
    
    return res.status(200).json({
      success: true,
      message: 'docx-templates loaded successfully',
      templateExists: true,
      templatePath: templatePath
    });
    
  } catch (error) {
    return res.status(500).json({
      error: 'docx-templates
