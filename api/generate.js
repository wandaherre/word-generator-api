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
    // Check if docx-templates can be imported
    const { createReport } = require('docx-templates');
    
    return res.status(200).json({
      success: true,
      message: 'docx-templates loaded successfully',
      hasTemplate: fs.existsSync(path.resolve(__dirname, 'template.docx'))
    });
  } catch (error) {
    return res.status(500).json({
      error: 'docx-templates failed to load',
      details: error.message
    });
  }
};
