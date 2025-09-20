module.exports = async (req, res) => {
  try {
    // CORS Headers setzen
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
    res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

    // OPTIONS Request beantworten
    if (req.method === 'OPTIONS') {
      return res.status(200).end();
    }

    // Nur POST erlauben
    if (req.method !== 'POST') {
      return res.status(405).json({ error: 'Method not allowed' });
    }

    // Einfache Antwort erstmal
    return res.status(200).json({
      success: true,
      message: 'API is working',
      received: req.body || {}
    });

  } catch (error) {
    console.error('API Error:', error);
    return res.status(500).json({
      error: 'Internal server error',
      details: error.message
    });
  }
};
