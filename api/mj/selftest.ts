import type { VercelRequest, VercelResponse } from '@vercel/node';

function withCORS(req: VercelRequest, res: VercelResponse) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST,OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type, Authorization');
  if (req.method === 'OPTIONS') { res.status(200).end(); return true; }
  return false;
}

export default async function handler(req: VercelRequest, res: VercelResponse) {
  if (withCORS(req, res)) return;
  if (req.method !== 'POST') return res.status(405).json({ error: 'Method not allowed' });

  const API_KEY = process.env.KIE_API_KEY;
  if (!API_KEY) return res.status(500).json({ error: 'Missing KIE_API_KEY env var' });

  try {
    const payload = {
      taskType: 'mj_txt2img',
      prompt: 'test illustration, clean composition, high quality',
      version: '7',
      aspectRatio: '16:9',
      speed: 'relaxed'
    };

    const upstream = await fetch('https://api.kie.ai/api/v1/mj/generate', {
      method: 'POST',
      headers: {
        'Authorization': `Bearer ${API_KEY}`,
        'Content-Type': 'application/json',
      },
      body: JSON.stringify(payload),
    });

    const raw = await upstream.text();
    console.log('[KIEAI][SELFTEST][GENERATE] status=%s body=%s', upstream.status, raw);

    let data: any = {};
    try { data = raw ? JSON.parse(raw) : {}; } catch {}

    return res.status(upstream.status).json(data || { raw });
  } catch (err: any) {
    console.error('[KIEAI][SELFTEST][ERROR]', err);
    return res.status(500).json({ error: err?.message || 'Internal error' });
  }
}
