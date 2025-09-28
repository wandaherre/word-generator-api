// /api/mj/generate.ts
// Zweck: Kie.ai-Task anlegen und taskId zurückgeben
// Methode: POST
// Body: { prompt: string, version?: string, aspectRatio?: string, speed?: string, webhook?: boolean, enableTranslation?: boolean }

import type { VercelRequest, VercelResponse } from '@vercel/node';

function withCORS(req: VercelRequest, res: VercelResponse) {
  res.setHeader('Access-Control-Allow-Origin', '*'); // bei Bedarf Origin einschränken
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
    const body = typeof req.body === 'string' ? JSON.parse(req.body || '{}') : (req.body || {});
    const {
      prompt,
      version = '7',
      aspectRatio = '16:9',
      speed = 'fast',            // für Tests schneller
      webhook = false,
      enableTranslation = true,  // WICHTIG: deutschsprachige Prompts automatisch übersetzen
    } = body;

    if (!prompt || typeof prompt !== 'string') {
      return res.status(422).json({ error: 'Missing or invalid "prompt"' });
    }

    const callBackUrl = webhook && process.env.APP_BASE_URL
      ? `${process.env.APP_BASE_URL.replace(/\/+$/, '')}/api/mj/webhook`
      : undefined;

    const payload: Record<string, any> = {
      taskType: 'mj_txt2img',
      prompt,
      version,
      aspectRatio,
      speed,
      enableTranslation,
      ...(callBackUrl ? { callBackUrl } : {})
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
    console.log('[KIEAI][GENERATE] status=%s body=%s', upstream.status, raw);

    let data: any = {};
    try { data = raw ? JSON.parse(raw) : {}; } catch {}

    if (!upstream.ok) {
      return res.status(upstream.status).json(data || { error: `Upstream ${upstream.status}`, raw });
    }

    // Erwartet: { code:200, data:{ taskId:"..." } }
    return res.status(200).json(data);
  } catch (err: any) {
    console.error('[KIEAI][GENERATE][ERROR]', err);
    return res.status(500).json({ error: err?.message || 'Internal error' });
  }
}
