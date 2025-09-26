import type { NextApiRequest, NextApiResponse } from 'next';

export default async function handler(req: NextApiRequest, res: NextApiResponse) {
  if (req.method !== 'POST') return res.status(405).json({ error: 'Method not allowed' });
  const { prompt, version = '7', aspectRatio = '16:9', speed = 'relaxed', webhook = true } = req.body || {};
  if (!process.env.KIE_API_KEY) return res.status(500).json({ error: 'Missing KIE_API_KEY' });
  if (!prompt) return res.status(400).json({ error: 'Missing prompt' });

  const callBackUrl = webhook ? `${process.env.APP_BASE_URL}/api/mj/webhook` : undefined;

  const body = {
    taskType: 'mj_txt2img',     // Text→Bild
    prompt,
    version,                    // '7' | '6.1' | '6' | '5.2' | '5.1' | 'niji6'
    aspectRatio,                // '1:1' | '16:9' | ... (siehe Docs)
    speed,                      // 'relaxed' | 'fast' | 'turbo'
    ...(callBackUrl ? { callBackUrl } : {})
  };

  try {
    const r = await fetch('https://api.kie.ai/api/v1/mj/generate', {
      method: 'POST',
      headers: {
        'Authorization': `Bearer ${process.env.KIE_API_KEY}`,
        'Content-Type': 'application/json'
      },
      body: JSON.stringify(body)
    });

    const json = await r.json();
    if (!r.ok) return res.status(r.status).json(json);
    // Erwartet: { code:200, data:{ taskId:"..." } }
    return res.status(200).json(json);
  } catch (e:any) {
    return res.status(500).json({ error: e?.message || 'Upstream error' });
  }
}
