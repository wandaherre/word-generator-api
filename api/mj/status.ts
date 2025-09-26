import type { NextApiRequest, NextApiResponse } from 'next';

export default async function handler(req: NextApiRequest, res: NextApiResponse) {
  if (req.method !== 'GET') return res.status(405).json({ error: 'Method not allowed' });
  const { taskId } = req.query;
  if (!taskId || Array.isArray(taskId)) return res.status(400).json({ error: 'Missing taskId' });

  try {
    const u = new URL('https://api.kie.ai/api/v1/mj/record-info');
    u.searchParams.set('taskId', String(taskId));
    const r = await fetch(u, { headers: { Authorization: `Bearer ${process.env.KIE_API_KEY}` }});
    const json = await r.json();
    return res.status(r.ok ? 200 : r.status).json(json);
  } catch (e:any) {
    return res.status(500).json({ error: e?.message || 'Upstream error' });
  }
}
