export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', 'https://satco-mobilisation.vercel.app');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type, Authorization');
  if (req.method === 'OPTIONS') return res.status(200).end();
  // The AI key is paid for by SATCO: only signed-in portal users (HR database login) may use it.
  // The portal adds its session token automatically (storage-guard.js).
  const token = String(req.headers.authorization || '').replace(/^Bearer\s+/i, '');
  const ok = token && (await fetch('https://oaerqjrkdpuhiproppaz.supabase.co/auth/v1/user', {
    headers: { apikey: 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Im9hZXJxanJrZHB1aGlwcm9wcGF6Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3Nzk5NTQ0NjksImV4cCI6MjA5NTUzMDQ2OX0.qBtb3OV1aFGX8e1QUg19qZmOwIIjipF6IZwBOLXY3YI', Authorization: 'Bearer ' + token },
  })).ok;
  if (!ok) return res.status(401).json({ error: { message: 'Not signed in' } });
  try {
    const response = await fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'x-api-key': process.env.ANTHROPIC_API_KEY,
        'anthropic-version': '2023-06-01',
      },
      body: JSON.stringify(req.body),
    });
    const data = await response.json();
    res.status(response.status).json(data);
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
}
