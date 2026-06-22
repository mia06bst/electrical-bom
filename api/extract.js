export default async function handler(req, res) {
  if (req.method !== 'POST') return res.status(405).end();

  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

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
    const text = await response.text();
    let data;
    try { data = JSON.parse(text); } catch { return res.status(500).json({ error: { message: text } }); }
    res.status(response.status).json(data);
  } catch (e) {
    res.status(500).json({ error: { message: e.message } });
  }
}
