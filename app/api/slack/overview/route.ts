import { NextResponse } from 'next/server'

const CHANNEL_FALLBACK = 'C0968BGH8RZ' // #code_pod

export async function POST(request: Request) {
  const token = process.env.SLACK_BOT_TOKEN
  const channel = process.env.SLACK_CHANNEL_ID || CHANNEL_FALLBACK
  if (!token) {
    return NextResponse.json({ error: 'SLACK_BOT_TOKEN is not configured on the server.' }, { status: 500 })
  }

  const { text } = await request.json()
  if (!text || typeof text !== 'string') {
    return NextResponse.json({ error: 'Missing message text.' }, { status: 400 })
  }

  const res = await fetch('https://slack.com/api/chat.postMessage', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json; charset=utf-8', Authorization: `Bearer ${token}` },
    body: JSON.stringify({ channel, text }),
  }).then(r => r.json())

  if (!res.ok) {
    return NextResponse.json({ error: `Slack error: ${res.error}` }, { status: 500 })
  }

  let permalink: string | null = null
  try {
    const pl = await fetch(
      `https://slack.com/api/chat.getPermalink?channel=${encodeURIComponent(res.channel)}&message_ts=${res.ts}`,
      { headers: { Authorization: `Bearer ${token}` } }
    ).then(r => r.json())
    if (pl.ok) permalink = pl.permalink
  } catch { /* best-effort */ }

  return NextResponse.json({ ok: true, ts: res.ts, permalink })
}
