import { NextResponse } from 'next/server'

const CHANNEL_FALLBACK = 'C0968BGH8RZ' // #code_pod

// Posts a short parent message, then the full text as a threaded reply (staffing update + milestones & PP overview)
export async function POST(request: Request) {
  const token = process.env.SLACK_BOT_TOKEN
  const channel = process.env.SLACK_CHANNEL_ID || CHANNEL_FALLBACK
  if (!token) {
    return NextResponse.json({ error: 'SLACK_BOT_TOKEN is not configured on the server.' }, { status: 500 })
  }

  const { parent, thread } = await request.json()
  if (!parent || typeof parent !== 'string' || !thread || typeof thread !== 'string') {
    return NextResponse.json({ error: 'Missing message text.' }, { status: 400 })
  }

  const post = (body: Record<string, unknown>) =>
    fetch('https://slack.com/api/chat.postMessage', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json; charset=utf-8', Authorization: `Bearer ${token}` },
      body: JSON.stringify({ channel, ...body }),
    }).then(r => r.json())

  const top = await post({ text: parent })
  if (!top.ok) {
    return NextResponse.json({ error: `Slack error: ${top.error}` }, { status: 500 })
  }

  const reply = await post({ thread_ts: top.ts, text: thread })
  if (!reply.ok) {
    return NextResponse.json({ error: `Slack error (thread): ${reply.error}`, ts: top.ts }, { status: 500 })
  }

  let permalink: string | null = null
  try {
    const pl = await fetch(
      `https://slack.com/api/chat.getPermalink?channel=${encodeURIComponent(top.channel)}&message_ts=${top.ts}`,
      { headers: { Authorization: `Bearer ${token}` } }
    ).then(r => r.json())
    if (pl.ok) permalink = pl.permalink
  } catch { /* best-effort */ }

  return NextResponse.json({ ok: true, ts: top.ts, permalink })
}
