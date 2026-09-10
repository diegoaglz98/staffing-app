import { NextResponse } from 'next/server'

const CHANNEL_FALLBACK = 'C0968BGH8RZ' // #code_pod

export async function POST(request: Request) {
  const token = process.env.SLACK_BOT_TOKEN
  const channel = process.env.SLACK_CHANNEL_ID || CHANNEL_FALLBACK
  if (!token) {
    return NextResponse.json({ error: 'SLACK_BOT_TOKEN is not configured on the server.' }, { status: 500 })
  }

  const { projectName, milestones, weeklyTarget, stoMention } = await request.json()

  const post = (body: Record<string, unknown>) =>
    fetch('https://slack.com/api/chat.postMessage', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json; charset=utf-8', Authorization: `Bearer ${token}` },
      body: JSON.stringify({ channel, ...body }),
    }).then(r => r.json())

  // 1) Parent message starts the thread
  const parent = await post({ text: `:mega: Update requested on *${projectName}*` })
  if (!parent.ok) {
    return NextResponse.json({ error: `Slack error: ${parent.error}` }, { status: 500 })
  }

  // 2) Threaded reply with the details
  const lines: string[] = ['Please provide an update on the following milestones:']
  if (Array.isArray(milestones) && milestones.length > 0) {
    lines.push(milestones.map((m: string) => `  • ${m}`).join('\n'))
  } else {
    lines.push('  _(no open milestones listed)_')
  }
  lines.push('')
  lines.push(
    `And on pacing towards the production plan for this week of *${weeklyTarget ?? 'N/A'}* task${weeklyTarget === 1 ? '' : 's'}.`
  )
  if (stoMention) {
    lines.push('')
    lines.push(`STO: ${stoMention}`)
  }

  const reply = await post({ thread_ts: parent.ts, text: lines.join('\n') })
  if (!reply.ok) {
    return NextResponse.json({ error: `Slack error (thread): ${reply.error}`, ts: parent.ts }, { status: 500 })
  }

  return NextResponse.json({ ok: true, ts: parent.ts })
}
