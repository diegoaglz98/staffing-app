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
  const parent = await post({ text: `:thread: Update requested on *${projectName}*` })
  if (!parent.ok) {
    return NextResponse.json({ error: `Slack error: ${parent.error}` }, { status: 500 })
  }

  // 2) Threaded reply with the details
  const hasMilestones = Array.isArray(milestones) && milestones.length > 0
  const hasPP = weeklyTarget !== null && weeklyTarget !== undefined
  const lines: string[] = []

  if (!hasMilestones && !hasPP) {
    // Neither set → standard "overall update" template
    lines.push('Please provide an overall status update on this project:')
    lines.push('  • Current progress since the last update')
    lines.push('  • Any blockers or risks')
    lines.push('  • Focus / next steps for this week')
    lines.push('')
    lines.push('Also, is there a confirmed production plan for this week? If so, please share the target and how you are pacing against it.')
  } else {
    if (hasMilestones) {
      lines.push('Please provide an update on the following milestones:')
      lines.push(milestones.map((m: string) => `  • ${m}`).join('\n'))
    }
    if (hasPP) {
      if (hasMilestones) lines.push('')
      lines.push(
        `${hasMilestones ? 'And on' : 'Please provide an update on'} pacing towards the production plan for this week of *${weeklyTarget}* task${weeklyTarget === 1 ? '' : 's'}.`
      )
    } else if (hasMilestones) {
      // Milestones but no PP → ask whether a confirmed PP already exists
      lines.push('')
      lines.push('The production plan for this week has not been set here — is there a confirmed production plan already? If so, please share the target and how you are pacing against it.')
    }
  }

  if (stoMention) {
    lines.push('')
    lines.push(`STO: ${stoMention}`)
  }

  const reply = await post({ thread_ts: parent.ts, text: lines.join('\n') })
  if (!reply.ok) {
    return NextResponse.json({ error: `Slack error (thread): ${reply.error}`, ts: parent.ts }, { status: 500 })
  }

  // Fetch a clickable permalink to the thread
  let permalink: string | null = null
  try {
    const pl = await fetch(
      `https://slack.com/api/chat.getPermalink?channel=${encodeURIComponent(parent.channel)}&message_ts=${parent.ts}`,
      { headers: { Authorization: `Bearer ${token}` } }
    ).then(r => r.json())
    if (pl.ok) permalink = pl.permalink
  } catch { /* permalink is best-effort */ }

  return NextResponse.json({ ok: true, ts: parent.ts, permalink })
}
