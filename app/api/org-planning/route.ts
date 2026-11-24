import { NextResponse } from 'next/server'
import { cookies } from 'next/headers'
import { createClient } from '@supabase/supabase-js'
import { createHash, timingSafeEqual } from 'crypto'

// Org planning is behind its own password. Its table has RLS on with no policies, so the
// browser's anon key can't read it — every read/write goes through this route using the
// service-role key, and only after the planning password has been entered.
const PASSWORD = process.env.ORG_PLANNING_PASSWORD || 'SpartanPlanning2027'
const COOKIE = 'org_planning'
const TABLE = 'org_plan_people'
const token = () => createHash('sha256').update(`org-planning:${PASSWORD}`).digest('hex')

async function unlocked() {
  const v = (await cookies()).get(COOKIE)?.value ?? ''
  const t = token()
  return v.length === t.length && timingSafeEqual(Buffer.from(v), Buffer.from(t))
}
// Only a service-role / secret key bypasses RLS. Returns why the configured key won't work, or null.
function keyProblem(key: string) {
  if (key.startsWith('sb_secret_')) return null
  if (key.startsWith('sb_publishable_')) return 'SUPABASE_SERVICE_ROLE_KEY is set to the publishable key — it needs the secret key (sb_secret_…).'
  try {
    const role = JSON.parse(Buffer.from(key.split('.')[1], 'base64url').toString()).role
    if (role === 'service_role') return null
    return `SUPABASE_SERVICE_ROLE_KEY is set to the "${role}" key — it needs the service_role key.`
  } catch {
    return "SUPABASE_SERVICE_ROLE_KEY doesn't look like a Supabase key."
  }
}
function db() {
  const url = process.env.NEXT_PUBLIC_SUPABASE_URL
  const key = (process.env.SUPABASE_SERVICE_ROLE_KEY ?? '').trim()
  if (!url || !key) return { error: 'SUPABASE_SERVICE_ROLE_KEY is not configured on the server.' }
  const problem = keyProblem(key)
  if (problem) return { error: problem }
  return { sb: createClient(url, key, { auth: { persistSession: false } }) }
}
const fail = (error: string, status = 500) => NextResponse.json({ error }, { status })

export async function GET() {
  if (!(await unlocked())) return fail('locked', 401)
  const { sb, error: keyError } = db()
  if (!sb) return fail(`setup: ${keyError}`)
  const { data, error } = await sb.from(TABLE).select('*').order('name', { ascending: true })
  if (error) return fail(`setup: ${error.message}`)
  return NextResponse.json({ people: data })
}

type Person = { id?: string; name?: string; position?: string | null; team_name?: string | null; parent_id?: string | null; in_chart?: boolean }

export async function POST(request: Request) {
  const body = await request.json().catch(() => ({}))

  // Unlock / lock
  if (body.action === 'unlock') {
    const given = String(body.password ?? '')
    const ok = given.length === PASSWORD.length && timingSafeEqual(Buffer.from(given), Buffer.from(PASSWORD))
    if (!ok) return fail('Wrong password', 401)
    const res = NextResponse.json({ ok: true })
    res.cookies.set(COOKIE, token(), {
      httpOnly: true, secure: process.env.NODE_ENV === 'production', sameSite: 'strict', maxAge: 60 * 60 * 12, path: '/',
    })
    return res
  }
  if (body.action === 'lock') {
    const res = NextResponse.json({ ok: true })
    res.cookies.set(COOKIE, '', { httpOnly: true, maxAge: 0, path: '/' })
    return res
  }

  if (!(await unlocked())) return fail('locked', 401)
  const { sb, error: keyError } = db()
  if (!sb) return fail(`setup: ${keyError}`)
  const clean = (p: Person) => {
    const out: Person = {}
    if (p.name !== undefined) out.name = String(p.name).trim()
    if (p.position !== undefined) out.position = (p.position ?? '').toString().trim() || null
    if (p.team_name !== undefined) out.team_name = (p.team_name ?? '').toString().trim() || null
    if (p.parent_id !== undefined) out.parent_id = p.parent_id || null
    if (p.in_chart !== undefined) out.in_chart = !!p.in_chart
    return out
  }

  switch (body.action) {
    case 'add': {
      const row = clean(body.person ?? {})
      if (!row.name) return fail('Name is required', 400)
      const { data, error } = await sb.from(TABLE).insert([row]).select().single()
      return error ? fail(error.message) : NextResponse.json({ person: data })
    }
    case 'update': {
      const { data, error } = await sb.from(TABLE).update(clean(body.patch ?? {})).eq('id', body.id).select().single()
      return error ? fail(error.message) : NextResponse.json({ person: data })
    }
    case 'updateMany': {
      // [{ id, patch }] — used for reattaching reports / removing branches
      for (const u of (body.updates ?? []) as { id: string; patch: Person }[]) {
        const { error } = await sb.from(TABLE).update(clean(u.patch)).eq('id', u.id)
        if (error) return fail(error.message)
      }
      return NextResponse.json({ ok: true })
    }
    case 'delete': {
      const { error } = await sb.from(TABLE).delete().eq('id', body.id)
      return error ? fail(error.message) : NextResponse.json({ ok: true })
    }
    case 'import': {
      // [{ name, position }] — new names are added, existing names (case-insensitive) get their position updated
      const rows = ((body.rows ?? []) as Person[]).map(clean).filter(r => r.name)
      const { data: existing, error: e1 } = await sb.from(TABLE).select('id,name')
      if (e1) return fail(e1.message)
      const byName = new Map((existing ?? []).map(r => [String(r.name).toLowerCase(), r.id as string]))
      const inserts: Person[] = []
      let updated = 0
      for (const r of rows) {
        const id = byName.get(r.name!.toLowerCase())
        if (id) {
          if (r.position) { await sb.from(TABLE).update({ position: r.position }).eq('id', id); updated++ }
        } else {
          inserts.push(r)
          byName.set(r.name!.toLowerCase(), 'pending')
        }
      }
      if (inserts.length) {
        const { error } = await sb.from(TABLE).insert(inserts)
        if (error) return fail(error.message)
      }
      return NextResponse.json({ added: inserts.length, updated })
    }
    case 'clear': {
      const { error } = await sb.from(TABLE).delete().neq('id', '00000000-0000-0000-0000-000000000000')
      return error ? fail(error.message) : NextResponse.json({ ok: true })
    }
    default:
      return fail('Unknown action', 400)
  }
}
