import { NextResponse } from 'next/server'
import { cookies } from 'next/headers'
import { createClient } from '@supabase/supabase-js'
import { createHash, randomUUID, timingSafeEqual } from 'crypto'

// Org planning is behind its own password. Its table has RLS on with no policies, so the
// browser's anon key can't read it — every read/write goes through this route using the
// service-role key, and only after the planning password has been entered.
const PASSWORD = process.env.ORG_PLANNING_PASSWORD || 'SpartanPlanning2027'
const COOKIE = 'org_planning'
const TABLE = 'org_plan_people'
const PODS = 'org_plan_pods'
const VERSIONS = 'org_plan_versions'
const POD_LINKS = 'org_plan_pod_links'
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

export async function GET(request: Request) {
  if (!(await unlocked())) return fail('locked', 401)
  const { sb, error: keyError } = db()
  if (!sb) return fail(`setup: ${keyError}`)
  // Versions: each one is a full, independent copy of the roster + chart.
  let { data: versions, error: vErr } = await sb.from(VERSIONS).select('*').order('created_at', { ascending: true })
  if (vErr) return fail(`setup: ${vErr.message}`)
  if (!versions?.length) {
    const { data: v, error } = await sb.from(VERSIONS).insert([{ name: 'Version 1' }]).select().single()
    if (error) return fail(`setup: ${error.message}`)
    await sb.from(TABLE).update({ version_id: v.id }).is('version_id', null)
    versions = [v]
  }
  const wanted = new URL(request.url).searchParams.get('version')
  const versionId = versions.some(v => v.id === wanted) ? wanted! : versions[0].id
  const { data, error } = await sb.from(TABLE).select('*').eq('version_id', versionId).order('name', { ascending: true })
  if (error) return fail(`setup: ${error.message}`)
  return NextResponse.json({ people: data, versions, versionId })
}

type Person = { id?: string; version_id?: string; name?: string; position?: string | null; team_name?: string | null; parent_id?: string | null; in_chart?: boolean; sort_order?: number | null }

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
    if (p.sort_order !== undefined) out.sort_order = p.sort_order == null ? null : Number(p.sort_order)
    return out
  }

  const vid = String(body.versionId ?? '')
  const needsVersion = ['add', 'update', 'updateMany', 'delete', 'import', 'clear'].includes(body.action)
  if (needsVersion && !vid) return fail('No version selected', 400)

  switch (body.action) {
    // ---- Versions ----
    case 'versionCreate': {
      const name = String(body.name ?? '').trim() || 'New version'
      const { data, error } = await sb.from(VERSIONS).insert([{ name }]).select().single()
      return error ? fail(error.message) : NextResponse.json({ version: data })
    }
    case 'versionRename': {
      const name = String(body.name ?? '').trim()
      if (!name) return fail('Version name is required', 400)
      const { data, error } = await sb.from(VERSIONS).update({ name }).eq('id', body.id).select().single()
      return error ? fail(error.message) : NextResponse.json({ version: data })
    }
    case 'versionDelete': {
      const { count } = await sb.from(VERSIONS).select('id', { count: 'exact', head: true })
      if ((count ?? 0) <= 1) return fail("Can't delete the only version", 400)
      const { error: e1 } = await sb.from(TABLE).delete().eq('version_id', body.id)
      if (e1) return fail(e1.message)
      const { error } = await sb.from(VERSIONS).delete().eq('id', body.id)
      return error ? fail(error.message) : NextResponse.json({ ok: true })
    }
    case 'versionClone': {
      // Copy every person in the source version with fresh ids, remapping manager links.
      const name = String(body.name ?? '').trim() || 'Copy'
      const { data: src, error: e1 } = await sb.from(TABLE).select('*').eq('version_id', body.sourceId)
      if (e1) return fail(e1.message)
      const { data: v, error: e2 } = await sb.from(VERSIONS).insert([{ name }]).select().single()
      if (e2) return fail(e2.message)
      const idMap = new Map((src ?? []).map(p => [p.id as string, randomUUID()]))
      const rows = (src ?? []).map(p => ({
        id: idMap.get(p.id)!, version_id: v.id, name: p.name, position: p.position, team_name: p.team_name,
        parent_id: p.parent_id ? idMap.get(p.parent_id) ?? null : null, in_chart: p.in_chart, sort_order: p.sort_order ?? null,
      }))
      // Insert managers before their reports so the self-reference is always satisfied.
      const byId = new Map(rows.map(r => [r.id, r]))
      const ordered: typeof rows = []
      const done = new Set<string>()
      const visit = (r: (typeof rows)[number], guard = new Set<string>()) => {
        if (done.has(r.id) || guard.has(r.id)) return
        guard.add(r.id)
        const parent = r.parent_id ? byId.get(r.parent_id) : undefined
        if (parent) visit(parent, guard)
        done.add(r.id); ordered.push(r)
      }
      rows.forEach(r => visit(r))
      if (ordered.length) {
        const { error } = await sb.from(TABLE).insert(ordered)
        if (error) { await sb.from(VERSIONS).delete().eq('id', v.id); return fail(error.message) }
      }
      return NextResponse.json({ version: v })
    }
    // ---- Everything, for the HTML viewer export ----
    case 'exportAll': {
      const [v, ppl, pods, links] = await Promise.all([
        sb.from(VERSIONS).select('*').order('created_at', { ascending: true }),
        sb.from(TABLE).select('*').order('name', { ascending: true }),
        sb.from(PODS).select('*').order('sort_order', { ascending: true, nullsFirst: false }).order('created_at', { ascending: true }),
        sb.from(POD_LINKS).select('*'),
      ])
      if (v.error) return fail(v.error.message)
      if (ppl.error) return fail(ppl.error.message)
      return NextResponse.json({ versions: v.data ?? [], people: ppl.data ?? [], pods: pods.data ?? [], links: links.data ?? [] })
    }
    // ---- Pods definition (separate table: org_plan_pods) ----
    case 'pods': {
      const { data, error } = await sb.from(PODS).select('*').order('sort_order', { ascending: true, nullsFirst: false }).order('created_at', { ascending: true })
      if (error) return fail(`pods-setup: ${error.message}`)
      const { data: links, error: lErr } = await sb.from(POD_LINKS).select('*')
      return NextResponse.json({ pods: data, links: links ?? [], linksSetup: lErr ? lErr.message : null })
    }
    case 'podAdd': {
      const name = String(body.pod?.name ?? '').trim()
      if (!name) return fail('Pod name is required', 400)
      const { data, error } = await sb.from(PODS).insert([{ name, description: String(body.pod?.description ?? '').trim() || null, sort_order: body.pod?.sort_order ?? null }]).select().single()
      return error ? fail(error.message) : NextResponse.json({ pod: data })
    }
    case 'podUpdate': {
      const patch: Record<string, unknown> = {}
      if (body.patch?.name !== undefined) { patch.name = String(body.patch.name).trim(); if (!patch.name) return fail('Pod name is required', 400) }
      if (body.patch?.description !== undefined) patch.description = String(body.patch.description ?? '').trim() || null
      if (body.patch?.sort_order !== undefined) patch.sort_order = body.patch.sort_order == null ? null : Number(body.patch.sort_order)
      const { data, error } = await sb.from(PODS).update(patch).eq('id', body.id).select().single()
      return error ? fail(error.message) : NextResponse.json({ pod: data })
    }
    case 'podLink': {
      // Links are undirected: store them with the smaller id first so A↔B is only saved once.
      const [a, b] = [String(body.a ?? ''), String(body.b ?? '')].sort()
      if (!a || !b || a === b) return fail('Pick two different pods', 400)
      const { data: existing } = await sb.from(POD_LINKS).select('*').eq('pod_a', a).eq('pod_b', b).maybeSingle()
      if (existing) return NextResponse.json({ link: existing })
      const { data, error } = await sb.from(POD_LINKS).insert([{ pod_a: a, pod_b: b }]).select().single()
      return error ? fail(error.message) : NextResponse.json({ link: data })
    }
    case 'podUnlink': {
      const { error } = await sb.from(POD_LINKS).delete().eq('id', body.id)
      return error ? fail(error.message) : NextResponse.json({ ok: true })
    }
    case 'podDelete': {
      const { error } = await sb.from(PODS).delete().eq('id', body.id)
      return error ? fail(error.message) : NextResponse.json({ ok: true })
    }
    case 'add': {
      const row = clean(body.person ?? {})
      if (!row.name) return fail('Name is required', 400)
      const { data, error } = await sb.from(TABLE).insert([{ ...row, version_id: vid }]).select().single()
      return error ? fail(error.message) : NextResponse.json({ person: data })
    }
    case 'update': {
      const { data, error } = await sb.from(TABLE).update(clean(body.patch ?? {})).eq('id', body.id).eq('version_id', vid).select().single()
      return error ? fail(error.message) : NextResponse.json({ person: data })
    }
    case 'updateMany': {
      // [{ id, patch }] — used for reattaching reports / removing branches
      for (const u of (body.updates ?? []) as { id: string; patch: Person }[]) {
        const { error } = await sb.from(TABLE).update(clean(u.patch)).eq('id', u.id).eq('version_id', vid)
        if (error) return fail(error.message)
      }
      return NextResponse.json({ ok: true })
    }
    case 'delete': {
      const { error } = await sb.from(TABLE).delete().eq('id', body.id).eq('version_id', vid)
      return error ? fail(error.message) : NextResponse.json({ ok: true })
    }
    case 'import': {
      // [{ name, position }] — new names are added, existing names (case-insensitive) get their position updated
      const rows = ((body.rows ?? []) as Person[]).map(clean).filter(r => r.name)
      const { data: existing, error: e1 } = await sb.from(TABLE).select('id,name').eq('version_id', vid)
      if (e1) return fail(e1.message)
      const byName = new Map((existing ?? []).map(r => [String(r.name).toLowerCase(), r.id as string]))
      const inserts: Person[] = []
      let updated = 0
      for (const r of rows) {
        const id = byName.get(r.name!.toLowerCase())
        if (id) {
          if (r.position) { await sb.from(TABLE).update({ position: r.position }).eq('id', id); updated++ }
        } else {
          inserts.push({ ...r, version_id: vid })
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
      const { error } = await sb.from(TABLE).delete().eq('version_id', vid)
      return error ? fail(error.message) : NextResponse.json({ ok: true })
    }
    default:
      return fail('Unknown action', 400)
  }
}
