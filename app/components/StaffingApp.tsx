'use client'

import { useState, useEffect } from 'react'
import { supabase } from '@/lib/supabase'
import { Document, Packer, Paragraph, Table, TableRow, TableCell, TextRun, HeadingLevel, WidthType, BorderStyle } from 'docx'

type Project = {
  id: string
  name: string
  customer_codename: string | null
  status: string
  start_date: string | null
  end_date: string | null
  flagged: boolean
}

type Staff = {
  id: string
  name: string
  position: string | null
  ooo: boolean
  ooo_return_date: string | null
  flexed: boolean
  onboarding: boolean
  flex_notes: string | null
}

type Assignment = {
  id: string
  project_id: string
  staff_id: string
  assignment_role: string | null
}

type Scenario = {
  id: string
  name: string
  created_at: string
}

type ScenarioAssignment = {
  id: string
  scenario_id: string
  project_id: string
  staff_id: string
  assignment_role: string | null
}

type Milestone = {
  id: string
  project_id: string
  title: string
  priority: string
  done: boolean
  due_date: string | null
  created_at: string
}

const PRIORITIES = ['P0', 'P1', 'P2']
const priorityRank = (p: string) => { const i = PRIORITIES.indexOf(p); return i === -1 ? 99 : i }
const priorityColor = (p: string) =>
  p === 'P0' ? 'bg-red-500/10 text-red-400' :
  p === 'P1' ? 'bg-amber-500/10 text-amber-400' :
  'bg-sky-500/10 text-sky-400'

const VALID_POSITIONS = ['SPA', 'SPL I', 'SPL II', 'Manager, Delivery', 'Senior SPL', 'Head of Delivery', 'GenAI Consultant']

const STATUS_OPTIONS: { value: string; label: string }[] = [
  { value: 'active', label: 'Active' },
  { value: 'starting-soon', label: 'Starting Soon' },
  { value: 'paused', label: 'Paused' },
  { value: 'on-hold', label: 'On Hold' },
  { value: 'completed', label: 'Completed' },
]
const STATUS_ORDER = STATUS_OPTIONS.map(s => s.value)
const statusRank = (status: string) => {
  const i = STATUS_ORDER.indexOf(status)
  return i === -1 ? 99 : i
}

function parseCSVLine(line: string): string[] {
  const cols: string[] = []
  let cur = ''
  let inQuotes = false
  for (let i = 0; i < line.length; i++) {
    const ch = line[i]
    if (ch === '"') { inQuotes = !inQuotes }
    else if (ch === ',' && !inQuotes) { cols.push(cur.trim()); cur = '' }
    else { cur += ch }
  }
  cols.push(cur.trim())
  return cols
}

export default function StaffingApp() {
  const [tab, setTab] = useState<'projects' | 'staff' | 'assignments' | 'dashboard' | 'scenarios' | 'milestones'>('dashboard')
  const [projects, setProjects] = useState<Project[]>([])
  const [staff, setStaff] = useState<Staff[]>([])
  const [assignments, setAssignments] = useState<Assignment[]>([])
  const [scenarios, setScenarios] = useState<Scenario[]>([])
  const [scenarioAssignments, setScenarioAssignments] = useState<ScenarioAssignment[]>([])
  const [milestones, setMilestones] = useState<Milestone[]>([])
  const [milestoneDrafts, setMilestoneDrafts] = useState<Record<string, { title: string; priority: string; due_date: string }>>({})
  const [hideEmptyMilestoneProjects, setHideEmptyMilestoneProjects] = useState(false)
  const [addingMilestoneProjectId, setAddingMilestoneProjectId] = useState<string | null>(null)
  const [editingMilestoneId, setEditingMilestoneId] = useState<string | null>(null)
  const [editMilestone, setEditMilestone] = useState({ title: '', due_date: '' })
  const [loadError, setLoadError] = useState(false)
  const [selectedScenarioId, setSelectedScenarioId] = useState<string | null>(null)
  const [newScenarioName, setNewScenarioName] = useState('')
  const [newScenarioCopyCurrent, setNewScenarioCopyCurrent] = useState(true)
  const [scenarioAssign, setScenarioAssign] = useState({ project_id: '', staff_id: '', assignment_role: '' })
  const [confirmApplyScenario, setConfirmApplyScenario] = useState<string | null>(null)
  const [dragOverScenarioProjectId, setDragOverScenarioProjectId] = useState<string | null>(null)
  const [pendingScenarioDrop, setPendingScenarioDrop] = useState<{ staffId: string; projectId: string } | null>(null)
  const [addingToScenarioProjectId, setAddingToScenarioProjectId] = useState<string | null>(null)
  const [scenarioQuickAdd, setScenarioQuickAdd] = useState({ staff_id: '', assignment_role: '' })
  const [loading, setLoading] = useState(true)
  const [theme, setTheme] = useState<'dark' | 'light' | 'system'>('dark')

  const [newProject, setNewProject] = useState({ name: '', customer_codename: '', status: 'active', duration_weeks: '' })
  const [newStaff, setNewStaff] = useState({ name: '', position: '', ooo: false, ooo_return_date: '' })
  const [newAssignment, setNewAssignment] = useState({ project_id: '', staff_id: '', assignment_role: '' })

  const [editingProjectId, setEditingProjectId] = useState<string | null>(null)
  const [editProject, setEditProject] = useState({ name: '', customer_codename: '', status: 'active', duration_weeks: '' })

  const [editingStaffId, setEditingStaffId] = useState<string | null>(null)
  const [editStaff, setEditStaff] = useState({ name: '', position: '', ooo: false, ooo_return_date: '', flex_notes: '' })
  const [inlineAssignment, setInlineAssignment] = useState({ project_id: '', assignment_role: '' })
  const [csvErrors, setCsvErrors] = useState<string[]>([])
  const [staffSort, setStaffSort] = useState<'default' | 'az' | 'za'>('az')
  const [hideAssigned, setHideAssigned] = useState(false)
  const [hideAssignedInDropdown, setHideAssignedInDropdown] = useState(false)
  const [projectSort, setProjectSort] = useState<{ col: string; dir: 'az' | 'za' }>({ col: 'status', dir: 'az' })
  const [showSupervisorsInChart, setShowSupervisorsInChart] = useState(false)
  const [visibleStatuses, setVisibleStatuses] = useState<Record<string, boolean>>({
    'active': true, 'starting-soon': true, 'paused': true, 'on-hold': false, 'completed': false,
  })
  const [statusFilterOpen, setStatusFilterOpen] = useState(false)
  const [addingToProjectId, setAddingToProjectId] = useState<string | null>(null)
  const [showAddProjectInline, setShowAddProjectInline] = useState(false)
  const [quickAdd, setQuickAdd] = useState({ staff_id: '', assignment_role: '' })
  const [draggedStaffId, setDraggedStaffId] = useState<string | null>(null)
  const [draggedAssignmentId, setDraggedAssignmentId] = useState<string | null>(null)
  const [dragOverProjectId, setDragOverProjectId] = useState<string | null>(null)
  const [dragOverFlexed, setDragOverFlexed] = useState(false)
  const [dragOverUnassigned, setDragOverUnassigned] = useState(false)
  const [dragOverOOO, setDragOverOOO] = useState(false)
  const [dragOverOnboarding, setDragOverOnboarding] = useState(false)
  const [pendingDrop, setPendingDrop] = useState<{ staffId: string; projectId: string } | null>(null)
  const [pendingMove, setPendingMove] = useState<{ assignmentId: string; projectId: string } | null>(null)
  const [confirmFreeStaff, setConfirmFreeStaff] = useState<{ projectId: string; count: number } | null>(null)
  const [pendingOOO, setPendingOOO] = useState<{ staffId: string; assignmentId: string | null } | null>(null)
  const [oooDate, setOooDate] = useState('')
  const [dropRole, setDropRole] = useState('')

  useEffect(() => {
    const saved = localStorage.getItem('theme') as 'dark' | 'light' | 'system' | null
    if (saved) setTheme(saved)
  }, [])

  useEffect(() => {
    localStorage.setItem('theme', theme)
    const root = document.documentElement
    const applyDark = () => root.classList.remove('light')
    const applyLight = () => root.classList.add('light')
    if (theme === 'light') { applyLight() }
    else if (theme === 'dark') { applyDark() }
    else {
      const mq = window.matchMedia('(prefers-color-scheme: light)')
      mq.matches ? applyLight() : applyDark()
      const handler = (e: MediaQueryListEvent) => e.matches ? applyLight() : applyDark()
      mq.addEventListener('change', handler)
      return () => mq.removeEventListener('change', handler)
    }
  }, [theme])

  useEffect(() => { loadAll() }, [])

  async function loadAll() {
    setLoading(true)
    setLoadError(false)
    try {
      const timeout = new Promise<never>((_, reject) => setTimeout(() => reject(new Error('Request timed out')), 15000))
      const queries = Promise.all([
        supabase.from('projects').select('*').order('created_at', { ascending: false }),
        supabase.from('staff').select('*').order('created_at', { ascending: false }),
        supabase.from('assignments').select('*'),
        supabase.from('scenarios').select('*').order('created_at', { ascending: true }),
        supabase.from('scenario_assignments').select('*'),
        supabase.from('milestones').select('*').order('created_at', { ascending: true }),
      ])
      const [p, s, a, sc, sa, ms] = await Promise.race([queries, timeout])
      // Core tables must load; scenario tables are optional (may not exist yet)
      if (p.error || s.error || a.error) throw (p.error || s.error || a.error)
      if (p.data) setProjects(p.data)
      if (s.data) {
        // Auto-clear OOO for anyone whose return date has arrived (date <= today)
        const today = fmt(new Date())
        const expired = s.data.filter((m: Staff) => m.ooo && m.ooo_return_date && m.ooo_return_date <= today)
        if (expired.length > 0) {
          await supabase.from('staff').update({ ooo: false, ooo_return_date: null }).in('id', expired.map((m: Staff) => m.id))
          s.data = s.data.map((m: Staff) => expired.some(e => e.id === m.id) ? { ...m, ooo: false, ooo_return_date: null } : m)
        }
        setStaff(s.data)
      }
      if (a.data) setAssignments(a.data)
      if (sc.data) setScenarios(sc.data)
      if (sa.data) setScenarioAssignments(sa.data)
      if (ms.data) setMilestones(ms.data)
    } catch {
      setLoadError(true)
    } finally {
      setLoading(false)
    }
  }

  const fmt = (d: Date) => `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`

  async function addProject() {
    if (!newProject.name.trim()) return
    const start = new Date()
    const end = new Date()
    end.setDate(end.getDate() + (Number(newProject.duration_weeks) || 0) * 7)
    const { data } = await supabase.from('projects').insert([{
      name: newProject.name.trim(),
      customer_codename: newProject.customer_codename.trim() || null,
      status: newProject.status,
      start_date: fmt(start),
      end_date: newProject.duration_weeks ? fmt(end) : null,
    }]).select().single()
    if (data) {
      setProjects([data, ...projects])
      setNewProject({ name: '', customer_codename: '', status: 'active', duration_weeks: '' })
    }
  }

  function startEditProject(p: Project) {
    const weeksRemaining = p.end_date
      ? String(Math.round((new Date(p.end_date).getTime() - new Date(p.start_date ?? '').getTime()) / (7 * 24 * 60 * 60 * 1000)))
      : ''
    setEditProject({ name: p.name, customer_codename: p.customer_codename ?? '', status: p.status, duration_weeks: weeksRemaining })
    setEditingProjectId(p.id)
  }

  async function saveProject(id: string) {
    if (!editProject.name.trim()) return
    const start = projects.find(p => p.id === id)?.start_date ?? fmt(new Date())
    const end = new Date(start)
    end.setDate(end.getDate() + (Number(editProject.duration_weeks) || 0) * 7)
    const updates = {
      name: editProject.name.trim(),
      customer_codename: editProject.customer_codename.trim() || null,
      status: editProject.status,
      end_date: editProject.duration_weeks ? fmt(end) : null,
    }
    const wasCompleted = projects.find(p => p.id === id)?.status === 'completed'
    const { data } = await supabase.from('projects').update(updates).eq('id', id).select().single()
    if (data) {
      setProjects(projects.map(p => p.id === id ? data : p))
      setEditingProjectId(null)
      // Newly completed project with staff → ask whether to free them up
      const assignedCount = assignments.filter(a => a.project_id === id).length
      if (data.status === 'completed' && !wasCompleted && assignedCount > 0) {
        setConfirmFreeStaff({ projectId: id, count: assignedCount })
      }
    }
  }

  async function updateProjectStatus(id: string, status: string) {
    const prevStatus = projects.find(p => p.id === id)?.status
    const { data } = await supabase.from('projects').update({ status }).eq('id', id).select().single()
    if (data) {
      setProjects(projects.map(p => p.id === id ? data : p))
      const assignedCount = assignments.filter(a => a.project_id === id).length
      if (status === 'completed' && prevStatus !== 'completed' && assignedCount > 0) {
        setConfirmFreeStaff({ projectId: id, count: assignedCount })
      }
    }
  }

  async function toggleProjectFlag(id: string, flagged: boolean) {
    const { data } = await supabase.from('projects').update({ flagged }).eq('id', id).select().single()
    if (data) setProjects(projects.map(p => p.id === id ? data : p))
  }

  async function deleteProject(id: string) {
    await supabase.from('projects').delete().eq('id', id)
    setProjects(projects.filter(p => p.id !== id))
    setAssignments(assignments.filter(a => a.project_id !== id))
  }

  async function addStaff() {
    if (!newStaff.name.trim()) return
    const { data } = await supabase.from('staff').insert([{
      name: newStaff.name.trim(),
      position: newStaff.position || null,
      ooo: newStaff.ooo,
      ooo_return_date: newStaff.ooo_return_date || null,
    }]).select().single()
    if (data) {
      setStaff([data, ...staff])
      setNewStaff({ name: '', position: '', ooo: false, ooo_return_date: '' })
    }
  }

  function startEditStaff(s: Staff) {
    setEditStaff({ name: s.name, position: s.position ?? '', ooo: s.ooo, ooo_return_date: s.ooo_return_date ?? '', flex_notes: s.flex_notes ?? '' })
    setEditingStaffId(s.id)
  }

  async function saveStaff(id: string) {
    if (!editStaff.name.trim()) return
    const updates = { name: editStaff.name.trim(), position: editStaff.position || null, ooo: editStaff.ooo, ooo_return_date: editStaff.ooo_return_date || null, flex_notes: editStaff.flex_notes.trim() || null }
    const { data } = await supabase.from('staff').update(updates).eq('id', id).select().single()
    if (data) {
      setStaff(staff.map(s => s.id === id ? data : s))
      setEditingStaffId(null)
    }
  }

  async function deleteStaff(id: string) {
    await supabase.from('staff').delete().eq('id', id)
    setStaff(staff.filter(s => s.id !== id))
    setAssignments(assignments.filter(a => a.staff_id !== id))
  }

  async function addAssignment() {
    if (!newAssignment.project_id || !newAssignment.staff_id) return
    const alreadyAssigned = assignments.some(
      a => a.project_id === newAssignment.project_id && a.staff_id === newAssignment.staff_id
    )
    if (alreadyAssigned) return
    const { data } = await supabase.from('assignments').insert([newAssignment]).select().single()
    if (data) {
      setAssignments([...assignments, data])
      setNewAssignment({ project_id: '', staff_id: '', assignment_role: '' })
    }
  }

  async function importStaffCSV(e: React.ChangeEvent<HTMLInputElement>) {
    const file = e.target.files?.[0]
    if (!file) return
    setCsvErrors([])
    const text = await file.text()
    const lines = text.trim().split('\n')
    const headers = parseCSVLine(lines[0]).map(h => h.toUpperCase())
    const nameIdx = headers.indexOf('NAME')
    const positionIdx = headers.indexOf('POSITION')
    if (nameIdx === -1) { setCsvErrors(['CSV must have a NAME column.']); return }

    const rows = lines.slice(1).map((line, i) => {
      const cols = parseCSVLine(line)
      return { line: i + 2, name: cols[nameIdx], position: positionIdx !== -1 ? cols[positionIdx] || null : null }
    }).filter(r => r.name)

    const invalid = rows.filter(r => r.position && !VALID_POSITIONS.includes(r.position))
    if (invalid.length > 0) {
      setCsvErrors(invalid.map(r => `Row ${r.line} — "${r.name}" has invalid position: "${r.position}"`))
      e.target.value = ''
      return
    }

    for (const row of rows) {
      const existing = staff.find(s => s.name.toLowerCase() === row.name.toLowerCase())
      if (existing) {
        const { data } = await supabase.from('staff').update({ position: row.position }).eq('id', existing.id).select().single()
        if (data) setStaff(prev => prev.map(s => s.id === existing.id ? data : s))
      } else {
        const { data } = await supabase.from('staff').insert([{ name: row.name, position: row.position }]).select().single()
        if (data) setStaff(prev => [data, ...prev])
      }
    }
    e.target.value = ''
  }

  function exportAssignmentsHTML() {
    const rows = [...projects].filter(p => p.status === 'active').sort((a, b) => a.name.localeCompare(b.name)).map(p => {
      const pa = assignments.filter(a => a.project_id === p.id)
      if (pa.length === 0) return `<h2>${p.name}</h2><p><em>No staff assigned.</em></p>`
      const rows = pa.map(a => {
        const member = staff.find(s => s.id === a.staff_id)
        return `<tr><td>${member?.name ?? 'Unknown'}</td><td>${member?.position ?? '—'}</td><td>${a.assignment_role ?? '—'}</td></tr>`
      }).join('')
      return `<h2>${p.name}${p.customer_codename ? ` <small>(${p.customer_codename})</small>` : ''}</h2>
<p>Status: ${p.status}${p.end_date ? ` &nbsp;|&nbsp; Ends: ${p.end_date}` : ''}</p>
<table border="1" cellpadding="6" cellspacing="0" style="border-collapse:collapse;width:100%">
  <thead><tr><th>Name</th><th>Position</th><th>Role</th></tr></thead>
  <tbody>${rows}</tbody>
</table>`
    }).join('<br/>')

    const html = `<!DOCTYPE html><html><head><meta charset="utf-8"><title>Staffing Assignments</title>
<style>body{font-family:Arial,sans-serif;max-width:900px;margin:40px auto;color:#111}h1{margin-bottom:8px}h2{margin-top:32px;margin-bottom:4px}table{width:100%;margin-top:8px}th{background:#f0f0f0;text-align:left}td,th{padding:6px 10px}</style>
</head><body>
<h1>Code Pod Staffing — Assignments</h1>
<p style="color:#666">Exported ${new Date().toLocaleDateString()}</p>
${rows}
</body></html>`

    const blob = new Blob([html], { type: 'text/html' })
    const url = URL.createObjectURL(blob)
    const a = document.createElement('a')
    a.href = url
    a.download = `staffing-assignments-${new Date().toISOString().split('T')[0]}.html`
    a.click()
    URL.revokeObjectURL(url)
  }

  function buildMilestonesHTML() {
    const esc = (t: string) => t.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    const today = fmt(new Date())
    const activeProjects = [...projects].filter(p => p.status === 'active').sort((a, b) => a.name.localeCompare(b.name))
    const activeIds = new Set(activeProjects.map(p => p.id))

    const openP0 = milestones
      .filter(m => !m.done && m.priority === 'P0' && activeIds.has(m.project_id))
      .sort((a, b) => (a.due_date ?? '9999').localeCompare(b.due_date ?? '9999'))
    const p0Section = openP0.length === 0 ? '' : `
<h2 style="color:#c0392b">Open P0 — Needs Attention (${openP0.length})</h2>
<ul>${openP0.map(m => {
      const proj = projects.find(p => p.id === m.project_id)
      const overdue = m.due_date && m.due_date < today
      return `<li><strong>${esc(m.title)}</strong> — <em>${esc(proj?.name ?? '')}</em>${m.due_date ? ` <span style="color:${overdue ? '#c0392b' : '#666'}">(${overdue ? 'overdue · ' : 'due '}${m.due_date})</span>` : ''}</li>`
    }).join('')}</ul>`

    const sections = activeProjects.map(p => {
      const ms = [...milestones.filter(m => m.project_id === p.id)].sort((a, b) =>
        (a.done === b.done ? 0 : a.done ? 1 : -1) || priorityRank(a.priority) - priorityRank(b.priority) || a.created_at.localeCompare(b.created_at)
      )
      if (ms.length === 0) return ''
      const done = ms.filter(m => m.done).length
      const rows = ms.map(m => {
        const overdue = !m.done && m.due_date && m.due_date < today
        return `<tr>
  <td style="text-align:center">${m.done ? '✓' : '☐'}</td>
  <td>${m.priority}</td>
  <td${m.done ? ' style="text-decoration:line-through;color:#999"' : ''}>${esc(m.title)}</td>
  <td style="color:${overdue ? '#c0392b' : '#666'}">${m.due_date ? `${overdue ? '⚠ ' : ''}${m.due_date}` : '—'}</td>
</tr>`
      }).join('')
      return `<h2>${esc(p.name)} <small style="color:#888;font-weight:normal">(${done}/${ms.length} done)</small></h2>
<table border="1" cellpadding="6" cellspacing="0" style="border-collapse:collapse;width:100%">
  <thead><tr><th>Done</th><th>Priority</th><th>Milestone</th><th>Due</th></tr></thead>
  <tbody>${rows}</tbody>
</table>`
    }).filter(Boolean).join('<br/>')

    const html = `<!DOCTYPE html><html><head><meta charset="utf-8"><title>Milestones Summary</title>
<style>body{font-family:Arial,sans-serif;max-width:900px;margin:40px auto;color:#111}h1{margin-bottom:8px}h2{margin-top:32px;margin-bottom:4px}table{width:100%;margin-top:8px}th{background:#f0f0f0;text-align:left}td,th{padding:6px 10px}ul{line-height:1.6}</style>
</head><body>
<h1>Code Pod — Milestones Summary</h1>
<p style="color:#666">Exported ${new Date().toLocaleDateString()} · Active projects only</p>
${p0Section}
${sections || '<p><em>No milestones yet.</em></p>'}
</body></html>`
    return html
  }

  function exportMilestonesHTML() {
    const blob = new Blob([buildMilestonesHTML()], { type: 'text/html' })
    const url = URL.createObjectURL(blob)
    const a = document.createElement('a')
    a.href = url
    a.download = `milestones-${new Date().toISOString().split('T')[0]}.html`
    a.click()
    URL.revokeObjectURL(url)
  }

  function exportMilestonesPDF() {
    const iframe = document.createElement('iframe')
    iframe.style.position = 'fixed'
    iframe.style.right = '0'
    iframe.style.bottom = '0'
    iframe.style.width = '0'
    iframe.style.height = '0'
    iframe.style.border = '0'
    document.body.appendChild(iframe)
    const doc = iframe.contentWindow?.document
    if (!doc) { document.body.removeChild(iframe); return }
    doc.open()
    doc.write(buildMilestonesHTML())
    doc.close()
    iframe.contentWindow?.focus()
    setTimeout(() => {
      iframe.contentWindow?.print()
      setTimeout(() => document.body.removeChild(iframe), 1000)
    }, 300)
  }

  async function exportAssignmentsDOCX() {
    const cellBorder = { style: BorderStyle.SINGLE, size: 4, color: 'DDDDDD' }
    const border = { top: cellBorder, bottom: cellBorder, left: cellBorder, right: cellBorder }
    // Page usable width ~9360 twips (letter, 1" margins each side)
    const colWidths = [4500, 2860, 2000] // Name, Position, Role

    const makeRow = (vals: string[], isHeader = false) =>
      new TableRow({
        children: vals.map((val, i) =>
          new TableCell({
            children: [new Paragraph({ children: [new TextRun({ text: val, bold: isHeader, size: 20 })] })],
            borders: border,
            shading: { fill: isHeader ? 'EEEEEE' : 'FFFFFF' },
            width: { size: colWidths[i], type: WidthType.DXA },
          })
        ),
        tableHeader: isHeader,
      })

    const children: (Paragraph | Table)[] = [
      new Paragraph({ text: 'Code Pod Staffing — Assignments', heading: HeadingLevel.TITLE }),
      new Paragraph({ text: `Exported ${new Date().toLocaleDateString()}`, spacing: { after: 400 } }),
    ]

    for (const p of [...projects].filter(p => p.status === 'active').sort((a, b) => a.name.localeCompare(b.name))) {
      const pa = assignments.filter(a => a.project_id === p.id)
      children.push(new Paragraph({ text: p.name + (p.customer_codename ? ` (${p.customer_codename})` : ''), heading: HeadingLevel.HEADING_2, spacing: { before: 400 } }))
      children.push(new Paragraph({ children: [new TextRun({ text: `Status: ${p.status}${p.end_date ? `   |   Ends: ${p.end_date}` : ''}`, color: '666666', size: 18 })], spacing: { after: 160 } }))

      if (pa.length === 0) {
        children.push(new Paragraph({ text: 'No staff assigned.', spacing: { after: 200 } }))
        continue
      }

      const dataRows = pa.map(a => {
        const member = staff.find(s => s.id === a.staff_id)
        return makeRow([member?.name ?? 'Unknown', member?.position ?? '—', a.assignment_role ?? '—'])
      })

      children.push(new Table({
        rows: [makeRow(['Name', 'Position', 'Role'], true), ...dataRows],
        width: { size: 9360, type: WidthType.DXA },
        columnWidths: colWidths,
      }))
      children.push(new Paragraph({ text: '' }))
    }

    const doc = new Document({ sections: [{ children }] })
    const blob = await Packer.toBlob(doc)
    const url = URL.createObjectURL(blob)
    const a = document.createElement('a')
    a.href = url
    a.download = `staffing-assignments-${new Date().toISOString().split('T')[0]}.docx`
    a.click()
    URL.revokeObjectURL(url)
  }

  async function setStaffZone(id: string, zone: 'unassigned' | 'flexed' | 'ooo' | 'onboarding') {
    const { data } = await supabase.from('staff').update({ flexed: zone === 'flexed', ooo: zone === 'ooo', onboarding: zone === 'onboarding' }).eq('id', id).select().single()
    if (data) setStaff(prev => prev.map(s => s.id === id ? data : s))
  }

  async function updateAssignmentRole(id: string, role: string) {
    const { data } = await supabase.from('assignments').update({ assignment_role: role || null }).eq('id', id).select().single()
    if (data) setAssignments(prev => prev.map(a => a.id === id ? data : a))
  }

  async function removeAssignment(id: string) {
    await supabase.from('assignments').delete().eq('id', id)
    setAssignments(assignments.filter(a => a.id !== id))
  }

  async function addScenario() {
    if (!newScenarioName.trim()) return
    const { data } = await supabase.from('scenarios').insert([{ name: newScenarioName.trim() }]).select().single()
    if (data) {
      setScenarios([...scenarios, data])
      setSelectedScenarioId(data.id)
      // Optionally seed from current live assignments
      if (newScenarioCopyCurrent && assignments.length > 0) {
        const rows = assignments.map(a => ({ scenario_id: data.id, project_id: a.project_id, staff_id: a.staff_id, assignment_role: a.assignment_role }))
        const { data: seeded } = await supabase.from('scenario_assignments').insert(rows).select()
        if (seeded) setScenarioAssignments(prev => [...prev, ...seeded])
      }
      setNewScenarioName('')
    }
  }

  async function deleteScenario(id: string) {
    await supabase.from('scenarios').delete().eq('id', id)
    setScenarios(scenarios.filter(s => s.id !== id))
    setScenarioAssignments(prev => prev.filter(sa => sa.scenario_id !== id))
    if (selectedScenarioId === id) setSelectedScenarioId(null)
  }

  async function addScenarioAssignment() {
    if (!selectedScenarioId || !scenarioAssign.project_id || !scenarioAssign.staff_id) return
    const exists = scenarioAssignments.some(sa => sa.scenario_id === selectedScenarioId && sa.project_id === scenarioAssign.project_id && sa.staff_id === scenarioAssign.staff_id)
    if (exists) return
    const { data } = await supabase.from('scenario_assignments').insert([{
      scenario_id: selectedScenarioId,
      project_id: scenarioAssign.project_id,
      staff_id: scenarioAssign.staff_id,
      assignment_role: scenarioAssign.assignment_role || null,
    }]).select().single()
    if (data) {
      setScenarioAssignments([...scenarioAssignments, data])
      setScenarioAssign({ project_id: '', staff_id: '', assignment_role: '' })
    }
  }

  async function removeScenarioAssignment(id: string) {
    await supabase.from('scenario_assignments').delete().eq('id', id)
    setScenarioAssignments(scenarioAssignments.filter(sa => sa.id !== id))
  }

  async function addScenarioAssignmentDirect(scenarioId: string, projectId: string, staffId: string, role: string) {
    const exists = scenarioAssignments.some(sa => sa.scenario_id === scenarioId && sa.project_id === projectId && sa.staff_id === staffId)
    if (exists) return
    const { data } = await supabase.from('scenario_assignments').insert([{
      scenario_id: scenarioId, project_id: projectId, staff_id: staffId, assignment_role: role || null,
    }]).select().single()
    if (data) setScenarioAssignments(prev => [...prev, data])
  }

  async function moveScenarioAssignment(id: string, projectId: string) {
    const { data } = await supabase.from('scenario_assignments').update({ project_id: projectId }).eq('id', id).select().single()
    if (data) setScenarioAssignments(prev => prev.map(sa => sa.id === id ? data : sa))
  }

  async function updateScenarioAssignmentRole(id: string, role: string) {
    const { data } = await supabase.from('scenario_assignments').update({ assignment_role: role || null }).eq('id', id).select().single()
    if (data) setScenarioAssignments(prev => prev.map(sa => sa.id === id ? data : sa))
  }

  async function applyScenarioToLive(scenarioId: string) {
    // Replace all live assignments with this scenario's assignments
    await supabase.from('assignments').delete().neq('id', '00000000-0000-0000-0000-000000000000')
    const sas = scenarioAssignments.filter(sa => sa.scenario_id === scenarioId)
    let inserted: Assignment[] = []
    if (sas.length > 0) {
      const rows = sas.map(sa => ({ project_id: sa.project_id, staff_id: sa.staff_id, assignment_role: sa.assignment_role }))
      const { data } = await supabase.from('assignments').insert(rows).select()
      if (data) inserted = data
    }
    setAssignments(inserted)
    setConfirmApplyScenario(null)
  }

  function getMilestoneDraft(projectId: string) {
    return milestoneDrafts[projectId] ?? { title: '', priority: 'P1', due_date: '' }
  }
  function setMilestoneDraft(projectId: string, partial: Partial<{ title: string; priority: string; due_date: string }>) {
    setMilestoneDrafts(prev => ({ ...prev, [projectId]: { ...getMilestoneDraft(projectId), ...partial } }))
  }

  async function addMilestone(projectId: string) {
    const draft = getMilestoneDraft(projectId)
    if (!draft.title.trim()) return
    const { data } = await supabase.from('milestones').insert([{
      project_id: projectId,
      title: draft.title.trim(),
      priority: draft.priority || 'P1',
      due_date: draft.due_date || null,
    }]).select().single()
    if (data) {
      setMilestones([...milestones, data])
      setMilestoneDrafts(prev => ({ ...prev, [projectId]: { title: '', priority: draft.priority, due_date: '' } }))
      setAddingMilestoneProjectId(null)
    }
  }

  async function toggleMilestone(id: string, done: boolean) {
    const { data } = await supabase.from('milestones').update({ done }).eq('id', id).select().single()
    if (data) setMilestones(prev => prev.map(m => m.id === id ? data : m))
  }

  async function updateMilestonePriority(id: string, priority: string) {
    const { data } = await supabase.from('milestones').update({ priority }).eq('id', id).select().single()
    if (data) setMilestones(prev => prev.map(m => m.id === id ? data : m))
  }

  async function deleteMilestone(id: string) {
    await supabase.from('milestones').delete().eq('id', id)
    setMilestones(milestones.filter(m => m.id !== id))
  }

  function startEditMilestone(m: Milestone) {
    setEditMilestone({ title: m.title, due_date: m.due_date ?? '' })
    setEditingMilestoneId(m.id)
  }

  async function saveMilestone(id: string) {
    if (!editMilestone.title.trim()) return
    const { data } = await supabase.from('milestones').update({ title: editMilestone.title.trim(), due_date: editMilestone.due_date || null }).eq('id', id).select().single()
    if (data) {
      setMilestones(prev => prev.map(m => m.id === id ? data : m))
      setEditingMilestoneId(null)
    }
  }

  function sortedList<T extends { name: string }>(list: T[], sort: 'default' | 'az' | 'za') {
    if (sort === 'az') return [...list].sort((a, b) => a.name.localeCompare(b.name))
    if (sort === 'za') return [...list].sort((a, b) => b.name.localeCompare(a.name))
    return list
  }

  function nextSort(current: 'default' | 'az' | 'za') {
    return current === 'default' ? 'az' : current === 'az' ? 'za' : 'default'
  }

  function sortLabel(sort: 'default' | 'az' | 'za') {
    return sort === 'az' ? 'A→Z' : sort === 'za' ? 'Z→A' : 'Sort'
  }

  function toggleProjectSort(col: string) {
    setProjectSort(prev => prev.col === col ? { col, dir: prev.dir === 'az' ? 'za' : 'az' } : { col, dir: 'az' })
  }

  function sortedProjects(list: Project[]) {
    const { col, dir } = projectSort
    const mult = dir === 'az' ? 1 : -1
    return [...list].sort((a, b) => {
      let valA: string | number = ''
      let valB: string | number = ''
      if (col === 'name')     { valA = a.name; valB = b.name }
      else if (col === 'customer') { valA = a.customer_codename ?? ''; valB = b.customer_codename ?? '' }
      else if (col === 'status')   { valA = statusRank(a.status); valB = statusRank(b.status) }
      else if (col === 'start')    { valA = a.start_date ?? ''; valB = b.start_date ?? '' }
      else if (col === 'end')      { valA = a.end_date ?? ''; valB = b.end_date ?? '' }
      else if (col === 'duration') {
        valA = a.start_date && a.end_date ? (new Date(a.end_date).getTime() - new Date(a.start_date).getTime()) : -1
        valB = b.start_date && b.end_date ? (new Date(b.end_date).getTime() - new Date(b.start_date).getTime()) : -1
      }
      else if (col === 'staff') {
        valA = assignments.filter(x => x.project_id === a.id).length
        valB = assignments.filter(x => x.project_id === b.id).length
      }
      if (valA < valB) return -1 * mult
      if (valA > valB) return 1 * mult
      return a.name.localeCompare(b.name)
    })
  }

  function colSortLabel(col: string) {
    if (projectSort.col !== col) return '↕'
    return projectSort.dir === 'az' ? '↑' : '↓'
  }

  function statusFilterDropdown() {
    const shownCount = STATUS_OPTIONS.filter(o => visibleStatuses[o.value]).length
    return (
      <div className="relative">
        <button
          onClick={() => setStatusFilterOpen(o => !o)}
          className="text-xs px-3 py-1.5 rounded-lg border border-gray-700 text-gray-400 hover:text-gray-200 transition-colors flex items-center gap-1.5"
        >
          Statuses <span className="text-[10px] bg-gray-800 rounded px-1.5 py-0.5">{shownCount}/{STATUS_OPTIONS.length}</span>
        </button>
        {statusFilterOpen && (
          <>
            <div className="fixed inset-0 z-10" onClick={() => setStatusFilterOpen(false)} />
            <div className="absolute right-0 mt-1 z-20 bg-gray-900 border border-gray-700 rounded-lg p-2 shadow-xl w-44">
              {STATUS_OPTIONS.map(o => (
                <label key={o.value} className="flex items-center gap-2 px-2 py-1.5 rounded hover:bg-gray-800 cursor-pointer text-sm text-gray-300">
                  <input
                    type="checkbox"
                    checked={!!visibleStatuses[o.value]}
                    onChange={e => setVisibleStatuses(prev => ({ ...prev, [o.value]: e.target.checked }))}
                    className="accent-[#193a29] w-4 h-4"
                  />
                  {o.label}
                </label>
              ))}
            </div>
          </>
        )}
      </div>
    )
  }

  const roleColors = [
    'bg-violet-500/10 text-violet-400',
    'bg-blue-500/10 text-blue-400',
    'bg-emerald-500/10 text-emerald-400',
    'bg-amber-500/10 text-amber-400',
    'bg-rose-500/10 text-rose-400',
    'bg-cyan-500/10 text-cyan-400',
    'bg-orange-500/10 text-orange-400',
    'bg-pink-500/10 text-pink-400',
  ]

  function roleColor(role: string) {
    const hash = role.split('').reduce((acc, c) => acc + c.charCodeAt(0), 0)
    return roleColors[hash % roleColors.length]
  }

  const inputClass = 'bg-gray-800 border border-gray-700 text-gray-100 placeholder-gray-500 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-[#193a29] focus:border-transparent'
  const inputSmClass = 'bg-gray-800 border border-gray-700 text-gray-100 placeholder-gray-500 rounded px-2 py-1 text-sm focus:outline-none focus:ring-1 focus:ring-[#193a29] focus:border-transparent w-full'
  const selectSmClass = inputSmClass + ' cursor-pointer'
  const selectClass = inputClass + ' cursor-pointer'
  const btnPrimary = 'text-white px-4 py-2 rounded-lg text-sm font-medium transition-colors hover:brightness-125'
  const btnDanger = 'text-gray-500 hover:text-red-400 text-sm font-medium transition-colors'
  const btnEdit = 'text-gray-500 hover:text-gray-200 text-sm font-medium transition-colors'
  const btnSave = 'text-emerald-400 hover:text-emerald-300 text-sm font-medium transition-colors'
  const btnCancel = 'text-gray-600 hover:text-gray-400 text-sm font-medium transition-colors'

  if (loading) {
    return (
      <div className="min-h-screen bg-gray-950 flex items-center justify-center text-gray-400 text-sm">
        Loading...
      </div>
    )
  }

  if (loadError) {
    return (
      <div className="min-h-screen bg-gray-950 flex items-center justify-center px-6">
        <div className="max-w-md text-center">
          <p className="text-4xl mb-4">📡</p>
          <h1 className="text-lg font-semibold text-gray-100 mb-2">Can&apos;t reach the database</h1>
          <p className="text-sm text-gray-400 mb-6">
            The app couldn&apos;t connect to Supabase. This is usually a network issue — a VPN, corporate firewall, or DNS block on your device preventing access to <span className="text-gray-300">supabase.co</span>. Your data is safe.
          </p>
          <button
            onClick={loadAll}
            className="text-white px-5 py-2 rounded-lg text-sm font-medium transition-colors hover:brightness-125"
            style={{ backgroundColor: '#193a29' }}
          >
            Retry
          </button>
          <p className="text-xs text-gray-600 mt-4">If it keeps failing, try a different network (e.g. phone hotspot) or disconnect any VPN.</p>
        </div>
      </div>
    )
  }

  return (
    <div className="min-h-screen bg-gray-950 text-gray-100">
      <div className="border-b px-8 py-5 flex items-center justify-between" style={{ borderColor: '#193a29' }}>
        <h1 className="text-xl font-semibold text-gray-100 tracking-tight">Code Pod Project Manager</h1>
        <div className="flex items-center gap-1 bg-gray-900 rounded-lg p-1">
          {(['dark', 'system', 'light'] as const).map(t => (
            <button
              key={t}
              onClick={() => setTheme(t)}
              className={`px-3 py-1 rounded-md text-xs font-medium capitalize transition-all ${
                theme === t ? 'text-white' : 'text-gray-500 hover:text-gray-300'
              }`}
              style={theme === t ? { backgroundColor: '#193a29' } : {}}
            >
              {t === 'dark' ? '🌙' : t === 'light' ? '☀️' : '⚙️'} {t}
            </button>
          ))}
        </div>
      </div>

      <div className="max-w-4xl mx-auto px-8 py-8">
        {/* Tabs */}
        <div className="flex gap-1 mb-8 bg-gray-900 rounded-xl p-1 w-fit">
          {([
            { key: 'dashboard', label: 'Dashboard' },
            { key: 'projects', label: 'Projects' },
            { key: 'staff', label: 'Staff' },
            { key: 'assignments', label: 'Assignments' },
            { key: 'scenarios', label: 'Staffing Scenarios' },
            { key: 'milestones', label: 'Milestones' },
          ] as const).map(({ key, label }) => (
            <button
              key={key}
              onClick={() => setTab(key)}
              className={`px-5 py-2 text-sm font-medium rounded-lg transition-all ${
                tab === key ? 'text-white shadow' : 'text-gray-400 hover:text-gray-200'
              }`}
              style={tab === key ? { backgroundColor: '#193a29' } : {}}
            >
              {label}
            </button>
          ))}
        </div>

        {/* Projects Tab */}
        {tab === 'projects' && (
          <div>
            <div className="bg-gray-900 rounded-xl p-5 mb-6 border border-gray-800">
              <div className="flex items-center justify-between mb-4">
                <h2 className="text-xs font-semibold text-gray-400 uppercase tracking-wider">Add Project</h2>
                {statusFilterDropdown()}
              </div>
              <div className="flex flex-wrap gap-2">
                <input
                  className={inputClass}
                  placeholder="Project name *"
                  value={newProject.name}
                  onChange={e => setNewProject({ ...newProject, name: e.target.value })}
                  onKeyDown={e => e.key === 'Enter' && addProject()}
                />
                <input
                  className={inputClass}
                  placeholder="Customer Codename"
                  value={newProject.customer_codename}
                  onChange={e => setNewProject({ ...newProject, customer_codename: e.target.value })}
                />
                <select
                  className={selectClass}
                  value={newProject.status}
                  onChange={e => setNewProject({ ...newProject, status: e.target.value })}
                >
                  <option value="active">Active</option>
                  <option value="starting-soon">Starting Soon</option>
                  <option value="on-hold">On Hold</option>
                  <option value="paused">Paused</option>
                  <option value="completed">Completed</option>
                </select>
                <input
                  className={inputClass}
                  type="number"
                  placeholder="Expected Duration (Weeks)"
                  value={newProject.duration_weeks}
                  onChange={e => setNewProject({ ...newProject, duration_weeks: e.target.value })}
                  style={{ width: 220 }}
                />
                <button className={btnPrimary} style={{ backgroundColor: '#193a29' }} onClick={addProject}>Add</button>
              </div>
            </div>

            {projects.length === 0 ? (
              <p className="text-gray-600 text-sm">No projects yet.</p>
            ) : (
              <table className="w-full text-sm">
                <thead>
                  <tr className="text-left text-gray-500 border-b border-gray-800">
                    {[
                      { key: 'name', label: 'Name' },
                      { key: 'customer', label: 'Customer' },
                      { key: 'status', label: 'Status' },
                      { key: 'start', label: 'Start' },
                      { key: 'end', label: 'End' },
                      { key: 'duration', label: 'Duration (Wks)' },
                      { key: 'staff', label: 'Staff' },
                    ].map(({ key, label }) => (
                      <th key={key} className="pb-3 pr-4 font-medium text-xs uppercase tracking-wider">
                        <button onClick={() => toggleProjectSort(key)} className={`flex items-center gap-1.5 hover:text-gray-300 transition-colors ${projectSort.col === key ? 'text-gray-300' : ''}`}>
                          {label} <span className="text-[10px] border border-gray-700 rounded px-1 py-0.5">{colSortLabel(key)}</span>
                        </button>
                      </th>
                    ))}
                    <th></th>
                  </tr>
                </thead>
                <tbody>
                  {sortedProjects(projects.filter(p => visibleStatuses[p.status])).map(p => {
                    const count = assignments.filter(a => a.project_id === p.id).length
                    const isEditing = editingProjectId === p.id
                    return (
                      <tr key={p.id} className="border-b border-gray-800/60 hover:bg-gray-900/60 transition-colors">
                        {isEditing ? (
                          <>
                            <td className="py-2 pr-2"><input className={inputSmClass} value={editProject.name} onChange={e => setEditProject({ ...editProject, name: e.target.value })} /></td>
                            <td className="py-2 pr-2"><input className={inputSmClass} value={editProject.customer_codename} onChange={e => setEditProject({ ...editProject, customer_codename: e.target.value })} /></td>
                            <td className="py-2 pr-2">
                              <select className={selectSmClass} value={editProject.status} onChange={e => setEditProject({ ...editProject, status: e.target.value })}>
                                <option value="active">Active</option>
                                <option value="starting-soon">Starting Soon</option>
                                <option value="on-hold">On Hold</option>
                                <option value="paused">Paused</option>
                                <option value="completed">Completed</option>
                              </select>
                            </td>
                            <td className="py-2 text-gray-500">{p.start_date ?? '—'}</td>
                            <td className="py-2 text-gray-500 text-xs">auto</td>
                            <td className="py-2 pr-2"><input className={inputSmClass} type="number" placeholder="Weeks" value={editProject.duration_weeks} onChange={e => setEditProject({ ...editProject, duration_weeks: e.target.value })} /></td>
                            <td className="py-2 text-gray-500">{count} assigned</td>
                            <td className="py-2 text-right">
                              <div className="flex justify-end gap-3">
                                <button className={btnSave} onClick={() => saveProject(p.id)}>Save</button>
                                <button className={btnCancel} onClick={() => setEditingProjectId(null)}>Cancel</button>
                              </div>
                            </td>
                          </>
                        ) : (
                          <>
                            <td className="py-3.5 font-medium text-gray-100">{p.name}</td>
                            <td className="py-3.5 text-gray-500">{p.customer_codename ?? '—'}</td>
                            <td className="py-3.5">
                              <select
                                value={p.status}
                                onChange={e => updateProjectStatus(p.id, e.target.value)}
                                title="Change project status"
                                className={`px-2.5 py-1 pr-6 rounded-full text-xs font-medium cursor-pointer appearance-none focus:outline-none focus:ring-1 focus:ring-gray-500 ${
                                  p.status === 'active' ? 'bg-emerald-500/10 text-emerald-400' :
                                  p.status === 'starting-soon' ? 'bg-sky-500/10 text-sky-400' :
                                  p.status === 'on-hold' ? 'bg-amber-500/10 text-amber-400' :
                                  p.status === 'paused' ? 'bg-rose-500/10 text-rose-400' :
                                  'bg-gray-700 text-gray-400'
                                }`}
                                style={{ backgroundImage: 'none' }}
                              >
                                {STATUS_OPTIONS.map(o => <option key={o.value} value={o.value} className="bg-gray-900 text-gray-100">{o.label}</option>)}
                              </select>
                            </td>
                            <td className="py-3.5 text-gray-500">{p.start_date ?? '—'}</td>
                            <td className="py-3.5 text-gray-500">{p.end_date ?? '—'}</td>
                            <td className="py-3.5 text-gray-500">
                              {p.start_date && p.end_date
                                ? `${Math.round((new Date(p.end_date).getTime() - new Date(p.start_date).getTime()) / (7 * 24 * 60 * 60 * 1000))} wks`
                                : '—'}
                            </td>
                            <td className="py-3.5 text-gray-500">{count} assigned</td>
                            <td className="py-3.5 text-right">
                              <div className="flex justify-end gap-3">
                                <button className={btnEdit} onClick={() => startEditProject(p)}>Edit</button>
                                <button className={btnDanger} onClick={() => deleteProject(p.id)}>Remove</button>
                              </div>
                            </td>
                          </>
                        )}
                      </tr>
                    )
                  })}
                </tbody>
              </table>
            )}
          </div>
        )}

        {/* Staff Tab */}
        {tab === 'staff' && (
          <div>
            <div className="bg-gray-900 rounded-xl p-5 mb-6 border border-gray-800">
              <div className="flex items-center justify-between mb-4">
                <h2 className="text-xs font-semibold text-gray-400 uppercase tracking-wider">Add Staff Member</h2>
                <div className="flex items-center gap-2">
                  <button
                    onClick={() => setHideAssigned(!hideAssigned)}
                    className={`text-xs px-3 py-1.5 rounded-lg border transition-colors ${
                      hideAssigned ? 'text-emerald-400' : 'border-gray-700 text-gray-400 hover:text-gray-200'
                    }`}
                    style={hideAssigned ? { borderColor: '#193a29' } : {}}
                  >
                    {hideAssigned ? 'Unassigned only' : 'Show unassigned only'}
                  </button>
                  <label className="cursor-pointer text-xs text-gray-400 hover:text-gray-200 transition-colors border border-gray-700 rounded-lg px-3 py-1.5">
                    Import CSV
                    <input type="file" accept=".csv" className="hidden" onChange={importStaffCSV} />
                  </label>
                </div>
              </div>
              <div className="flex flex-wrap gap-2">
                <input
                  className={inputClass}
                  placeholder="Full name *"
                  value={newStaff.name}
                  onChange={e => setNewStaff({ ...newStaff, name: e.target.value })}
                  onKeyDown={e => e.key === 'Enter' && addStaff()}
                />
                <select
                  className={selectClass}
                  value={newStaff.position}
                  onChange={e => setNewStaff({ ...newStaff, position: e.target.value })}
                >
                  <option value="">Position</option>
                  {VALID_POSITIONS.map(p => <option key={p}>{p}</option>)}
                </select>
                <label className="flex items-center gap-2 text-sm text-gray-400 cursor-pointer select-none">
                  <input type="checkbox" checked={newStaff.ooo} onChange={e => setNewStaff({ ...newStaff, ooo: e.target.checked, ooo_return_date: e.target.checked ? newStaff.ooo_return_date : '' })} className="accent-[#193a29] w-4 h-4" />
                  OOO
                </label>
                {newStaff.ooo && (
                  <input
                    type="date"
                    className={inputClass}
                    placeholder="Return date"
                    value={newStaff.ooo_return_date}
                    onChange={e => setNewStaff({ ...newStaff, ooo_return_date: e.target.value })}
                  />
                )}
                <button className={btnPrimary} style={{ backgroundColor: '#193a29' }} onClick={addStaff}>Add</button>
              </div>
              {csvErrors.length > 0 && (
                <div className="mt-4 bg-red-500/10 border border-red-500/20 rounded-lg p-3">
                  <p className="text-red-400 text-xs font-semibold mb-1">CSV import failed — fix these issues and try again:</p>
                  {csvErrors.map((err, i) => <p key={i} className="text-red-400 text-xs">{err}</p>)}
                </div>
              )}
            </div>

            {staff.length === 0 ? (
              <p className="text-gray-600 text-sm">No staff yet.</p>
            ) : (
              <>
              <table className="w-full text-sm">
                <thead>
                  <tr className="text-left text-gray-500 border-b border-gray-800">
                    <th className="pb-3 pr-6 font-medium text-xs uppercase tracking-wider w-40">
                      <button onClick={() => setStaffSort(nextSort(staffSort))} className="flex items-center gap-1.5 hover:text-gray-300 transition-colors">
                        Name <span className="text-[10px] border border-gray-700 rounded px-1 py-0.5">{sortLabel(staffSort)}</span>
                      </button>
                    </th>
                    <th className="pb-3 pr-6 font-medium text-xs uppercase tracking-wider w-36">Position</th>
                    <th className="pb-3 pr-6 font-medium text-xs uppercase tracking-wider w-28">OOO</th>
                    <th className="pb-3 pr-6 font-medium text-xs uppercase tracking-wider">Assigned To</th>
                    <th></th>
                  </tr>
                </thead>
                <tbody>
                  {sortedList(staff.filter(s => !hideAssigned || !assignments.some(a => a.staff_id === s.id)), staffSort).map(s => {
                    const isEditing = editingStaffId === s.id
                    const staffAssignments = assignments.filter(a => a.staff_id === s.id)
                    return (
                      <tr key={s.id} className="border-b border-gray-800/60 hover:bg-gray-900/60 transition-colors">
                        {isEditing ? (
                          <>
                            <td className="py-2 pr-2"><input className={inputSmClass} value={editStaff.name} onChange={e => setEditStaff({ ...editStaff, name: e.target.value })} /></td>
                            <td className="py-2 pr-2">
                              <select className={selectSmClass} value={editStaff.position} onChange={e => setEditStaff({ ...editStaff, position: e.target.value })}>
                                <option value="">Position</option>
                                <option>SPA</option>
                                <option>SPL I</option>
                                <option>SPL II</option>
                                <option>Manager, Delivery</option>
                                <option>Senior SPL</option>
                                <option>Head of Delivery</option>
                                <option>GenAI Consultant</option>
                              </select>
                            </td>
                            <td className="py-2 pr-2 align-top">
                              <label className="flex items-center gap-2 text-sm text-gray-400 cursor-pointer mb-1">
                                <input type="checkbox" checked={editStaff.ooo} onChange={e => setEditStaff({ ...editStaff, ooo: e.target.checked, ooo_return_date: e.target.checked ? editStaff.ooo_return_date : '' })} className="accent-[#193a29] w-4 h-4" />
                                OOO
                              </label>
                              {editStaff.ooo && (
                                <input type="date" className={inputSmClass} value={editStaff.ooo_return_date} onChange={e => setEditStaff({ ...editStaff, ooo_return_date: e.target.value })} />
                              )}
                              <input
                                type="text"
                                className={inputSmClass + ' mt-1'}
                                placeholder="Flex notes"
                                value={editStaff.flex_notes}
                                onChange={e => setEditStaff({ ...editStaff, flex_notes: e.target.value })}
                              />
                            </td>
                            <td className="py-2 pr-2 align-top">
                              <div className="flex flex-wrap gap-1 mb-1.5">
                                {staffAssignments.map(a => {
                                  const project = projects.find(p => p.id === a.project_id)
                                  return project ? (
                                    <span key={a.id} className="text-xs px-2 py-0.5 rounded-full bg-gray-800 text-gray-300 flex items-center gap-1">
                                      {project.name}
                                      {a.assignment_role && <span className={`px-1.5 py-0.5 rounded-full text-[10px] font-medium ${roleColor(a.assignment_role)}`}>{a.assignment_role}</span>}
                                      <button onClick={() => removeAssignment(a.id)} className="text-gray-500 hover:text-red-400 ml-0.5 leading-none">×</button>
                                    </span>
                                  ) : null
                                })}
                              </div>
                              <div className="flex gap-1">
                                <select
                                  className={selectSmClass}
                                  value={inlineAssignment.project_id}
                                  onChange={e => setInlineAssignment({ ...inlineAssignment, project_id: e.target.value })}
                                >
                                  <option value="">Add project…</option>
                                  {[...projects]
                                    .sort((a, b) => a.name.localeCompare(b.name))
                                    .filter(p => !staffAssignments.some(a => a.project_id === p.id))
                                    .map(p => <option key={p.id} value={p.id}>{p.name}</option>)}
                                </select>
                                <select
                                  className={selectSmClass}
                                  value={inlineAssignment.assignment_role}
                                  onChange={e => setInlineAssignment({ ...inlineAssignment, assignment_role: e.target.value })}
                                >
                                  <option value="">Role</option>
                                  <option>Supervisor</option>
                                  <option>STO</option>
                                  <option>Ops Support</option>
                                </select>
                                <button
                                  className="text-xs px-2 py-1 rounded bg-gray-700 text-gray-300 hover:text-white transition-colors whitespace-nowrap"
                                  onClick={async () => {
                                    if (!inlineAssignment.project_id) return
                                    const { data } = await supabase.from('assignments').insert([{ project_id: inlineAssignment.project_id, staff_id: s.id, assignment_role: inlineAssignment.assignment_role || null }]).select().single()
                                    if (data) { setAssignments(prev => [...prev, data]); setInlineAssignment({ project_id: '', assignment_role: '' }) }
                                  }}
                                >+ Add</button>
                              </div>
                            </td>
                            <td className="py-2 text-right align-top">
                              <div className="flex justify-end gap-3">
                                <button className={btnSave} onClick={() => saveStaff(s.id)}>Save</button>
                                <button className={btnCancel} onClick={() => { setEditingStaffId(null); setInlineAssignment({ project_id: '', assignment_role: '' }) }}>Cancel</button>
                              </div>
                            </td>
                          </>
                        ) : (
                          <>
                            <td className="py-3.5 pr-6 font-medium text-gray-100">
                              <span className="inline-flex items-center gap-1.5">
                                {s.name}
                                {s.flexed && <span className="text-[10px] font-medium px-1.5 py-0.5 rounded-full bg-violet-500/10 text-violet-400">⤢ Flexed</span>}
                              </span>
                              {s.flexed && s.flex_notes && <span className="block text-xs font-normal text-violet-400 mt-0.5">📝 {s.flex_notes}</span>}
                            </td>
                            <td className="py-3.5 pr-6 text-gray-500">{s.position ?? '—'}</td>
                            <td className="py-3.5 pr-6">
                              {s.ooo ? (
                                <div className="flex flex-col gap-0.5">
                                  <span className="text-xs font-medium px-2 py-0.5 rounded-full bg-amber-500/10 text-amber-400 w-fit">OOO</span>
                                  {s.ooo_return_date && <span className="text-xs text-gray-500">Back {s.ooo_return_date}</span>}
                                </div>
                              ) : (
                                <span className="text-xs text-gray-600">—</span>
                              )}
                            </td>
                            <td className="py-3.5">
                              {staffAssignments.length === 0 ? (
                                <span className="text-gray-600 text-xs">Unassigned</span>
                              ) : (
                                <div className="flex flex-wrap gap-1">
                                  {staffAssignments.map(a => {
                                    const project = projects.find(p => p.id === a.project_id)
                                    return project ? (
                                      <span key={a.id} className="text-xs px-2 py-0.5 rounded-full bg-gray-800 text-gray-300 flex items-center gap-1">
                                        {project.name}
                                        {a.assignment_role && (
                                          <span className={`px-1.5 py-0.5 rounded-full text-[10px] font-medium ${roleColor(a.assignment_role)}`}>{a.assignment_role}</span>
                                        )}
                                      </span>
                                    ) : null
                                  })}
                                </div>
                              )}
                            </td>
                            <td className="py-3.5 text-right">
                              <div className="flex justify-end gap-3">
                                <button className={btnEdit} onClick={() => startEditStaff(s)}>Edit</button>
                                <button className={btnDanger} onClick={() => deleteStaff(s.id)}>Remove</button>
                              </div>
                            </td>
                          </>
                        )}
                      </tr>
                    )
                  })}
                </tbody>
              </table>
              <p className="text-xs text-gray-500 mt-4 text-right">Total headcount: <span className="text-gray-300 font-medium">{staff.length}</span></p>
              </>
            )}
          </div>
        )}

        {/* Assignments Tab */}
        {tab === 'assignments' && (
          <div>
            <div className="bg-gray-900 rounded-xl p-5 mb-6 border border-gray-800">
              <div className="flex items-center justify-between mb-4">
                <h2 className="text-xs font-semibold text-gray-400 uppercase tracking-wider">Assign Staff to Project</h2>
                <div className="flex items-center gap-2">
                <button
                  onClick={() => setShowAddProjectInline(v => !v)}
                  className={`text-xs px-3 py-1.5 rounded-lg border transition-colors ${
                    showAddProjectInline ? 'text-emerald-400' : 'border-gray-700 text-gray-400 hover:text-gray-200'
                  }`}
                  style={showAddProjectInline ? { borderColor: '#193a29' } : {}}
                >
                  + New Project
                </button>
                {statusFilterDropdown()}
                <button
                  onClick={exportAssignmentsHTML}
                  className="text-xs px-3 py-1.5 rounded-lg border border-gray-700 text-gray-400 hover:text-gray-200 transition-colors"
                >
                  Download as HTML
                </button>
                <button
                  onClick={exportAssignmentsDOCX}
                  className="text-xs px-3 py-1.5 rounded-lg border border-gray-700 text-gray-400 hover:text-gray-200 transition-colors"
                >
                  Download as DOCX
                </button>
                <button
                  onClick={() => setHideAssignedInDropdown(!hideAssignedInDropdown)}
                  className={`text-xs px-3 py-1.5 rounded-lg border transition-colors ${
                    hideAssignedInDropdown ? 'text-emerald-400' : 'border-gray-700 text-gray-400 hover:text-gray-200'
                  }`}
                  style={hideAssignedInDropdown ? { borderColor: '#193a29' } : {}}
                >
                  {hideAssignedInDropdown ? 'Unassigned only' : 'Show unassigned only'}
                </button>
                </div>
              </div>
              <div className="flex flex-wrap gap-2">
                <select
                  className={selectClass}
                  value={newAssignment.project_id}
                  onChange={e => setNewAssignment({ ...newAssignment, project_id: e.target.value })}
                >
                  <option value="">Select project *</option>
                  {[...projects].filter(p => visibleStatuses[p.status]).sort((a, b) => statusRank(a.status) - statusRank(b.status) || a.name.localeCompare(b.name)).map(p => <option key={p.id} value={p.id}>{p.name}</option>)}
                </select>
                <select
                  className={selectClass}
                  value={newAssignment.staff_id}
                  onChange={e => setNewAssignment({ ...newAssignment, staff_id: e.target.value })}
                >
                  <option value="">Select staff *</option>
                  {staff
                    .filter(s => !hideAssignedInDropdown || !assignments.some(a => a.staff_id === s.id))
                    .sort((a, b) => a.name.localeCompare(b.name))
                    .map(s => <option key={s.id} value={s.id}>{s.name}</option>)}
                </select>
                <select
                  className={selectClass}
                  value={newAssignment.assignment_role}
                  onChange={e => setNewAssignment({ ...newAssignment, assignment_role: e.target.value })}
                >
                  <option value="">Role</option>
                  <option>Supervisor</option>
                  <option>STO</option>
                  <option>Ops Support</option>
                </select>
                <button className={btnPrimary} style={{ backgroundColor: '#193a29' }} onClick={addAssignment}>Assign</button>
              </div>

              {showAddProjectInline && (
                <div className="mt-4 pt-4 border-t border-gray-800">
                  <p className="text-xs font-semibold text-gray-400 uppercase tracking-wider mb-3">Add Project</p>
                  <div className="flex flex-wrap gap-2">
                    <input
                      className={inputClass}
                      placeholder="Project name *"
                      value={newProject.name}
                      onChange={e => setNewProject({ ...newProject, name: e.target.value })}
                      onKeyDown={e => e.key === 'Enter' && addProject()}
                    />
                    <input
                      className={inputClass}
                      placeholder="Customer Codename"
                      value={newProject.customer_codename}
                      onChange={e => setNewProject({ ...newProject, customer_codename: e.target.value })}
                    />
                    <select
                      className={selectClass}
                      value={newProject.status}
                      onChange={e => setNewProject({ ...newProject, status: e.target.value })}
                    >
                      {STATUS_OPTIONS.map(o => <option key={o.value} value={o.value}>{o.label}</option>)}
                    </select>
                    <input
                      className={inputClass}
                      type="number"
                      placeholder="Expected Duration (Weeks)"
                      value={newProject.duration_weeks}
                      onChange={e => setNewProject({ ...newProject, duration_weeks: e.target.value })}
                      style={{ width: 220 }}
                    />
                    <button className={btnPrimary} style={{ backgroundColor: '#193a29' }} onClick={addProject}>Add</button>
                  </div>
                </div>
              )}
            </div>

            {/* Two-column layout */}
            <div className="flex gap-6 items-start">

            {/* Drop role modal */}
            {pendingDrop && (
              <div className="fixed inset-0 bg-black/50 flex items-center justify-center z-50" onClick={() => { setPendingDrop(null); setDropRole('') }}>
                <div className="bg-gray-900 border border-gray-700 rounded-xl p-6 w-80 shadow-xl" onClick={e => e.stopPropagation()}>
                  <p className="text-sm font-semibold text-gray-100 mb-1">Assign to project</p>
                  <p className="text-xs text-gray-500 mb-4">
                    <span className="text-gray-300">{staff.find(s => s.id === pendingDrop.staffId)?.name}</span>
                    {' → '}
                    <span className="text-gray-300">{projects.find(p => p.id === pendingDrop.projectId)?.name}</span>
                  </p>
                  <select
                    className={selectClass + ' w-full mb-4'}
                    value={dropRole}
                    onChange={e => setDropRole(e.target.value)}
                    autoFocus
                  >
                    <option value="">No role</option>
                    <option>Supervisor</option>
                    <option>STO</option>
                    <option>Ops Support</option>
                  </select>
                  <div className="flex gap-2 justify-end">
                    <button className={btnCancel} onClick={() => { setPendingDrop(null); setDropRole('') }}>Cancel</button>
                    <button
                      className={btnPrimary}
                      style={{ backgroundColor: '#193a29' }}
                      onClick={async () => {
                        const already = assignments.some(a => a.project_id === pendingDrop.projectId && a.staff_id === pendingDrop.staffId)
                        if (!already) {
                          const { data } = await supabase.from('assignments').insert([{
                            project_id: pendingDrop.projectId,
                            staff_id: pendingDrop.staffId,
                            assignment_role: dropRole || null,
                          }]).select().single()
                          if (data) setAssignments(prev => [...prev, data])
                        }
                        const member = staff.find(s => s.id === pendingDrop.staffId)
                        if (member?.flexed) {
                          const { data } = await supabase.from('staff').update({ flexed: false }).eq('id', pendingDrop.staffId).select().single()
                          if (data) setStaff(prev => prev.map(s => s.id === pendingDrop.staffId ? data : s))
                        }
                        setPendingDrop(null)
                        setDropRole('')
                      }}
                    >
                      Assign
                    </button>
                  </div>
                </div>
              </div>
            )}

            {pendingMove && (() => {
              const assignment = assignments.find(a => a.id === pendingMove.assignmentId)
              const member = staff.find(s => s.id === assignment?.staff_id)
              return (
                <div className="fixed inset-0 bg-black/50 flex items-center justify-center z-50" onClick={() => { setPendingMove(null); setDropRole('') }}>
                  <div className="bg-gray-900 border border-gray-700 rounded-xl p-6 w-80 shadow-xl" onClick={e => e.stopPropagation()}>
                    <p className="text-sm font-semibold text-gray-100 mb-1">Move to project</p>
                    <p className="text-xs text-gray-500 mb-4">
                      <span className="text-gray-300">{member?.name}</span>
                      {' → '}
                      <span className="text-gray-300">{projects.find(p => p.id === pendingMove.projectId)?.name}</span>
                    </p>
                    <p className="text-xs text-gray-500 mb-2">Confirm role for this assignment:</p>
                    <select
                      className={selectClass + ' w-full mb-4'}
                      value={dropRole}
                      onChange={e => setDropRole(e.target.value)}
                      autoFocus
                    >
                      <option value="">No role</option>
                      <option>Supervisor</option>
                      <option>STO</option>
                      <option>Ops Support</option>
                    </select>
                    <div className="flex gap-2 justify-end">
                      <button className={btnCancel} onClick={() => { setPendingMove(null); setDropRole('') }}>Cancel</button>
                      <button
                        className={btnPrimary}
                        style={{ backgroundColor: '#193a29' }}
                        onClick={async () => {
                          const { data } = await supabase.from('assignments').update({ project_id: pendingMove.projectId, assignment_role: dropRole || null }).eq('id', pendingMove.assignmentId).select().single()
                          if (data) setAssignments(prev => prev.map(a => a.id === pendingMove.assignmentId ? data : a))
                          setPendingMove(null)
                          setDropRole('')
                        }}
                      >
                        Move
                      </button>
                    </div>
                  </div>
                </div>
              )
            })()}

            {/* Left: project list */}
            <div className="flex-1 min-w-0">
            {projects.length === 0 ? (
              <p className="text-gray-600 text-sm">Add projects first.</p>
            ) : (
              <div className="space-y-3">
                {[...projects].filter(p => visibleStatuses[p.status]).sort((a, b) => statusRank(a.status) - statusRank(b.status) || a.name.localeCompare(b.name)).map(p => {
                  const projectAssignments = assignments.filter(a => a.project_id === p.id)
                  const isOver = dragOverProjectId === p.id
                  const roles = projectAssignments.map(a => a.assignment_role)
                  const stoCount = roles.filter(r => r === 'STO').length
                  const opsCount = roles.filter(r => r === 'Ops Support').length
                  const staffingStatus =
                    stoCount === 0 || (stoCount >= 1 && opsCount === 0) ? 'understaffed' :
                    stoCount >= 2 || opsCount > 2 ? 'overstaffed' :
                    'good'
                  return (
                    <div
                      key={p.id}
                      className={`border rounded-xl p-5 transition-all ${
                        isOver
                          ? 'border-[#193a29] bg-[#193a29]/10 scale-[1.01]'
                          : p.flagged
                            ? 'border-red-500/40 bg-red-500/5'
                            : 'border-gray-800 bg-gray-900/40'
                      }`}
                      onDragOver={e => { e.preventDefault(); setDragOverProjectId(p.id) }}
                      onDragLeave={() => setDragOverProjectId(null)}
                      onDrop={async e => {
                        e.preventDefault()
                        setDragOverProjectId(null)
                        const type = e.dataTransfer.getData('type')
                        const id = e.dataTransfer.getData('id')
                        if (type === 'staff' && id) {
                          const already = assignments.some(a => a.project_id === p.id && a.staff_id === id)
                          if (!already) { setPendingDrop({ staffId: id, projectId: p.id }) }
                          setDraggedStaffId(null)
                        } else if (type === 'assignment' && id) {
                          const assignment = assignments.find(a => a.id === id)
                          if (assignment && assignment.project_id !== p.id) {
                            const already = assignments.some(a => a.project_id === p.id && a.staff_id === assignment.staff_id)
                            if (!already) {
                              setPendingMove({ assignmentId: id, projectId: p.id })
                              setDropRole(assignment.assignment_role ?? '')
                            }
                          }
                          setDraggedAssignmentId(null)
                        }
                      }}
                    >
                      <div className="flex items-center justify-between mb-4">
                        <div className="flex items-center gap-2">
                          <h3 className="font-semibold text-gray-100">{p.name}</h3>
                          <button
                            onClick={() => { setAddingToProjectId(addingToProjectId === p.id ? null : p.id); setQuickAdd({ staff_id: '', assignment_role: '' }) }}
                            title="Add a person to this project"
                            className="w-5 h-5 flex items-center justify-center rounded-full border border-gray-600 text-gray-400 hover:text-white hover:border-gray-400 transition-colors text-sm leading-none"
                          >
                            +
                          </button>
                          <button
                            onClick={() => toggleProjectFlag(p.id, !p.flagged)}
                            title={p.flagged ? 'Flagged: needs more support — click to clear' : 'Flag: needs more support'}
                            className={`text-sm leading-none transition-colors ${p.flagged ? '' : 'opacity-40 hover:opacity-100 grayscale'}`}
                          >
                            🚩
                          </button>
                          {p.flagged && (
                            <span className="px-2 py-0.5 rounded-full text-[11px] font-medium bg-red-500/10 text-red-400">Needs more support</span>
                          )}
                          {p.status === 'active' && (
                            <span className={`px-2 py-0.5 rounded-full text-[11px] font-medium ${
                              staffingStatus === 'understaffed' ? 'bg-amber-500/10 text-amber-400' :
                              staffingStatus === 'overstaffed' ? 'bg-blue-500/10 text-blue-400' :
                              'bg-emerald-500/10 text-emerald-400'
                            }`}>
                              {staffingStatus === 'understaffed' ? 'Potentially understaffed' :
                               staffingStatus === 'overstaffed' ? 'Potentially overstaffed' :
                               'Well staffed'}
                            </span>
                          )}
                        </div>
                        <select
                          value={p.status}
                          onChange={e => updateProjectStatus(p.id, e.target.value)}
                          title="Change project status"
                          className={`px-2.5 py-1 pr-6 rounded-full text-xs font-medium cursor-pointer appearance-none focus:outline-none focus:ring-1 focus:ring-gray-500 ${
                            p.status === 'active' ? 'bg-emerald-500/10 text-emerald-400' :
                            p.status === 'starting-soon' ? 'bg-sky-500/10 text-sky-400' :
                            p.status === 'on-hold' ? 'bg-amber-500/10 text-amber-400' :
                            p.status === 'paused' ? 'bg-rose-500/10 text-rose-400' :
                            'bg-gray-700 text-gray-400'
                          }`}
                          style={{ backgroundImage: 'none' }}
                        >
                          {STATUS_OPTIONS.map(o => <option key={o.value} value={o.value} className="bg-gray-900 text-gray-100">{o.label}</option>)}
                        </select>
                      </div>
                      {addingToProjectId === p.id && (
                        <div className="flex flex-wrap gap-2 mb-3 p-3 rounded-lg bg-gray-800/40 border border-gray-700">
                          <select
                            className={selectSmClass + ' flex-1 min-w-[160px]'}
                            value={quickAdd.staff_id}
                            onChange={e => setQuickAdd({ ...quickAdd, staff_id: e.target.value })}
                          >
                            <option value="">Select staff…</option>
                            {[...staff]
                              .filter(s => !projectAssignments.some(a => a.staff_id === s.id))
                              .sort((a, b) => a.name.localeCompare(b.name))
                              .map(s => <option key={s.id} value={s.id}>{s.name}{s.position ? ` — ${s.position}` : ''}</option>)}
                          </select>
                          <select
                            className={selectSmClass + ' w-32'}
                            value={quickAdd.assignment_role}
                            onChange={e => setQuickAdd({ ...quickAdd, assignment_role: e.target.value })}
                          >
                            <option value="">Role</option>
                            <option>Supervisor</option>
                            <option>STO</option>
                            <option>Ops Support</option>
                          </select>
                          <button
                            className={btnPrimary}
                            style={{ backgroundColor: '#193a29' }}
                            onClick={async () => {
                              if (!quickAdd.staff_id) return
                              const { data } = await supabase.from('assignments').insert([{
                                project_id: p.id,
                                staff_id: quickAdd.staff_id,
                                assignment_role: quickAdd.assignment_role || null,
                              }]).select().single()
                              if (data) setAssignments(prev => [...prev, data])
                              setQuickAdd({ staff_id: '', assignment_role: '' })
                              setAddingToProjectId(null)
                            }}
                          >
                            Add
                          </button>
                          <button className={btnCancel} onClick={() => { setAddingToProjectId(null); setQuickAdd({ staff_id: '', assignment_role: '' }) }}>Cancel</button>
                        </div>
                      )}
                      {projectAssignments.length === 0 ? (
                        <p className={`text-sm ${isOver ? 'text-gray-400' : 'text-gray-600'}`}>
                          {isOver ? 'Drop to assign' : 'No staff assigned.'}
                        </p>
                      ) : (
                        <div className="space-y-2">
                          {[...projectAssignments].sort((a, b) => {
                            const order = ['Supervisor', 'STO', 'Ops Support']
                            return (order.indexOf(a.assignment_role ?? '') === -1 ? 99 : order.indexOf(a.assignment_role ?? '')) -
                                   (order.indexOf(b.assignment_role ?? '') === -1 ? 99 : order.indexOf(b.assignment_role ?? ''))
                          }).map(a => {
                            const member = staff.find(s => s.id === a.staff_id)
                            const nonSupProjectCount = new Set(
                              assignments.filter(x => x.staff_id === a.staff_id && x.assignment_role !== 'Supervisor').map(x => x.project_id)
                            ).size
                            const overAllocated = a.assignment_role !== 'Supervisor' && nonSupProjectCount >= 2
                            return (
                              <div
                                key={a.id}
                                draggable
                                onDragStart={e => { e.dataTransfer.setData('type', 'assignment'); e.dataTransfer.setData('id', a.id); setDraggedAssignmentId(a.id); setDraggedStaffId(null) }}
                                onDragEnd={() => setDraggedAssignmentId(null)}
                                className={`flex items-center justify-between bg-gray-800/60 rounded-lg px-4 py-2.5 text-sm cursor-grab active:cursor-grabbing transition-opacity ${draggedAssignmentId === a.id ? 'opacity-40' : ''}`}
                              >
                                <div className="flex items-center gap-3">
                                  <span className="font-medium text-gray-200">{member?.name ?? 'Unknown'}</span>
                                  {member?.ooo && (
                                    <span className="flex items-center gap-1 text-xs font-medium px-2 py-0.5 rounded-full bg-amber-500/10 text-amber-400" title={member.ooo_return_date ? `OOO — back ${member.ooo_return_date}` : 'Out of office'}>
                                      ⚠️ OOO{member.ooo_return_date ? ` · back ${member.ooo_return_date}` : ''}
                                    </span>
                                  )}
                                  {overAllocated && (
                                    <span className="flex items-center gap-1 text-xs font-medium px-2 py-0.5 rounded-full bg-red-500/10 text-red-400" title={`Assigned to ${nonSupProjectCount} projects in a non-Supervisor role`}>
                                      ⚠️ In {nonSupProjectCount} projects
                                    </span>
                                  )}
                                  <select
                                    value={a.assignment_role ?? ''}
                                    onChange={e => updateAssignmentRole(a.id, e.target.value)}
                                    onClick={e => e.stopPropagation()}
                                    title="Change role"
                                    className={`text-xs font-medium px-2 py-0.5 pr-5 rounded-full cursor-pointer appearance-none focus:outline-none focus:ring-1 focus:ring-gray-500 ${a.assignment_role ? roleColor(a.assignment_role) : 'bg-gray-700/60 text-gray-400'}`}
                                    style={{ backgroundImage: 'none' }}
                                  >
                                    <option value="" className="bg-gray-900 text-gray-100">No role</option>
                                    <option value="Supervisor" className="bg-gray-900 text-gray-100">Supervisor</option>
                                    <option value="STO" className="bg-gray-900 text-gray-100">STO</option>
                                    <option value="Ops Support" className="bg-gray-900 text-gray-100">Ops Support</option>
                                  </select>
                                  {member?.position && <span className="text-xs text-gray-500">{member.position}</span>}
                                </div>
                                <button className={btnDanger} onClick={() => removeAssignment(a.id)}>×</button>
                              </div>
                            )
                          })}
                        </div>
                      )}
                    </div>
                  )
                })}
              </div>
            )}
            </div>{/* end left column */}

            {/* Right: unassigned + flexed + OOO panels */}
            {(() => {
              const assignedIds = new Set(assignments.map(a => a.staff_id))
              const unassigned = [...staff].filter(s => !assignedIds.has(s.id) && !s.flexed && !s.ooo && !s.onboarding).sort((a, b) => a.name.localeCompare(b.name))
              const flexed = [...staff].filter(s => s.flexed && !assignedIds.has(s.id)).sort((a, b) => a.name.localeCompare(b.name))
              const oooStaff = [...staff].filter(s => s.ooo && !assignedIds.has(s.id)).sort((a, b) => a.name.localeCompare(b.name))
              const onboardingStaff = [...staff].filter(s => s.onboarding && !assignedIds.has(s.id)).sort((a, b) => a.name.localeCompare(b.name))

              const staffTile = (s: Staff, extraClass = '') => (
                <div
                  key={s.id}
                  draggable
                  onDragStart={e => { e.dataTransfer.setData('type', 'staff'); e.dataTransfer.setData('id', s.id); setDraggedStaffId(s.id); setDraggedAssignmentId(null) }}
                  onDragEnd={() => setDraggedStaffId(null)}
                  className={`cursor-grab active:cursor-grabbing select-none border rounded-lg px-3 py-2 text-sm transition-all ${
                    draggedStaffId === s.id ? 'opacity-40 border-gray-600' : `border-gray-700 bg-gray-900 hover:border-gray-500 ${extraClass}`
                  }`}
                >
                  <p className="font-medium text-gray-200 leading-tight">{s.name}</p>
                  {s.position && <p className="text-xs text-gray-500 mt-0.5">{s.position}</p>}
                  {s.ooo && <span className="text-[10px] px-1.5 py-0.5 rounded-full bg-amber-500/10 text-amber-400 mt-1 inline-block">OOO</span>}
                  {s.flexed && s.flex_notes && <p className="text-[10px] text-violet-400 mt-1 leading-snug">📝 {s.flex_notes}</p>}
                </div>
              )

              return (
                <div className="w-52 shrink-0 sticky top-6 flex flex-col gap-5">

                  {/* Unassigned drop zone */}
                  <div
                    onDragOver={e => { e.preventDefault(); setDragOverUnassigned(true) }}
                    onDragLeave={() => setDragOverUnassigned(false)}
                    onDrop={async e => {
                      e.preventDefault(); setDragOverUnassigned(false)
                      const type = e.dataTransfer.getData('type')
                      const id = e.dataTransfer.getData('id')
                      if (type === 'staff' && id) {
                        await setStaffZone(id, 'unassigned')
                        setDraggedStaffId(null)
                      } else if (type === 'assignment' && id) {
                        const assignment = assignments.find(a => a.id === id)
                        if (assignment) {
                          await setStaffZone(assignment.staff_id, 'unassigned')
                          await removeAssignment(id)
                        }
                        setDraggedAssignmentId(null)
                      }
                    }}
                    className={`rounded-lg p-2 transition-all min-h-[60px] ${dragOverUnassigned ? 'ring-1 ring-gray-500 bg-gray-800/40' : ''}`}
                  >
                    <p className="text-xs text-gray-500 uppercase tracking-wider mb-3">Unassigned</p>
                    {unassigned.length === 0
                      ? <p className="text-xs text-gray-600">None</p>
                      : <div className="flex flex-col gap-2">{unassigned.map(s => staffTile(s))}</div>
                    }
                  </div>

                  {/* Flexed drop zone */}
                  <div
                    onDragOver={e => { e.preventDefault(); setDragOverFlexed(true) }}
                    onDragLeave={() => setDragOverFlexed(false)}
                    onDrop={async e => {
                      e.preventDefault(); setDragOverFlexed(false)
                      const type = e.dataTransfer.getData('type')
                      const id = e.dataTransfer.getData('id')
                      if (type === 'staff' && id) {
                        await setStaffZone(id, 'flexed')
                        setDraggedStaffId(null)
                      } else if (type === 'assignment' && id) {
                        const assignment = assignments.find(a => a.id === id)
                        if (assignment) {
                          await setStaffZone(assignment.staff_id, 'flexed')
                          await removeAssignment(id)
                        }
                        setDraggedAssignmentId(null)
                      }
                    }}
                    className={`rounded-lg p-2 transition-all min-h-[60px] border border-dashed ${
                      dragOverFlexed ? 'border-violet-500 bg-violet-500/10' : 'border-gray-700'
                    }`}
                  >
                    <p className="text-xs text-gray-500 uppercase tracking-wider mb-3">Flexed</p>
                    {flexed.length === 0
                      ? <p className="text-xs text-gray-600">Drop here to flex</p>
                      : <div className="flex flex-col gap-2">{flexed.map(s => staffTile(s))}</div>
                    }
                  </div>

                  {/* Currently OOO drop zone */}
                  <div
                    onDragOver={e => { e.preventDefault(); setDragOverOOO(true) }}
                    onDragLeave={() => setDragOverOOO(false)}
                    onDrop={e => {
                      e.preventDefault(); setDragOverOOO(false)
                      const type = e.dataTransfer.getData('type')
                      const id = e.dataTransfer.getData('id')
                      if (type === 'staff' && id) {
                        const member = staff.find(s => s.id === id)
                        setOooDate(member?.ooo_return_date ?? '')
                        setPendingOOO({ staffId: id, assignmentId: null })
                        setDraggedStaffId(null)
                      } else if (type === 'assignment' && id) {
                        const assignment = assignments.find(a => a.id === id)
                        if (assignment) {
                          const member = staff.find(s => s.id === assignment.staff_id)
                          setOooDate(member?.ooo_return_date ?? '')
                          setPendingOOO({ staffId: assignment.staff_id, assignmentId: id })
                        }
                        setDraggedAssignmentId(null)
                      }
                    }}
                    className={`rounded-lg p-2 transition-all min-h-[60px] border border-dashed ${
                      dragOverOOO ? 'border-amber-500 bg-amber-500/10' : 'border-gray-700'
                    }`}
                  >
                    <p className="text-xs text-gray-500 uppercase tracking-wider mb-3">Currently OOO</p>
                    {oooStaff.length === 0
                      ? <p className="text-xs text-gray-600">Drop here to mark OOO</p>
                      : <div className="flex flex-col gap-2">{oooStaff.map(s => staffTile(s))}</div>
                    }
                  </div>

                  {/* Onboarding drop zone */}
                  <div
                    onDragOver={e => { e.preventDefault(); setDragOverOnboarding(true) }}
                    onDragLeave={() => setDragOverOnboarding(false)}
                    onDrop={async e => {
                      e.preventDefault(); setDragOverOnboarding(false)
                      const type = e.dataTransfer.getData('type')
                      const id = e.dataTransfer.getData('id')
                      if (type === 'staff' && id) {
                        await setStaffZone(id, 'onboarding')
                        setDraggedStaffId(null)
                      } else if (type === 'assignment' && id) {
                        const assignment = assignments.find(a => a.id === id)
                        if (assignment) {
                          await setStaffZone(assignment.staff_id, 'onboarding')
                          await removeAssignment(id)
                        }
                        setDraggedAssignmentId(null)
                      }
                    }}
                    className={`rounded-lg p-2 transition-all min-h-[60px] border border-dashed ${
                      dragOverOnboarding ? 'border-cyan-500 bg-cyan-500/10' : 'border-gray-700'
                    }`}
                  >
                    <p className="text-xs text-gray-500 uppercase tracking-wider mb-3">Onboarding</p>
                    {onboardingStaff.length === 0
                      ? <p className="text-xs text-gray-600">Drop here to onboard</p>
                      : <div className="flex flex-col gap-2">{onboardingStaff.map(s => staffTile(s))}</div>
                    }
                  </div>

                </div>
              )
            })()}

            </div>{/* end two-column layout */}
          </div>
        )}

        {/* Dashboard Tab */}
        {tab === 'dashboard' && (() => {
          const assignedStaffIds = new Set(assignments.map(a => a.staff_id))
          const availableStaff = staff.filter(s => !assignedStaffIds.has(s.id) && !s.flexed && !s.ooo && !s.onboarding)
          const oooCount = staff.filter(s => s.ooo).length
          const flexedCount = staff.filter(s => s.flexed).length
          const onboardingCount = staff.filter(s => s.onboarding).length

          const activeProjectIds = new Set(
            projects
              .filter(p => p.status === 'active' && assignments.some(a => a.project_id === p.id))
              .map(p => p.id)
          )
          const activeProjectCount = activeProjectIds.size

          const stoCount = assignments.filter(a => a.assignment_role === 'STO' && activeProjectIds.has(a.project_id)).length
          const avgSTOs = activeProjectCount ? (stoCount / activeProjectCount).toFixed(1) : '0'

          const opsCount = assignments.filter(a => a.assignment_role === 'Ops Support' && activeProjectIds.has(a.project_id)).length
          const avgOps = activeProjectCount ? (opsCount / activeProjectCount).toFixed(1) : '0'

          const supervisorMap = new Map<string, Set<string>>()
          assignments.filter(a => a.assignment_role === 'Supervisor').forEach(a => {
            if (!supervisorMap.has(a.staff_id)) supervisorMap.set(a.staff_id, new Set())
            supervisorMap.get(a.staff_id)!.add(a.project_id)
          })
          const avgProjectsPerSupervisor = supervisorMap.size
            ? (Array.from(supervisorMap.values()).reduce((sum, s) => sum + s.size, 0) / supervisorMap.size).toFixed(1)
            : '0'

          const today = new Date()
          const soon = new Date(); soon.setDate(today.getDate() + 30)
          const endingSoonProjects = projects.filter(p => {
            if (!p.end_date) return false
            const end = new Date(p.end_date)
            return end >= today && end <= soon
          })
          const becomingAvailable = staff.filter(s =>
            assignments.some(a => a.staff_id === s.id && endingSoonProjects.some(p => p.id === a.project_id))
          )

          const KEY_ROLES = ['Supervisor', 'STO', 'Ops Support']
          const understaffed = projects
            .filter(p => p.status === 'active')
            .map(p => {
              const projectAssignments = assignments.filter(a => a.project_id === p.id)
              const coveredRoles = new Set(projectAssignments.map(a => a.assignment_role).filter(Boolean))
              const missingRoles = KEY_ROLES.filter(r => !coveredRoles.has(r))
              const headcount = projectAssignments.length
              return { project: p, missingRoles, headcount, coveredRoles }
            })
            .filter(({ missingRoles, headcount }) => missingRoles.length > 0 || headcount === 0)
            .sort((a, b) => b.missingRoles.length - a.missingRoles.length)

          const statusCounts = projects.reduce((acc, p) => { acc[p.status] = (acc[p.status] || 0) + 1; return acc }, {} as Record<string, number>)
          const maxStatus = Math.max(...Object.values(statusCounts), 1)

          // Staffing-status distribution for active projects (same logic as Assignments tab)
          const staffingStatusCounts = { understaffed: 0, good: 0, overstaffed: 0 }
          projects.filter(p => p.status === 'active').forEach(p => {
            const roles = assignments.filter(a => a.project_id === p.id).map(a => a.assignment_role)
            const stoCount = roles.filter(r => r === 'STO').length
            const opsCount = roles.filter(r => r === 'Ops Support').length
            const status =
              stoCount === 0 || (stoCount >= 1 && opsCount === 0) ? 'understaffed' :
              stoCount >= 2 || opsCount > 2 ? 'overstaffed' :
              'good'
            staffingStatusCounts[status]++
          })
          const totalActiveForStaffing = staffingStatusCounts.understaffed + staffingStatusCounts.good + staffingStatusCounts.overstaffed

          const customerCounts = projects.reduce((acc, p) => {
            const key = p.customer_codename || 'No Codename'
            acc[key] = (acc[key] || 0) + 1
            return acc
          }, {} as Record<string, number>)
          const maxCustomer = Math.max(...Object.values(customerCounts), 1)

          const headcounts = projects.map(p => {
            const pa = assignments.filter(a => a.project_id === p.id)
            return {
              name: p.name,
              count: pa.length,
              roles: {
                Supervisor: pa.filter(a => a.assignment_role === 'Supervisor').length,
                STO: pa.filter(a => a.assignment_role === 'STO').length,
                'Ops Support': pa.filter(a => a.assignment_role === 'Ops Support').length,
              },
            }
          }).sort((a, b) => b.count - a.count)
          const maxHeadcount = Math.max(...headcounts.map(h => h.count), 1)

          const statCard = (label: string, value: string | number, sub?: string) => (
            <div className="bg-gray-900 border border-gray-800 rounded-xl p-5">
              <p className="text-xs text-gray-500 uppercase tracking-wider mb-1">{label}</p>
              <p className="text-3xl font-bold text-gray-100">{value}</p>
              {sub && <p className="text-xs text-gray-600 mt-1">{sub}</p>}
            </div>
          )

          const barChart = (title: string, data: Record<string, number>, max: number, colorClass: string) => (
            <div className="bg-gray-900 border border-gray-800 rounded-xl p-5">
              <p className="text-xs text-gray-500 uppercase tracking-wider mb-4">{title}</p>
              <div className="space-y-3">
                {Object.entries(data).sort((a, b) => b[1] - a[1]).map(([label, count]) => (
                  <div key={label}>
                    <div className="flex justify-between text-xs mb-1">
                      <span className="text-gray-300">{label}</span>
                      <span className="text-gray-500">{count}</span>
                    </div>
                    <div className="w-full bg-gray-800 rounded-full h-2">
                      <div className={`h-2 rounded-full ${colorClass}`} style={{ width: `${(count / max) * 100}%` }} />
                    </div>
                  </div>
                ))}
              </div>
            </div>
          )

          return (
            <div className="space-y-6">
              {/* Stat cards — averages */}
              <div className="grid grid-cols-2 gap-4 sm:grid-cols-3">
                {statCard('Avg STOs / Project', avgSTOs, 'active, staffed projects')}
                {statCard('Avg Ops Support / Project', avgOps, 'active, staffed projects')}
                {statCard('Avg Projects / Supervisor', avgProjectsPerSupervisor)}
              </div>

              {/* Stat cards — staff availability */}
              <div className="grid grid-cols-2 gap-4 sm:grid-cols-4">
                {statCard('Available Staff', availableStaff.length, `${staff.length} total`)}
                {statCard('Currently OOO', oooCount)}
                {statCard('Flexed', flexedCount)}
                {statCard('Onboarding', onboardingCount)}
              </div>

              {/* Charts row */}
              <div className="grid grid-cols-2 gap-4">
                {barChart('Projects by Status', statusCounts, maxStatus, 'bg-[#193a29]')}
                {barChart('Projects by Customer', customerCounts, maxCustomer, 'bg-violet-500')}
              </div>

              {/* Staffing status distribution (active projects) */}
              <div className="bg-gray-900 border border-gray-800 rounded-xl p-5">
                <p className="text-xs text-gray-500 uppercase tracking-wider mb-4">Staffing Status <span className="normal-case text-gray-600 ml-1">— active projects</span></p>
                {totalActiveForStaffing === 0 ? (
                  <p className="text-gray-600 text-sm">No active projects.</p>
                ) : (
                  <>
                    <div className="flex h-3 rounded-full overflow-hidden mb-4">
                      {staffingStatusCounts.understaffed > 0 && <div className="bg-amber-400" style={{ width: `${(staffingStatusCounts.understaffed / totalActiveForStaffing) * 100}%` }} />}
                      {staffingStatusCounts.good > 0 && <div className="bg-emerald-500" style={{ width: `${(staffingStatusCounts.good / totalActiveForStaffing) * 100}%` }} />}
                      {staffingStatusCounts.overstaffed > 0 && <div className="bg-blue-500" style={{ width: `${(staffingStatusCounts.overstaffed / totalActiveForStaffing) * 100}%` }} />}
                    </div>
                    <div className="grid grid-cols-3 gap-4">
                      {[
                        { key: 'understaffed', label: 'Potentially understaffed', dot: 'bg-amber-400', text: 'text-amber-400', count: staffingStatusCounts.understaffed },
                        { key: 'good', label: 'Well staffed', dot: 'bg-emerald-500', text: 'text-emerald-400', count: staffingStatusCounts.good },
                        { key: 'overstaffed', label: 'Potentially overstaffed', dot: 'bg-blue-500', text: 'text-blue-400', count: staffingStatusCounts.overstaffed },
                      ].map(({ key, label, dot, text, count }) => (
                        <div key={key}>
                          <div className="flex items-center gap-2 mb-1">
                            <span className={`w-2 h-2 rounded-full ${dot}`} />
                            <span className="text-xs text-gray-400">{label}</span>
                          </div>
                          <p className={`text-2xl font-bold ${text}`}>{count}</p>
                          <p className="text-xs text-gray-600">{Math.round((count / totalActiveForStaffing) * 100)}% of active</p>
                        </div>
                      ))}
                    </div>
                  </>
                )}
              </div>

              {/* Headcount per project */}
              <div className="bg-gray-900 border border-gray-800 rounded-xl p-5">
                <div className="flex items-center justify-between mb-4">
                  <p className="text-xs text-gray-500 uppercase tracking-wider">Total Headcount per Project</p>
                  <button
                    onClick={() => setShowSupervisorsInChart(v => !v)}
                    className={`text-xs px-3 py-1.5 rounded-lg border transition-colors ${
                      showSupervisorsInChart ? 'text-emerald-400' : 'border-gray-700 text-gray-400 hover:text-gray-200'
                    }`}
                    style={showSupervisorsInChart ? { borderColor: '#193a29' } : {}}
                  >
                    {showSupervisorsInChart ? 'Supervisors: on' : 'Supervisors: off'}
                  </button>
                </div>
                {headcounts.length === 0 ? (
                  <p className="text-gray-600 text-sm">No assignments yet.</p>
                ) : (
                  <div className="space-y-3">
                    {headcounts.map(({ name, count, roles }) => {
                      const displayCount = showSupervisorsInChart ? count : count - roles.Supervisor
                      const displayMax = showSupervisorsInChart ? maxHeadcount : Math.max(...headcounts.map(h => h.count - h.roles.Supervisor), 1)
                      return (
                      <div key={name}>
                        <div className="flex justify-between text-xs mb-1">
                          <span className="text-gray-300">{name}</span>
                          <div className="flex items-center gap-3">
                            {showSupervisorsInChart && roles.Supervisor > 0 && <span className="text-violet-400">{roles.Supervisor} Supervisor{roles.Supervisor > 1 ? 's' : ''}</span>}
                            {roles.STO > 0 && <span className="text-blue-400">{roles.STO} STO{roles.STO > 1 ? 's' : ''}</span>}
                            {roles['Ops Support'] > 0 && <span className="text-emerald-400">{roles['Ops Support']} Ops Support</span>}
                            <span className="text-gray-500">{displayCount} {displayCount === 1 ? 'person' : 'people'}</span>
                          </div>
                        </div>
                        <div className="w-full bg-gray-800 rounded-full h-2">
                          <div className="h-2 rounded-full bg-emerald-600" style={{ width: `${(displayCount / displayMax) * 100}%` }} />
                        </div>
                      </div>
                    )})}

                  </div>
                )}
              </div>

              {/* Becoming available soon */}
              <div className="bg-gray-900 border border-gray-800 rounded-xl p-5">
                <p className="text-xs text-gray-500 uppercase tracking-wider mb-4">Becoming Available Soon <span className="normal-case">(next 30 days)</span></p>
                {becomingAvailable.length === 0 ? (
                  <p className="text-gray-600 text-sm">No one becoming available in the next 30 days.</p>
                ) : (
                  <div className="space-y-2">
                    {becomingAvailable.map(s => {
                      const memberAssignments = assignments.filter(a => a.staff_id === s.id && endingSoonProjects.some(p => p.id === a.project_id))
                      return (
                        <div key={s.id} className="flex items-center justify-between bg-gray-800/60 rounded-lg px-4 py-2.5 text-sm">
                          <div className="flex items-center gap-3">
                            <span className="font-medium text-gray-200">{s.name}</span>
                            {s.position && <span className="text-xs text-gray-500">{s.position}</span>}
                          </div>
                          <div className="flex gap-2">
                            {memberAssignments.map(a => {
                              const p = projects.find(pr => pr.id === a.project_id)
                              return p ? (
                                <span key={a.id} className="text-xs px-2 py-0.5 rounded-full bg-amber-500/10 text-amber-400">
                                  {p.name} · ends {p.end_date}
                                </span>
                              ) : null
                            })}
                          </div>
                        </div>
                      )
                    })}
                  </div>
                )}
              </div>

              {/* Unassigned staff */}
              <div className="bg-gray-900 border border-gray-800 rounded-xl p-5">
                <div className="flex items-center justify-between mb-4">
                  <p className="text-xs text-gray-500 uppercase tracking-wider">People Without Assignment</p>
                  <span className={`text-xs font-semibold px-2 py-0.5 rounded-full ${availableStaff.length > 0 ? 'bg-amber-500/10 text-amber-400' : 'bg-emerald-500/10 text-emerald-400'}`}>
                    {availableStaff.length} {availableStaff.length === 1 ? 'person' : 'people'}
                  </span>
                </div>
                {availableStaff.length === 0 ? (
                  <p className="text-gray-600 text-sm">Everyone is currently assigned to a project.</p>
                ) : (
                  <div className="space-y-2">
                    {[...availableStaff].sort((a, b) => a.name.localeCompare(b.name)).map(s => (
                      <div key={s.id} className="flex items-center justify-between bg-gray-800/60 rounded-lg px-4 py-2.5 text-sm">
                        <div className="flex items-center gap-3">
                          <span className="font-medium text-gray-200">{s.name}</span>
                          {s.position && <span className="text-xs text-gray-500">{s.position}</span>}
                        </div>
                        {s.ooo && (
                          <div className="flex items-center gap-2">
                            <span className="text-xs px-2 py-0.5 rounded-full bg-amber-500/10 text-amber-400 font-medium">OOO</span>
                            {s.ooo_return_date && <span className="text-xs text-gray-500">Back {s.ooo_return_date}</span>}
                          </div>
                        )}
                      </div>
                    ))}
                  </div>
                )}
              </div>

              {/* Understaffed projects */}
              <div className="bg-gray-900 border border-gray-800 rounded-xl p-5">
                <div className="flex items-center justify-between mb-4">
                  <p className="text-xs text-gray-500 uppercase tracking-wider">Projects Understaffed</p>
                  <span className={`text-xs font-semibold px-2 py-0.5 rounded-full ${understaffed.length > 0 ? 'bg-red-500/10 text-red-400' : 'bg-emerald-500/10 text-emerald-400'}`}>
                    {understaffed.length} {understaffed.length === 1 ? 'project' : 'projects'}
                  </span>
                </div>
                {understaffed.length === 0 ? (
                  <p className="text-gray-600 text-sm">All active projects are fully staffed.</p>
                ) : (
                  <div className="space-y-3">
                    {understaffed.map(({ project: p, missingRoles, headcount }) => (
                      <div key={p.id} className="border border-gray-800 rounded-lg px-4 py-3">
                        <div className="flex items-center justify-between mb-2">
                          <div className="flex items-center gap-2">
                            <span className="font-medium text-gray-100 text-sm">{p.name}</span>
                            {p.customer_codename && <span className="text-xs text-gray-500">{p.customer_codename}</span>}
                          </div>
                          <div className="flex items-center gap-2">
                            {p.end_date && <span className="text-xs text-gray-500">ends {p.end_date}</span>}
                            <span className={`text-xs px-2 py-0.5 rounded-full font-medium ${
                              p.status === 'active' ? 'bg-emerald-500/10 text-emerald-400' : 'bg-amber-500/10 text-amber-400'
                            }`}>{p.status}</span>
                          </div>
                        </div>
                        <div className="flex items-center gap-3 flex-wrap">
                          <span className="text-xs text-gray-500">{headcount} {headcount === 1 ? 'person' : 'people'} assigned</span>
                          {headcount === 0 && (
                            <span className="text-xs px-2 py-0.5 rounded-full bg-red-500/10 text-red-400 font-medium">No staff</span>
                          )}
                          {missingRoles.map(r => (
                            <span key={r} className="text-xs px-2 py-0.5 rounded-full bg-gray-800 text-gray-400">
                              Missing: {r}
                            </span>
                          ))}
                        </div>
                      </div>
                    ))}
                  </div>
                )}
              </div>
            </div>
          )
        })()}

        {/* Staffing Scenarios Tab */}
        {tab === 'scenarios' && (
          <div>
            <div className="bg-gray-900 rounded-xl p-5 mb-6 border border-gray-800">
              <h2 className="text-xs font-semibold text-gray-400 uppercase tracking-wider mb-4">Staffing Scenarios</h2>
              <div className="flex flex-wrap gap-2 items-center">
                <input
                  className={inputClass}
                  placeholder="New scenario name…"
                  value={newScenarioName}
                  onChange={e => setNewScenarioName(e.target.value)}
                  onKeyDown={e => e.key === 'Enter' && addScenario()}
                />
                <label className="flex items-center gap-2 text-sm text-gray-400 cursor-pointer select-none">
                  <input type="checkbox" checked={newScenarioCopyCurrent} onChange={e => setNewScenarioCopyCurrent(e.target.checked)} className="accent-[#193a29] w-4 h-4" />
                  Start from current assignments
                </label>
                <button className={btnPrimary} style={{ backgroundColor: '#193a29' }} onClick={addScenario}>Create</button>
              </div>
              {scenarios.length > 0 && (
                <div className="flex flex-wrap gap-2 mt-4">
                  {scenarios.map(sc => {
                    const count = scenarioAssignments.filter(sa => sa.scenario_id === sc.id).length
                    const isSelected = selectedScenarioId === sc.id
                    return (
                      <div
                        key={sc.id}
                        onClick={() => setSelectedScenarioId(sc.id)}
                        className={`flex items-center gap-2 px-3 py-1.5 rounded-lg border text-sm cursor-pointer transition-colors ${isSelected ? 'text-white' : 'border-gray-700 text-gray-300 hover:text-white'}`}
                        style={isSelected ? { backgroundColor: '#193a29', borderColor: '#193a29' } : {}}
                      >
                        <span className="font-medium">{sc.name}</span>
                        <span className="text-[10px] bg-black/20 rounded px-1.5 py-0.5">{count}</span>
                        <button onClick={e => { e.stopPropagation(); deleteScenario(sc.id) }} className="text-gray-400 hover:text-red-400 leading-none">×</button>
                      </div>
                    )
                  })}
                </div>
              )}
            </div>

            {!selectedScenarioId ? (
              <p className="text-gray-600 text-sm">
                {scenarios.length === 0 ? 'No scenarios yet. Create one above to start planning a staffing version.' : 'Select a scenario above to view and edit its staffing.'}
              </p>
            ) : (() => {
              const sas = scenarioAssignments.filter(sa => sa.scenario_id === selectedScenarioId)
              const scenarioName = scenarios.find(s => s.id === selectedScenarioId)?.name
              return (
                <div>
                  <div className="bg-gray-900 rounded-xl p-5 mb-6 border border-gray-800">
                    <div className="flex items-center justify-between mb-4">
                      <h2 className="text-xs font-semibold text-gray-400 uppercase tracking-wider">Plan: <span className="text-gray-200 normal-case">{scenarioName}</span></h2>
                      <button
                        onClick={() => setConfirmApplyScenario(selectedScenarioId)}
                        className="text-xs px-3 py-1.5 rounded-lg border border-gray-700 text-gray-400 hover:text-gray-200 transition-colors"
                      >
                        Apply to live assignments
                      </button>
                    </div>
                    <div className="flex flex-wrap gap-2">
                      <select className={selectClass} value={scenarioAssign.project_id} onChange={e => setScenarioAssign({ ...scenarioAssign, project_id: e.target.value })}>
                        <option value="">Select project *</option>
                        {[...projects].filter(p => visibleStatuses[p.status]).sort((a, b) => statusRank(a.status) - statusRank(b.status) || a.name.localeCompare(b.name)).map(p => <option key={p.id} value={p.id}>{p.name}</option>)}
                      </select>
                      <select className={selectClass} value={scenarioAssign.staff_id} onChange={e => setScenarioAssign({ ...scenarioAssign, staff_id: e.target.value })}>
                        <option value="">Select staff *</option>
                        {[...staff].sort((a, b) => a.name.localeCompare(b.name)).map(s => <option key={s.id} value={s.id}>{s.name}</option>)}
                      </select>
                      <select className={selectClass} value={scenarioAssign.assignment_role} onChange={e => setScenarioAssign({ ...scenarioAssign, assignment_role: e.target.value })}>
                        <option value="">Role</option>
                        <option>Supervisor</option>
                        <option>STO</option>
                        <option>Ops Support</option>
                      </select>
                      <button className={btnPrimary} style={{ backgroundColor: '#193a29' }} onClick={addScenarioAssignment}>Assign</button>
                    </div>
                  </div>

                  <div className="flex gap-6 items-start">
                    {/* Left: scenario project cards */}
                    <div className="flex-1 min-w-0 space-y-3">
                    {[...projects].filter(p => visibleStatuses[p.status]).sort((a, b) => statusRank(a.status) - statusRank(b.status) || a.name.localeCompare(b.name)).map(p => {
                      const pa = sas.filter(sa => sa.project_id === p.id)
                      const isOver = dragOverScenarioProjectId === p.id
                      return (
                        <div
                          key={p.id}
                          className={`border rounded-xl p-5 transition-all ${isOver ? 'border-[#193a29] bg-[#193a29]/10 scale-[1.01]' : 'border-gray-800 bg-gray-900/40'}`}
                          onDragOver={e => { e.preventDefault(); setDragOverScenarioProjectId(p.id) }}
                          onDragLeave={() => setDragOverScenarioProjectId(null)}
                          onDrop={async e => {
                            e.preventDefault(); setDragOverScenarioProjectId(null)
                            const type = e.dataTransfer.getData('type')
                            const id = e.dataTransfer.getData('id')
                            if (type === 'scenario-staff' && id) {
                              if (!sas.some(sa => sa.project_id === p.id && sa.staff_id === id)) setPendingScenarioDrop({ staffId: id, projectId: p.id })
                              setDraggedStaffId(null)
                            } else if (type === 'scenario-assignment' && id) {
                              const sa = scenarioAssignments.find(x => x.id === id)
                              if (sa && sa.project_id !== p.id && !sas.some(x => x.project_id === p.id && x.staff_id === sa.staff_id)) {
                                await moveScenarioAssignment(id, p.id)
                              }
                              setDraggedAssignmentId(null)
                            }
                          }}
                        >
                          <div className="flex items-center justify-between mb-4">
                            <div className="flex items-center gap-2">
                              <h3 className="font-semibold text-gray-100">{p.name}</h3>
                              <button
                                onClick={() => { setAddingToScenarioProjectId(addingToScenarioProjectId === p.id ? null : p.id); setScenarioQuickAdd({ staff_id: '', assignment_role: '' }) }}
                                title="Add a person to this project"
                                className="w-5 h-5 flex items-center justify-center rounded-full border border-gray-600 text-gray-400 hover:text-white hover:border-gray-400 transition-colors text-sm leading-none"
                              >
                                +
                              </button>
                            </div>
                            <span className={`px-2.5 py-1 rounded-full text-xs font-medium ${
                              p.status === 'active' ? 'bg-emerald-500/10 text-emerald-400' :
                              p.status === 'starting-soon' ? 'bg-sky-500/10 text-sky-400' :
                              p.status === 'on-hold' ? 'bg-amber-500/10 text-amber-400' :
                              p.status === 'paused' ? 'bg-rose-500/10 text-rose-400' :
                              'bg-gray-700 text-gray-400'
                            }`}>{p.status}</span>
                          </div>
                          {addingToScenarioProjectId === p.id && (
                            <div className="flex flex-wrap gap-2 mb-3 p-3 rounded-lg bg-gray-800/40 border border-gray-700">
                              <select
                                className={selectSmClass + ' flex-1 min-w-[160px]'}
                                value={scenarioQuickAdd.staff_id}
                                onChange={e => setScenarioQuickAdd({ ...scenarioQuickAdd, staff_id: e.target.value })}
                              >
                                <option value="">Select staff…</option>
                                {[...staff]
                                  .filter(s => !pa.some(sa => sa.staff_id === s.id))
                                  .sort((a, b) => a.name.localeCompare(b.name))
                                  .map(s => <option key={s.id} value={s.id}>{s.name}{s.position ? ` — ${s.position}` : ''}</option>)}
                              </select>
                              <select
                                className={selectSmClass + ' w-32'}
                                value={scenarioQuickAdd.assignment_role}
                                onChange={e => setScenarioQuickAdd({ ...scenarioQuickAdd, assignment_role: e.target.value })}
                              >
                                <option value="">Role</option>
                                <option>Supervisor</option>
                                <option>STO</option>
                                <option>Ops Support</option>
                              </select>
                              <button
                                className="text-xs px-2 py-1 rounded bg-gray-700 text-gray-300 hover:text-white transition-colors whitespace-nowrap"
                                onClick={async () => {
                                  if (!scenarioQuickAdd.staff_id || !selectedScenarioId) return
                                  await addScenarioAssignmentDirect(selectedScenarioId, p.id, scenarioQuickAdd.staff_id, scenarioQuickAdd.assignment_role)
                                  setScenarioQuickAdd({ staff_id: '', assignment_role: '' })
                                  setAddingToScenarioProjectId(null)
                                }}
                              >+ Add</button>
                            </div>
                          )}
                          {pa.length === 0 ? (
                            <p className={`text-sm ${isOver ? 'text-gray-400' : 'text-gray-600'}`}>{isOver ? 'Drop to add' : 'No staff in this scenario.'}</p>
                          ) : (
                            <div className="space-y-2">
                              {[...pa].sort((a, b) => {
                                const order = ['Supervisor', 'STO', 'Ops Support']
                                return (order.indexOf(a.assignment_role ?? '') === -1 ? 99 : order.indexOf(a.assignment_role ?? '')) -
                                       (order.indexOf(b.assignment_role ?? '') === -1 ? 99 : order.indexOf(b.assignment_role ?? ''))
                              }).map(sa => {
                                const member = staff.find(s => s.id === sa.staff_id)
                                return (
                                  <div
                                    key={sa.id}
                                    draggable
                                    onDragStart={e => { e.dataTransfer.setData('type', 'scenario-assignment'); e.dataTransfer.setData('id', sa.id); setDraggedAssignmentId(sa.id); setDraggedStaffId(null) }}
                                    onDragEnd={() => setDraggedAssignmentId(null)}
                                    className={`flex items-center justify-between bg-gray-800/60 rounded-lg px-4 py-2.5 text-sm cursor-grab active:cursor-grabbing transition-opacity ${draggedAssignmentId === sa.id ? 'opacity-40' : ''}`}
                                  >
                                    <div className="flex items-center gap-3">
                                      <span className="font-medium text-gray-200">{member?.name ?? 'Unknown'}</span>
                                      <select
                                        value={sa.assignment_role ?? ''}
                                        onChange={e => updateScenarioAssignmentRole(sa.id, e.target.value)}
                                        title="Change role"
                                        className={`text-xs font-medium px-2 py-0.5 pr-5 rounded-full cursor-pointer appearance-none focus:outline-none focus:ring-1 focus:ring-gray-500 ${sa.assignment_role ? roleColor(sa.assignment_role) : 'bg-gray-700/60 text-gray-400'}`}
                                        style={{ backgroundImage: 'none' }}
                                      >
                                        <option value="" className="bg-gray-900 text-gray-100">No role</option>
                                        <option value="Supervisor" className="bg-gray-900 text-gray-100">Supervisor</option>
                                        <option value="STO" className="bg-gray-900 text-gray-100">STO</option>
                                        <option value="Ops Support" className="bg-gray-900 text-gray-100">Ops Support</option>
                                      </select>
                                      {member?.position && <span className="text-xs text-gray-500">{member.position}</span>}
                                    </div>
                                    <button className={btnDanger} onClick={() => removeScenarioAssignment(sa.id)}>×</button>
                                  </div>
                                )
                              })}
                            </div>
                          )}
                        </div>
                      )
                    })}
                    </div>{/* end left column */}

                    {/* Right: staff pool sidebar (not in this scenario) */}
                    {(() => {
                      // Only count someone as "in scenario" if their assignment is to a currently-visible project.
                      // This frees up people whose only scenario assignment is to a hidden/completed project.
                      const visibleProjectIds = new Set(projects.filter(p => visibleStatuses[p.status]).map(p => p.id))
                      const inScenario = new Set(sas.filter(sa => visibleProjectIds.has(sa.project_id)).map(sa => sa.staff_id))
                      const sUnassigned = [...staff].filter(s => !inScenario.has(s.id) && !s.flexed && !s.ooo && !s.onboarding).sort((a, b) => a.name.localeCompare(b.name))
                      const sFlexed = [...staff].filter(s => s.flexed && !inScenario.has(s.id)).sort((a, b) => a.name.localeCompare(b.name))
                      const sOOO = [...staff].filter(s => s.ooo && !inScenario.has(s.id)).sort((a, b) => a.name.localeCompare(b.name))
                      const sOnboarding = [...staff].filter(s => s.onboarding && !inScenario.has(s.id)).sort((a, b) => a.name.localeCompare(b.name))
                      const tile = (s: Staff) => (
                        <div
                          key={s.id}
                          draggable
                          onDragStart={e => { e.dataTransfer.setData('type', 'scenario-staff'); e.dataTransfer.setData('id', s.id); setDraggedStaffId(s.id); setDraggedAssignmentId(null) }}
                          onDragEnd={() => setDraggedStaffId(null)}
                          className={`cursor-grab active:cursor-grabbing select-none border rounded-lg px-3 py-2 text-sm transition-all ${draggedStaffId === s.id ? 'opacity-40 border-gray-600' : 'border-gray-700 bg-gray-900 hover:border-gray-500'}`}
                        >
                          <p className="font-medium text-gray-200 leading-tight">{s.name}</p>
                          {s.position && <p className="text-xs text-gray-500 mt-0.5">{s.position}</p>}
                          {s.ooo && <span className="text-[10px] px-1.5 py-0.5 rounded-full bg-amber-500/10 text-amber-400 mt-1 inline-block">OOO</span>}
                  {s.flexed && s.flex_notes && <p className="text-[10px] text-violet-400 mt-1 leading-snug">📝 {s.flex_notes}</p>}
                        </div>
                      )
                      const zoneDrop = (zone: 'unassigned' | 'flexed' | 'ooo' | 'onboarding') => async (e: React.DragEvent) => {
                        e.preventDefault()
                        setDragOverUnassigned(false); setDragOverFlexed(false); setDragOverOOO(false); setDragOverOnboarding(false)
                        const type = e.dataTransfer.getData('type')
                        const id = e.dataTransfer.getData('id')
                        let staffId = ''
                        if (type === 'scenario-staff' && id) { staffId = id; setDraggedStaffId(null) }
                        else if (type === 'scenario-assignment' && id) {
                          const sa = scenarioAssignments.find(x => x.id === id)
                          await removeScenarioAssignment(id)
                          if (sa) staffId = sa.staff_id
                          setDraggedAssignmentId(null)
                        }
                        if (!staffId) return
                        if (zone === 'ooo') {
                          const m = staff.find(s => s.id === staffId)
                          setOooDate(m?.ooo_return_date ?? '')
                          setPendingOOO({ staffId, assignmentId: null })
                        } else {
                          await setStaffZone(staffId, zone)
                        }
                      }
                      return (
                        <div className="w-52 shrink-0 sticky top-6 flex flex-col gap-5">
                          <div
                            onDragOver={e => { e.preventDefault(); setDragOverUnassigned(true) }}
                            onDragLeave={() => setDragOverUnassigned(false)}
                            onDrop={zoneDrop('unassigned')}
                            className={`rounded-lg p-2 min-h-[60px] transition-all ${dragOverUnassigned ? 'ring-1 ring-gray-500 bg-gray-800/40' : ''}`}
                          >
                            <p className="text-xs text-gray-500 uppercase tracking-wider mb-3">Not in scenario</p>
                            {sUnassigned.length === 0 ? <p className="text-xs text-gray-600">None</p> : <div className="flex flex-col gap-2">{sUnassigned.map(tile)}</div>}
                          </div>
                          <div
                            onDragOver={e => { e.preventDefault(); setDragOverFlexed(true) }}
                            onDragLeave={() => setDragOverFlexed(false)}
                            onDrop={zoneDrop('flexed')}
                            className={`rounded-lg p-2 min-h-[60px] border border-dashed transition-all ${dragOverFlexed ? 'border-violet-500 bg-violet-500/10' : 'border-gray-700'}`}
                          >
                            <p className="text-xs text-gray-500 uppercase tracking-wider mb-3">Flexed</p>
                            {sFlexed.length === 0 ? <p className="text-xs text-gray-600">Drop here to flex</p> : <div className="flex flex-col gap-2">{sFlexed.map(tile)}</div>}
                          </div>
                          <div
                            onDragOver={e => { e.preventDefault(); setDragOverOOO(true) }}
                            onDragLeave={() => setDragOverOOO(false)}
                            onDrop={zoneDrop('ooo')}
                            className={`rounded-lg p-2 min-h-[60px] border border-dashed transition-all ${dragOverOOO ? 'border-amber-500 bg-amber-500/10' : 'border-gray-700'}`}
                          >
                            <p className="text-xs text-gray-500 uppercase tracking-wider mb-3">Currently OOO</p>
                            {sOOO.length === 0 ? <p className="text-xs text-gray-600">Drop here to mark OOO</p> : <div className="flex flex-col gap-2">{sOOO.map(tile)}</div>}
                          </div>
                          <div
                            onDragOver={e => { e.preventDefault(); setDragOverOnboarding(true) }}
                            onDragLeave={() => setDragOverOnboarding(false)}
                            onDrop={zoneDrop('onboarding')}
                            className={`rounded-lg p-2 min-h-[60px] border border-dashed transition-all ${dragOverOnboarding ? 'border-cyan-500 bg-cyan-500/10' : 'border-gray-700'}`}
                          >
                            <p className="text-xs text-gray-500 uppercase tracking-wider mb-3">Onboarding</p>
                            {sOnboarding.length === 0 ? <p className="text-xs text-gray-600">Drop here to onboard</p> : <div className="flex flex-col gap-2">{sOnboarding.map(tile)}</div>}
                          </div>
                        </div>
                      )
                    })()}
                  </div>
                </div>
              )
            })()}
          </div>
        )}

        {/* Milestones Tab */}
        {tab === 'milestones' && (() => {
          const today = fmt(new Date())
          const activeProjects = [...projects].filter(p => p.status === 'active').sort((a, b) => a.name.localeCompare(b.name))
          const activeIds = new Set(activeProjects.map(p => p.id))
          const allMs = milestones.filter(m => activeIds.has(m.project_id))
          const total = allMs.length
          const doneCount = allMs.filter(m => m.done).length
          const pct = total ? Math.round((doneCount / total) * 100) : 0
          const openP0 = allMs.filter(m => !m.done && m.priority === 'P0')
          const openCounts = {
            P0: allMs.filter(m => !m.done && m.priority === 'P0').length,
            P1: allMs.filter(m => !m.done && m.priority === 'P1').length,
            P2: allMs.filter(m => !m.done && m.priority === 'P2').length,
          }
          return (
            <div className="flex gap-6 items-start">
              {/* Left: per-project milestone checklists */}
              <div className="flex-1 min-w-0 space-y-3">
                <div className="flex justify-end gap-2">
                  <button
                    onClick={exportMilestonesHTML}
                    className="text-xs px-3 py-1.5 rounded-lg border border-gray-700 text-gray-400 hover:text-gray-200 transition-colors"
                  >
                    Download as HTML
                  </button>
                  <button
                    onClick={exportMilestonesPDF}
                    className="text-xs px-3 py-1.5 rounded-lg border border-gray-700 text-gray-400 hover:text-gray-200 transition-colors"
                  >
                    Download as PDF
                  </button>
                  <button
                    onClick={() => setHideEmptyMilestoneProjects(v => !v)}
                    className={`text-xs px-3 py-1.5 rounded-lg border transition-colors ${
                      hideEmptyMilestoneProjects ? 'text-emerald-400' : 'border-gray-700 text-gray-400 hover:text-gray-200'
                    }`}
                    style={hideEmptyMilestoneProjects ? { borderColor: '#193a29' } : {}}
                  >
                    {hideEmptyMilestoneProjects ? 'Hiding projects without milestones' : 'Show all projects'}
                  </button>
                </div>
                {(() => {
                  const shownProjects = hideEmptyMilestoneProjects
                    ? activeProjects.filter(p => milestones.some(m => m.project_id === p.id))
                    : activeProjects
                  return shownProjects.length === 0 ? (
                  <p className="text-gray-600 text-sm">{activeProjects.length === 0 ? 'No active projects. Milestones are tracked for active projects only.' : 'No active projects have milestones yet.'}</p>
                ) : shownProjects.map(p => {
                  const ms = [...milestones.filter(m => m.project_id === p.id)].sort((a, b) =>
                    (a.done === b.done ? 0 : a.done ? 1 : -1) || priorityRank(a.priority) - priorityRank(b.priority) || a.created_at.localeCompare(b.created_at)
                  )
                  const pDone = ms.filter(m => m.done).length
                  const draft = getMilestoneDraft(p.id)
                  return (
                    <div key={p.id} className="border border-gray-800 rounded-xl p-5 bg-gray-900/40">
                      <div className="flex items-center justify-between mb-4">
                        <h3 className="font-semibold text-gray-100">{p.name}</h3>
                        {ms.length > 0 && <span className="text-xs text-gray-500">{pDone}/{ms.length} done</span>}
                      </div>
                      {ms.length > 0 && (
                        <div className="space-y-1.5 mb-3">
                          {ms.map(m => {
                            const overdue = !m.done && m.due_date && m.due_date < today
                            return (
                              <div key={m.id} className="flex items-center gap-3 bg-gray-800/40 rounded-lg px-3 py-2 text-sm">
                                {editingMilestoneId === m.id ? (
                                  <>
                                    <input
                                      className={inputSmClass + ' flex-1'}
                                      value={editMilestone.title}
                                      autoFocus
                                      onChange={e => setEditMilestone({ ...editMilestone, title: e.target.value })}
                                      onKeyDown={e => e.key === 'Enter' && saveMilestone(m.id)}
                                    />
                                    <input type="date" className={inputSmClass + ' w-40'} value={editMilestone.due_date} onChange={e => setEditMilestone({ ...editMilestone, due_date: e.target.value })} />
                                    <button className={btnSave} onClick={() => saveMilestone(m.id)}>Save</button>
                                    <button className={btnCancel} onClick={() => setEditingMilestoneId(null)}>Cancel</button>
                                  </>
                                ) : (
                                  <>
                                    <input type="checkbox" checked={m.done} onChange={e => toggleMilestone(m.id, e.target.checked)} className="accent-[#193a29] w-4 h-4 shrink-0" />
                                    <span
                                      onClick={() => startEditMilestone(m)}
                                      title="Click to edit"
                                      className={`flex-1 cursor-pointer hover:underline decoration-dotted underline-offset-4 ${m.done ? 'line-through text-gray-600' : 'text-gray-200'}`}
                                    >{m.title}</span>
                                    {m.due_date && <span className={`text-xs ${overdue ? 'text-red-400 font-medium' : 'text-gray-500'}`}>{overdue ? '⚠ ' : ''}{m.due_date}</span>}
                                    <select
                                      value={m.priority}
                                      onChange={e => updateMilestonePriority(m.id, e.target.value)}
                                      title="Priority"
                                      className={`text-xs font-medium px-2 py-0.5 pr-5 rounded-full cursor-pointer appearance-none focus:outline-none focus:ring-1 focus:ring-gray-500 ${priorityColor(m.priority)}`}
                                      style={{ backgroundImage: 'none' }}
                                    >
                                      {PRIORITIES.map(pr => <option key={pr} value={pr} className="bg-gray-900 text-gray-100">{pr}</option>)}
                                    </select>
                                    <button className={btnEdit} onClick={() => startEditMilestone(m)}>Edit</button>
                                    <button className={btnDanger} onClick={() => deleteMilestone(m.id)}>×</button>
                                  </>
                                )}
                              </div>
                            )
                          })}
                        </div>
                      )}
                      {addingMilestoneProjectId === p.id ? (
                        <div className="flex flex-wrap gap-2">
                          <input
                            className={inputSmClass + ' flex-1 min-w-[180px]'}
                            placeholder="New milestone…"
                            value={draft.title}
                            autoFocus
                            onChange={e => setMilestoneDraft(p.id, { title: e.target.value })}
                            onKeyDown={e => e.key === 'Enter' && addMilestone(p.id)}
                          />
                          <select className={selectSmClass + ' w-20'} value={draft.priority} onChange={e => setMilestoneDraft(p.id, { priority: e.target.value })}>
                            {PRIORITIES.map(pr => <option key={pr} value={pr}>{pr}</option>)}
                          </select>
                          <input type="date" className={inputSmClass + ' w-40'} value={draft.due_date} onChange={e => setMilestoneDraft(p.id, { due_date: e.target.value })} />
                          <button className="text-xs px-3 py-1 rounded bg-gray-700 text-gray-300 hover:text-white transition-colors" onClick={() => addMilestone(p.id)}>Add</button>
                          <button className={btnCancel} onClick={() => setAddingMilestoneProjectId(null)}>Cancel</button>
                        </div>
                      ) : (
                        <button
                          onClick={() => setAddingMilestoneProjectId(p.id)}
                          className="text-xs text-gray-500 hover:text-gray-300 transition-colors"
                        >
                          + Add new milestone
                        </button>
                      )}
                    </div>
                  )
                })
                })()}
              </div>

              {/* Right: stats + P0 reminders */}
              <div className="w-64 shrink-0 sticky top-6 flex flex-col gap-4">
                <div className="bg-gray-900 border border-gray-800 rounded-xl p-5">
                  <p className="text-xs text-gray-500 uppercase tracking-wider mb-3">Overall</p>
                  <p className="text-3xl font-bold text-gray-100">{pct}%</p>
                  <p className="text-xs text-gray-600 mb-3">{doneCount} of {total} done</p>
                  <div className="w-full bg-gray-800 rounded-full h-2">
                    <div className="h-2 rounded-full bg-emerald-600" style={{ width: `${pct}%` }} />
                  </div>
                  <div className="flex gap-3 mt-4 text-xs">
                    <span className="text-red-400">{openCounts.P0} P0</span>
                    <span className="text-amber-400">{openCounts.P1} P1</span>
                    <span className="text-sky-400">{openCounts.P2} P2</span>
                    <span className="text-gray-600">open</span>
                  </div>
                </div>

                <div className="bg-gray-900 border border-gray-800 rounded-xl p-5">
                  <div className="flex items-center justify-between mb-3">
                    <p className="text-xs text-gray-500 uppercase tracking-wider">P0 — needs attention</p>
                    <span className={`text-xs font-semibold px-2 py-0.5 rounded-full ${openP0.length > 0 ? 'bg-red-500/10 text-red-400' : 'bg-emerald-500/10 text-emerald-400'}`}>{openP0.length}</span>
                  </div>
                  {openP0.length === 0 ? (
                    <p className="text-xs text-gray-600">No open P0 milestones. 🎉</p>
                  ) : (
                    <div className="space-y-2">
                      {[...openP0].sort((a, b) => (a.due_date ?? '9999').localeCompare(b.due_date ?? '9999')).map(m => {
                        const proj = projects.find(p => p.id === m.project_id)
                        const overdue = m.due_date && m.due_date < today
                        return (
                          <div key={m.id} className="border border-gray-800 rounded-lg px-3 py-2">
                            <p className="text-sm text-gray-200 leading-tight">{m.title}</p>
                            <div className="flex items-center justify-between mt-1">
                              <span className="text-xs text-gray-500">{proj?.name}</span>
                              {m.due_date && <span className={`text-xs ${overdue ? 'text-red-400 font-medium' : 'text-gray-500'}`}>{overdue ? '⚠ overdue · ' : 'due '}{m.due_date}</span>}
                            </div>
                          </div>
                        )
                      })}
                    </div>
                  )}
                </div>
              </div>
            </div>
          )
        })()}
      </div>

      {pendingOOO && (() => {
        const member = staff.find(s => s.id === pendingOOO.staffId)
        return (
          <div className="fixed inset-0 bg-black/50 flex items-center justify-center z-50" onClick={() => { setPendingOOO(null); setOooDate('') }}>
            <div className="bg-gray-900 border border-gray-700 rounded-xl p-6 w-80 shadow-xl" onClick={e => e.stopPropagation()}>
              <p className="text-sm font-semibold text-gray-100 mb-1">Mark as OOO</p>
              <p className="text-xs text-gray-500 mb-4"><span className="text-gray-300">{member?.name}</span> will be marked out of office.</p>
              <p className="text-xs text-gray-500 mb-2">Expected return date (optional):</p>
              <input
                type="date"
                className={inputClass + ' w-full mb-4'}
                value={oooDate}
                onChange={e => setOooDate(e.target.value)}
                autoFocus
              />
              <div className="flex gap-2 justify-end">
                <button className={btnCancel} onClick={() => { setPendingOOO(null); setOooDate('') }}>Cancel</button>
                <button
                  className={btnPrimary}
                  style={{ backgroundColor: '#193a29' }}
                  onClick={async () => {
                    const { data } = await supabase.from('staff').update({ ooo: true, flexed: false, onboarding: false, ooo_return_date: oooDate || null }).eq('id', pendingOOO.staffId).select().single()
                    if (data) setStaff(prev => prev.map(s => s.id === pendingOOO.staffId ? data : s))
                    if (pendingOOO.assignmentId) await removeAssignment(pendingOOO.assignmentId)
                    setPendingOOO(null)
                    setOooDate('')
                  }}
                >
                  Mark OOO
                </button>
              </div>
            </div>
          </div>
        )
      })()}

      {pendingScenarioDrop && selectedScenarioId && (
        <div className="fixed inset-0 bg-black/50 flex items-center justify-center z-50" onClick={() => { setPendingScenarioDrop(null); setDropRole('') }}>
          <div className="bg-gray-900 border border-gray-700 rounded-xl p-6 w-80 shadow-xl" onClick={e => e.stopPropagation()}>
            <p className="text-sm font-semibold text-gray-100 mb-1">Add to scenario</p>
            <p className="text-xs text-gray-500 mb-4">
              <span className="text-gray-300">{staff.find(s => s.id === pendingScenarioDrop.staffId)?.name}</span>
              {' → '}
              <span className="text-gray-300">{projects.find(p => p.id === pendingScenarioDrop.projectId)?.name}</span>
            </p>
            <select className={selectClass + ' w-full mb-4'} value={dropRole} onChange={e => setDropRole(e.target.value)} autoFocus>
              <option value="">No role</option>
              <option>Supervisor</option>
              <option>STO</option>
              <option>Ops Support</option>
            </select>
            <div className="flex gap-2 justify-end">
              <button className={btnCancel} onClick={() => { setPendingScenarioDrop(null); setDropRole('') }}>Cancel</button>
              <button
                className={btnPrimary}
                style={{ backgroundColor: '#193a29' }}
                onClick={async () => {
                  await addScenarioAssignmentDirect(selectedScenarioId, pendingScenarioDrop.projectId, pendingScenarioDrop.staffId, dropRole)
                  setPendingScenarioDrop(null)
                  setDropRole('')
                }}
              >
                Add
              </button>
            </div>
          </div>
        </div>
      )}

      {confirmApplyScenario && (
        <div className="fixed inset-0 bg-black/50 flex items-center justify-center z-50" onClick={() => setConfirmApplyScenario(null)}>
          <div className="bg-gray-900 border border-gray-700 rounded-xl p-6 w-96 shadow-xl" onClick={e => e.stopPropagation()}>
            <p className="text-sm font-semibold text-gray-100 mb-2">Apply scenario to live?</p>
            <p className="text-xs text-gray-400 mb-5">
              This will <span className="text-red-400 font-medium">replace all current live assignments</span> with the staffing from
              {' '}<span className="text-gray-200">{scenarios.find(s => s.id === confirmApplyScenario)?.name}</span>. This can't be undone.
            </p>
            <div className="flex gap-2 justify-end">
              <button className={btnCancel} onClick={() => setConfirmApplyScenario(null)}>Cancel</button>
              <button className={btnPrimary} style={{ backgroundColor: '#193a29' }} onClick={() => applyScenarioToLive(confirmApplyScenario)}>Apply to live</button>
            </div>
          </div>
        </div>
      )}

      {confirmFreeStaff && (
        <div className="fixed inset-0 bg-black/50 flex items-center justify-center z-50" onClick={() => setConfirmFreeStaff(null)}>
          <div className="bg-gray-900 border border-gray-700 rounded-xl p-6 w-96 shadow-xl" onClick={e => e.stopPropagation()}>
            <p className="text-sm font-semibold text-gray-100 mb-2">Project completed</p>
            <p className="text-xs text-gray-400 mb-5">
              <span className="text-gray-200">{projects.find(p => p.id === confirmFreeStaff.projectId)?.name}</span> has{' '}
              <span className="text-gray-200">{confirmFreeStaff.count}</span> {confirmFreeStaff.count === 1 ? 'person' : 'people'} assigned.
              Free them up by clearing their assignments?
            </p>
            <div className="flex gap-2 justify-end">
              <button className={btnCancel} onClick={() => setConfirmFreeStaff(null)}>Keep assignments</button>
              <button
                className={btnPrimary}
                style={{ backgroundColor: '#193a29' }}
                onClick={async () => {
                  await supabase.from('assignments').delete().eq('project_id', confirmFreeStaff.projectId)
                  setAssignments(prev => prev.filter(a => a.project_id !== confirmFreeStaff.projectId))
                  setConfirmFreeStaff(null)
                }}
              >
                Free up staff
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  )
}
