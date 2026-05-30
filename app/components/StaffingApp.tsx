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
}

type Staff = {
  id: string
  name: string
  position: string | null
  ooo: boolean
  ooo_return_date: string | null
  flexed: boolean
}

type Assignment = {
  id: string
  project_id: string
  staff_id: string
  assignment_role: string | null
}

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
  const [tab, setTab] = useState<'projects' | 'staff' | 'assignments' | 'dashboard'>('dashboard')
  const [projects, setProjects] = useState<Project[]>([])
  const [staff, setStaff] = useState<Staff[]>([])
  const [assignments, setAssignments] = useState<Assignment[]>([])
  const [loading, setLoading] = useState(true)
  const [theme, setTheme] = useState<'dark' | 'light' | 'system'>('dark')

  const [newProject, setNewProject] = useState({ name: '', customer_codename: '', status: 'active', duration_weeks: '' })
  const [newStaff, setNewStaff] = useState({ name: '', position: '', ooo: false, ooo_return_date: '' })
  const [newAssignment, setNewAssignment] = useState({ project_id: '', staff_id: '', assignment_role: '' })

  const [editingProjectId, setEditingProjectId] = useState<string | null>(null)
  const [editProject, setEditProject] = useState({ name: '', customer_codename: '', status: 'active', duration_weeks: '' })

  const [editingStaffId, setEditingStaffId] = useState<string | null>(null)
  const [editStaff, setEditStaff] = useState({ name: '', position: '', ooo: false, ooo_return_date: '' })
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
  const [draggedStaffId, setDraggedStaffId] = useState<string | null>(null)
  const [draggedAssignmentId, setDraggedAssignmentId] = useState<string | null>(null)
  const [dragOverProjectId, setDragOverProjectId] = useState<string | null>(null)
  const [dragOverFlexed, setDragOverFlexed] = useState(false)
  const [dragOverUnassigned, setDragOverUnassigned] = useState(false)
  const [dragOverOOO, setDragOverOOO] = useState(false)
  const [pendingDrop, setPendingDrop] = useState<{ staffId: string; projectId: string } | null>(null)
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
    const [p, s, a] = await Promise.all([
      supabase.from('projects').select('*').order('created_at', { ascending: false }),
      supabase.from('staff').select('*').order('created_at', { ascending: false }),
      supabase.from('assignments').select('*'),
    ])
    if (p.data) setProjects(p.data)
    if (s.data) setStaff(s.data)
    if (a.data) setAssignments(a.data)
    setLoading(false)
  }

  const fmt = (d: Date) => d.toISOString().split('T')[0]

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
    const { data } = await supabase.from('projects').update(updates).eq('id', id).select().single()
    if (data) {
      setProjects(projects.map(p => p.id === id ? data : p))
      setEditingProjectId(null)
    }
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
    setEditStaff({ name: s.name, position: s.position ?? '', ooo: s.ooo, ooo_return_date: s.ooo_return_date ?? '' })
    setEditingStaffId(s.id)
  }

  async function saveStaff(id: string) {
    if (!editStaff.name.trim()) return
    const updates = { name: editStaff.name.trim(), position: editStaff.position || null, ooo: editStaff.ooo, ooo_return_date: editStaff.ooo_return_date || null }
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

  async function setStaffZone(id: string, zone: 'unassigned' | 'flexed' | 'ooo') {
    const { data } = await supabase.from('staff').update({ flexed: zone === 'flexed', ooo: zone === 'ooo' }).eq('id', id).select().single()
    if (data) setStaff(prev => prev.map(s => s.id === id ? data : s))
  }

  async function removeAssignment(id: string) {
    await supabase.from('assignments').delete().eq('id', id)
    setAssignments(assignments.filter(a => a.id !== id))
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

  return (
    <div className="min-h-screen bg-gray-950 text-gray-100">
      <div className="border-b px-8 py-5 flex items-center justify-between" style={{ borderColor: '#193a29' }}>
        <h1 className="text-xl font-semibold text-gray-100 tracking-tight">Code Pod Staffing Manager</h1>
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
          {(['dashboard', 'projects', 'staff', 'assignments'] as const).map(t => (
            <button
              key={t}
              onClick={() => setTab(t)}
              className={`px-5 py-2 text-sm font-medium capitalize rounded-lg transition-all ${
                tab === t ? 'text-white shadow' : 'text-gray-400 hover:text-gray-200'
              }`}
              style={tab === t ? { backgroundColor: '#193a29' } : {}}
            >
              {t}
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
                              <span className={`px-2.5 py-1 rounded-full text-xs font-medium ${
                                p.status === 'active' ? 'bg-emerald-500/10 text-emerald-400' :
                                p.status === 'starting-soon' ? 'bg-sky-500/10 text-sky-400' :
                                p.status === 'on-hold' ? 'bg-amber-500/10 text-amber-400' :
                                p.status === 'paused' ? 'bg-rose-500/10 text-rose-400' :
                                'bg-gray-700 text-gray-400'
                              }`}>
                                {p.status}
                              </span>
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
                            <td className="py-3.5 pr-6 font-medium text-gray-100">{s.name}</td>
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
                        if (member?.flexed || member?.ooo) await setStaffZone(pendingDrop.staffId, 'unassigned')
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
                              const { data } = await supabase.from('assignments').update({ project_id: p.id }).eq('id', id).select().single()
                              if (data) setAssignments(prev => prev.map(a => a.id === id ? data : a))
                            }
                          }
                          setDraggedAssignmentId(null)
                        }
                      }}
                    >
                      <div className="flex items-center justify-between mb-4">
                        <div className="flex items-center gap-2">
                          <h3 className="font-semibold text-gray-100">{p.name}</h3>
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
                        <span className={`px-2.5 py-1 rounded-full text-xs font-medium ${
                          p.status === 'active' ? 'bg-emerald-500/10 text-emerald-400' :
                          p.status === 'starting-soon' ? 'bg-sky-500/10 text-sky-400' :
                          p.status === 'on-hold' ? 'bg-amber-500/10 text-amber-400' :
                          p.status === 'paused' ? 'bg-rose-500/10 text-rose-400' :
                          'bg-gray-700 text-gray-400'
                        }`}>
                          {p.status}
                        </span>
                      </div>
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
                                  {a.assignment_role && <span className={`text-xs font-medium px-2 py-0.5 rounded-full ${roleColor(a.assignment_role)}`}>{a.assignment_role}</span>}
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
              const unassigned = [...staff].filter(s => !assignedIds.has(s.id) && !s.flexed && !s.ooo).sort((a, b) => a.name.localeCompare(b.name))
              const flexed = [...staff].filter(s => s.flexed && !assignedIds.has(s.id)).sort((a, b) => a.name.localeCompare(b.name))
              const oooStaff = [...staff].filter(s => s.ooo && !assignedIds.has(s.id)).sort((a, b) => a.name.localeCompare(b.name))

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
                    onDrop={async e => {
                      e.preventDefault(); setDragOverOOO(false)
                      const type = e.dataTransfer.getData('type')
                      const id = e.dataTransfer.getData('id')
                      if (type === 'staff' && id) {
                        await setStaffZone(id, 'ooo')
                        setDraggedStaffId(null)
                      } else if (type === 'assignment' && id) {
                        const assignment = assignments.find(a => a.id === id)
                        if (assignment) {
                          await setStaffZone(assignment.staff_id, 'ooo')
                          await removeAssignment(id)
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

                </div>
              )
            })()}

            </div>{/* end two-column layout */}
          </div>
        )}

        {/* Dashboard Tab */}
        {tab === 'dashboard' && (() => {
          const assignedStaffIds = new Set(assignments.map(a => a.staff_id))
          const availableStaff = staff.filter(s => !assignedStaffIds.has(s.id) && !s.flexed && !s.ooo)
          const oooCount = staff.filter(s => s.ooo).length
          const flexedCount = staff.filter(s => s.flexed).length

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
              <div className="grid grid-cols-3 gap-4">
                {statCard('Available Staff', availableStaff.length, `${staff.length} total`)}
                {statCard('Currently OOO', oooCount)}
                {statCard('Flexed', flexedCount)}
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
      </div>
    </div>
  )
}
