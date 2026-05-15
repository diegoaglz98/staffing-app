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
}

type Assignment = {
  id: string
  project_id: string
  staff_id: string
  assignment_role: string | null
}

const VALID_POSITIONS = ['SPA', 'SPL I', 'SPL II', 'Manager, Delivery', 'Senior SPL', 'Head of Delivery', 'GenAI Consultant']

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
  const [newStaff, setNewStaff] = useState({ name: '', position: '' })
  const [newAssignment, setNewAssignment] = useState({ project_id: '', staff_id: '', assignment_role: '' })

  const [editingProjectId, setEditingProjectId] = useState<string | null>(null)
  const [editProject, setEditProject] = useState({ name: '', customer_codename: '', status: 'active', duration_weeks: '' })

  const [editingStaffId, setEditingStaffId] = useState<string | null>(null)
  const [editStaff, setEditStaff] = useState({ name: '', position: '' })
  const [csvErrors, setCsvErrors] = useState<string[]>([])
  const [staffSort, setStaffSort] = useState<'default' | 'az' | 'za'>('az')
  const [hideAssigned, setHideAssigned] = useState(false)
  const [hideAssignedInDropdown, setHideAssignedInDropdown] = useState(false)
  const [projectSort, setProjectSort] = useState<'default' | 'az' | 'za'>('az')

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
    }]).select().single()
    if (data) {
      setStaff([data, ...staff])
      setNewStaff({ name: '', position: '' })
    }
  }

  function startEditStaff(s: Staff) {
    setEditStaff({ name: s.name, position: s.position ?? '' })
    setEditingStaffId(s.id)
  }

  async function saveStaff(id: string) {
    if (!editStaff.name.trim()) return
    const updates = { name: editStaff.name.trim(), position: editStaff.position || null }
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
    const rows = [...projects].sort((a, b) => a.name.localeCompare(b.name)).map(p => {
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
    const noBorder = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' }
    const cellBorder = { style: BorderStyle.SINGLE, size: 4, color: 'DDDDDD' }
    const border = { top: cellBorder, bottom: cellBorder, left: cellBorder, right: cellBorder }

    const children: (Paragraph | Table)[] = [
      new Paragraph({ text: 'Code Pod Staffing — Assignments', heading: HeadingLevel.TITLE }),
      new Paragraph({ text: `Exported ${new Date().toLocaleDateString()}`, spacing: { after: 400 } }),
    ]

    for (const p of [...projects].sort((a, b) => a.name.localeCompare(b.name))) {
      const pa = assignments.filter(a => a.project_id === p.id)
      children.push(new Paragraph({ text: p.name + (p.customer_codename ? ` (${p.customer_codename})` : ''), heading: HeadingLevel.HEADING_2, spacing: { before: 400 } }))
      children.push(new Paragraph({ children: [new TextRun({ text: `Status: ${p.status}${p.end_date ? `   |   Ends: ${p.end_date}` : ''}`, color: '666666', size: 18 })], spacing: { after: 160 } }))

      if (pa.length === 0) {
        children.push(new Paragraph({ text: 'No staff assigned.', spacing: { after: 200 } }))
        continue
      }

      const headerRow = new TableRow({
        children: ['Name', 'Position', 'Role'].map(h =>
          new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: h, bold: true })] })], borders: border, shading: { fill: 'F0F0F0' } })
        ),
        tableHeader: true,
      })

      const dataRows = pa.map(a => {
        const member = staff.find(s => s.id === a.staff_id)
        return new TableRow({
          children: [member?.name ?? 'Unknown', member?.position ?? '—', a.assignment_role ?? '—'].map(val =>
            new TableCell({ children: [new Paragraph(val)], borders: border })
          ),
        })
      })

      children.push(new Table({ rows: [headerRow, ...dataRows], width: { size: 100, type: WidthType.PERCENTAGE } }))
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
              <h2 className="text-xs font-semibold text-gray-400 uppercase tracking-wider mb-4">Add Project</h2>
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
                  <option value="on-hold">On Hold</option>
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
                    <th className="pb-3 font-medium text-xs uppercase tracking-wider">
                      <button onClick={() => setProjectSort(nextSort(projectSort))} className="flex items-center gap-1.5 hover:text-gray-300 transition-colors">
                        Name <span className="text-[10px] border border-gray-700 rounded px-1 py-0.5">{sortLabel(projectSort)}</span>
                      </button>
                    </th>
                    <th className="pb-3 font-medium text-xs uppercase tracking-wider">Customer</th>
                    <th className="pb-3 font-medium text-xs uppercase tracking-wider">Status</th>
                    <th className="pb-3 font-medium text-xs uppercase tracking-wider">Start</th>
                    <th className="pb-3 font-medium text-xs uppercase tracking-wider">End</th>
                    <th className="pb-3 font-medium text-xs uppercase tracking-wider">Duration (Wks)</th>
                    <th className="pb-3 font-medium text-xs uppercase tracking-wider">Staff</th>
                    <th></th>
                  </tr>
                </thead>
                <tbody>
                  {sortedList(projects, projectSort).map(p => {
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
                                <option value="on-hold">On Hold</option>
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
                                p.status === 'on-hold' ? 'bg-amber-500/10 text-amber-400' :
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
                    <th className="pb-3 font-medium text-xs uppercase tracking-wider">
                      <button onClick={() => setStaffSort(nextSort(staffSort))} className="flex items-center gap-1.5 hover:text-gray-300 transition-colors">
                        Name <span className="text-[10px] border border-gray-700 rounded px-1 py-0.5">{sortLabel(staffSort)}</span>
                      </button>
                    </th>
                    <th className="pb-3 font-medium text-xs uppercase tracking-wider">Position</th>
                    <th className="pb-3 font-medium text-xs uppercase tracking-wider">Assigned To</th>
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
                            <td className="py-2"></td>
                            <td className="py-2 text-right">
                              <div className="flex justify-end gap-3">
                                <button className={btnSave} onClick={() => saveStaff(s.id)}>Save</button>
                                <button className={btnCancel} onClick={() => setEditingStaffId(null)}>Cancel</button>
                              </div>
                            </td>
                          </>
                        ) : (
                          <>
                            <td className="py-3.5 font-medium text-gray-100">{s.name}</td>
                            <td className="py-3.5 text-gray-500">{s.position ?? '—'}</td>
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
                  {[...projects].sort((a, b) => a.name.localeCompare(b.name)).map(p => <option key={p.id} value={p.id}>{p.name}</option>)}
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

            {projects.length === 0 ? (
              <p className="text-gray-600 text-sm">Add projects first.</p>
            ) : (
              <div className="space-y-3">
                {[...projects].sort((a, b) => a.name.localeCompare(b.name)).map(p => {
                  const projectAssignments = assignments.filter(a => a.project_id === p.id)
                  return (
                    <div key={p.id} className="border border-gray-800 rounded-xl p-5 bg-gray-900/40">
                      <div className="flex items-center justify-between mb-4">
                        <h3 className="font-semibold text-gray-100">{p.name}</h3>
                        <span className={`px-2.5 py-1 rounded-full text-xs font-medium ${
                          p.status === 'active' ? 'bg-emerald-500/10 text-emerald-400' :
                          p.status === 'on-hold' ? 'bg-amber-500/10 text-amber-400' :
                          'bg-gray-700 text-gray-400'
                        }`}>
                          {p.status}
                        </span>
                      </div>
                      {projectAssignments.length === 0 ? (
                        <p className="text-gray-600 text-sm">No staff assigned.</p>
                      ) : (
                        <div className="space-y-2">
                          {projectAssignments.map(a => {
                            const member = staff.find(s => s.id === a.staff_id)
                            return (
                              <div key={a.id} className="flex items-center justify-between bg-gray-800/60 rounded-lg px-4 py-2.5 text-sm">
                                <div className="flex items-center gap-3">
                                  <span className="font-medium text-gray-200">{member?.name ?? 'Unknown'}</span>
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
          </div>
        )}

        {/* Dashboard Tab */}
        {tab === 'dashboard' && (() => {
          const assignedStaffIds = new Set(assignments.map(a => a.staff_id))
          const availableStaff = staff.filter(s => !assignedStaffIds.has(s.id))

          const stoCount = assignments.filter(a => a.assignment_role === 'STO').length
          const avgSTOs = projects.length ? (stoCount / projects.length).toFixed(1) : '0'

          const opsCount = assignments.filter(a => a.assignment_role === 'Ops Support').length
          const avgOps = projects.length ? (opsCount / projects.length).toFixed(1) : '0'

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
            .filter(p => p.status !== 'completed')
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
              <p className="text-3xl font-bold text-white">{value}</p>
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
              {/* Stat cards */}
              <div className="grid grid-cols-2 gap-4 sm:grid-cols-4">
                {statCard('Avg STOs / Project', avgSTOs)}
                {statCard('Avg Ops Support / Project', avgOps)}
                {statCard('Avg Projects / Supervisor', avgProjectsPerSupervisor)}
                {statCard('Available Staff', availableStaff.length, `${staff.length} total`)}
              </div>

              {/* Charts row */}
              <div className="grid grid-cols-2 gap-4">
                {barChart('Projects by Status', statusCounts, maxStatus, 'bg-[#193a29]')}
                {barChart('Projects by Customer', customerCounts, maxCustomer, 'bg-violet-500')}
              </div>

              {/* Headcount per project */}
              <div className="bg-gray-900 border border-gray-800 rounded-xl p-5">
                <p className="text-xs text-gray-500 uppercase tracking-wider mb-4">Total Headcount per Project</p>
                {headcounts.length === 0 ? (
                  <p className="text-gray-600 text-sm">No assignments yet.</p>
                ) : (
                  <div className="space-y-3">
                    {headcounts.map(({ name, count, roles }) => (
                      <div key={name}>
                        <div className="flex justify-between text-xs mb-1">
                          <span className="text-gray-300">{name}</span>
                          <div className="flex items-center gap-3">
                            {roles.Supervisor > 0 && <span className="text-violet-400">{roles.Supervisor} Supervisor{roles.Supervisor > 1 ? 's' : ''}</span>}
                            {roles.STO > 0 && <span className="text-blue-400">{roles.STO} STO{roles.STO > 1 ? 's' : ''}</span>}
                            {roles['Ops Support'] > 0 && <span className="text-emerald-400">{roles['Ops Support']} Ops Support</span>}
                            <span className="text-gray-500">{count} {count === 1 ? 'person' : 'people'}</span>
                          </div>
                        </div>
                        <div className="w-full bg-gray-800 rounded-full h-2">
                          <div className="h-2 rounded-full bg-emerald-600" style={{ width: `${(count / maxHeadcount) * 100}%` }} />
                        </div>
                      </div>
                    ))}
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
