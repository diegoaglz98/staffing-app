'use client'

import { useState, useEffect, useRef, Fragment } from 'react'
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
  emoji: string | null
  is_pilot: boolean
  is_internal: boolean
  weekly_target: number | null
  project_type: string | null
}

type ProjectType = {
  id: string
  name: string
}

// A "team" (Lead / Supervisor / STO / Ops) that can run one queue — Utilization tab
type UtilTeam = {
  id: string
  lead: string | null
  supervisor: string | null
  sto: string | null
  ops: string | null
  queue_complexity: string | null
  assignment_name: string | null
  assignment_complexity: string | null
  tier: string | null
  assign_priority: number | null
  notes: string | null
  reason: string | null
  sort_order: number
}

const UTIL_TIERS = ['S', 'A', 'B', 'C'] as const
const UTIL_POINTS: Record<string, number> = { S: 10, A: 7, B: 4, C: 1 }
const UTIL_TIER_INFO: Record<string, { label: string; detail: string; cls: string }> = {
  S: { label: 'S Tier', detail: 'High revenue result ($XM/week)', cls: 'bg-violet-500/15 text-violet-300' },
  A: { label: 'A Tier', detail: 'Mid revenue result ($X00K/week) OR very high potential pilot', cls: 'bg-emerald-500/15 text-emerald-300' },
  B: { label: 'B Tier', detail: 'Low revenue result ($X0K/week) OR low–mid potential pilot', cls: 'bg-sky-500/15 text-sky-300' },
  C: { label: 'C Tier', detail: 'Very low revenue result (<$10K/week) OR not assigned', cls: 'bg-gray-600/40 text-gray-300' },
}
const COMPLEXITY_RANK: Record<string, number> = { Low: 1, Mid: 2, High: 3 }

// Baseline "Code Pod - Staffing x Queues Utilization" sheet (static). Seeds the table when it's empty
// and backs "Reset to original sheet".
// [lead, supervisor, sto, ops, queue complexity, current assignment, assignment complexity, tier, assign prio, notes]
const UTIL_SEED_ROWS: [string | null, string | null, string | null, string | null, string | null, string, string | null, string | null, number | null, string | null][] = [
  ["Lisiman Hua", "Abril Herrlein", "Daniel Cueva", null, "Low", "Free", null, null, 2, null],
  ["Lisiman Hua", null, "Abril Herrlein", "Prashant Kumar", "High", "RSI OTS", "High", "A", null, "Not getting good commercial traction"],
  ["Lisiman Hua", null, "Gonzalo Lopez", null, "Mid", "Free", null, null, 2, null],
  ["Lisiman Hua", null, "Ale Pérez", "Sophia Basturk", "Mid", "Quadrillion", "Mid", "S", null, "Big commitment with the customer, cannot back down"],
  ["Lisiman Hua", "Abril Herrlein", "Andrea Sosa", "Claudia Avendaño", "Mid", "PKJA Finance", "Mid", "B", null, "Seems low priority for PKJA"],
  ["Lisiman Hua", null, "Aldo Vergara", null, "High", "CUA", "High", "S", null, null],
  ["Lisiman Hua", null, "Katara Zhiduo Xie", null, "Mid", "PKJA Harness", "Mid", "B", null, "Low revenue, customer moving slow"],
  ["Lisiman Hua", null, "Mahek Hemani", null, "Low", "Free", null, null, 2, null],
  ["Cole McCollum", null, "Aditi Chatterjee", null, "Low", "FNA1 Document QA Code", "Mid", "B", 3, null],
  ["Cole McCollum", null, "Gautam Anand", null, "High", "PKJA Contamination", "High", "A", null, null],
  ["Cole McCollum", "Gautam Anand", "Christian Rojas", null, "Low", "Free", null, null, 2, null],
  ["Cole McCollum", null, "Pavit Singh", "Jahnavi Murthy Padukone", "High", "Free", null, null, 3, null],
  ["David Espinosa", null, "Killian Mannarelli", null, "Mid", "Free", null, null, 1, null],
  ["David Espinosa", null, "David Espinosa", "Francisco Castellano", "High", "OpenClaw MM", "High", "S", null, null],
  ["Ivan Guevara", null, "Fernando Carabedo", null, "Mid", "Free", null, null, 2, null],
  ["Ivan Guevara", null, "Ezequiel Manoukian", null, "Mid", "SWE Atlas", "High", "B", null, null],
  ["Ivan Guevara", null, "Ignacio Ramet", null, "Mid", "SWE Atlas", "High", "B", null, null],
  ["Ivan Guevara", null, "Sebastian Toranzo", null, "Mid", "Tau 3", "High", "A", null, null],
  ["Ivan Guevara", null, "Nicolas Bizzotto", null, "Low", "Free", null, null, 1, null],
  ["Ivan Guevara", null, "Julio Reyna", null, "Low", "Enterprise Atlas Advanced", "Mid", "B", null, null],
  ["Jose Gutierrez", null, "Servando Rodríguez", null, "Mid", "Free", null, null, 1, null],
  ["Jose Gutierrez", "Servando Rodríguez", "Jorge Jimenez", null, "Mid", "PKJA Web Dev RL", "High", "A", null, null],
  ["Jose Gutierrez", null, "Achal Gandhi", null, "High", "OTS CyberPentest RL", "High", "B", null, null],
  ["Jose Gutierrez", "Pablo Caneva", "Matin Mansouri", null, "High", "OTS Co-found bench", "High", "C", null, null],
  ["Jose Gutierrez", null, "Pablo Caneva", "Valeria Guzman", "High", "OTS Multimodal Artifact Coding", "High", "B", null, null],
  ["Jose Gutierrez", null, "Ivan Seira", null, "High", "PKJA Ultra Evals i18n", "High", "A", null, null],
  ["Rafael Cruz", "Mila Avagimova", "Pablo Kogan", null, "Mid", "PKJA Ultra GAB Agent", "Mid", "A", null, null],
  ["Rafael Cruz", null, "Evan Honggo", "Daniela Ortiz", "High", "Enterprise Atlas Commercial", "High", "S", null, null],
  ["Rafael Cruz", "Mila Avagimova", "Oluwasegun Madarikan", null, "Mid", "Free", null, null, 1, null],
  ["Rafael Cruz", "Mila Avagimova", "Santos Iparraguirre", null, "High", "FNA1 Project Aura", "High", "S", null, null],
  ["Rafael Cruz", "Mila Avagimova", "Muskan Rastogi", null, "Mid", "PKJA Ultra Evals SxS", "Mid", "S", null, null],
  ["Rafael Cruz", "Mila Avagimova", "Pablo Paniagua", "Marius Delahay", "High", "Universe Gen", "High", "S", null, null],
]
function utilSeed(): Omit<UtilTeam, 'id'>[] {
  return UTIL_SEED_ROWS.map(([lead, supervisor, sto, ops, queue_complexity, assignment_name, assignment_complexity, tier, assign_priority, notes], i) => ({
    lead, supervisor, sto, ops, queue_complexity, assignment_name, assignment_complexity, tier, assign_priority, notes, reason: null, sort_order: i + 1,
  }))
}

const PROJECT_EMOJIS = ['📁', '🚀', '🎯', '🔥', '⭐', '🧪', '🤖', '🛡️', '📊', '💡', '⚙️', '🧠', '🌐', '📈', '🏆', '🐛', '🔒', '📝', '🎨', '⚡', '🧩', '📦', '🔧', '🩺', '💬', '🎓', '🗂️', '✅', '🔬', '🛰️', '📡', '💻', '📱', '☁️', '🔑', '🧵', '📐', '🕹️', '🎬', '🎧', '📷', '🏗️', '🚦', '🧭', '⏱️', '📅', '💰', '🏦', '⚖️', '🩹', '🧬', '🔭', '🌟', '💎', '🎲', '🃏', '🐙', '🦾', '👾', '🦉']

// "Confirm assignments with leads": projects whose leads are fixed by name (key = project name, lowercase)
const LEAD_CONFIRMATION_OVERRIDES: Record<string, string[]> = {
  'cua v3': ['Rafael Cruz', 'Ivan Guevara'],
}

// The person who owns this tool — always cc'd on update requests
const REQUESTER_SLACK_MENTION = '<@U020EEUSPPT>'

type Staff = {
  id: string
  name: string
  position: string | null
  ooo: boolean
  ooo_return_date: string | null
  flexed: boolean
  onboarding: boolean
  flex_notes: string | null
  emoji: string | null
  slack_user_id: string | null
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

type OrgNode = {
  staff_id: string
  parent_staff_id: string | null
}

const PRIORITIES = ['P00', 'P0', 'P1', 'P2']
const priorityRank = (p: string) => { const i = PRIORITIES.indexOf(p); return i === -1 ? 99 : i }
const priorityColor = (p: string) =>
  p === 'P00' ? 'bg-red-600/25 text-red-300' :
  p === 'P0' ? 'bg-red-500/10 text-red-400' :
  p === 'P1' ? 'bg-amber-500/10 text-amber-400' :
  'bg-sky-500/10 text-sky-400'
const priorityLabel = (p: string) => p === 'P00' ? '⚠️ P00' : p

const VALID_POSITIONS = ['SPA', 'SPL I', 'SPL II', 'Manager, Delivery', 'Senior SPL', 'Head of Delivery', 'Director, Delivery', 'GenAI Consultant']

const STATUS_OPTIONS: { value: string; label: string }[] = [
  { value: 'active', label: 'Active' },
  { value: 'starting-soon', label: 'Starting Soon' },
  { value: 'paused', label: 'Paused' },
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
  const [tab, setTab] = useState<'projects' | 'staff' | 'assignments' | 'dashboard' | 'scenarios' | 'milestones' | 'orgchart' | 'consolidated' | 'utilization'>('dashboard')
  const [projects, setProjects] = useState<Project[]>([])
  const [staff, setStaff] = useState<Staff[]>([])
  const [assignments, setAssignments] = useState<Assignment[]>([])
  const [scenarios, setScenarios] = useState<Scenario[]>([])
  const [scenarioAssignments, setScenarioAssignments] = useState<ScenarioAssignment[]>([])
  const [milestones, setMilestones] = useState<Milestone[]>([])
  const [projectTypes, setProjectTypes] = useState<ProjectType[]>([])
  const [utilTeams, setUtilTeams] = useState<UtilTeam[]>([])
  const [utilFilter, setUtilFilter] = useState<'all' | 'utilized' | 'under' | 'not'>('all')
  const [utilLead, setUtilLead] = useState('')
  const [utilSearch, setUtilSearch] = useState('')
  const [utilEdit, setUtilEdit] = useState<(Omit<UtilTeam, 'id'> & { id: string | null }) | null>(null)
  const [utilLocalOnly, setUtilLocalOnly] = useState(false)
  const utilSeeding = useRef(false)
  const [manageTypesOpen, setManageTypesOpen] = useState(false)
  const [newTypeName, setNewTypeName] = useState('')
  const [orgChart, setOrgChart] = useState<OrgNode[]>([])
  const [orgUnassignedOnly, setOrgUnassignedOnly] = useState(true)
  const [dragOverOrgId, setDragOverOrgId] = useState<string | null>(null)
  const [collapsedOrg, setCollapsedOrg] = useState<Record<string, boolean>>({})
  const [orgShowConsultants, setOrgShowConsultants] = useState(true)
  const [orgColorsOn, setOrgColorsOn] = useState(true)
  const [orgZoom, setOrgZoom] = useState(1)
  const [orgShowProjects, setOrgShowProjects] = useState(false)
  const [orgLayout, setOrgLayout] = useState<'horizontal' | 'vertical'>('horizontal')
  const orgScrollRef = useRef<HTMLDivElement>(null)
  const orgScrollDir = useRef(0)
  const orgScrollTimer = useRef<number | null>(null)
  function stopOrgAutoScroll() {
    orgScrollDir.current = 0
    if (orgScrollTimer.current != null) { clearInterval(orgScrollTimer.current); orgScrollTimer.current = null }
  }
  function orgAutoScroll(e: React.DragEvent) {
    const el = orgScrollRef.current
    if (!el) return
    const rect = el.getBoundingClientRect()
    const threshold = 90
    let dir = 0
    if (e.clientX < rect.left + threshold) dir = -1
    else if (e.clientX > rect.right - threshold) dir = 1
    orgScrollDir.current = dir
    if (dir !== 0 && orgScrollTimer.current == null) {
      orgScrollTimer.current = window.setInterval(() => {
        const c = orgScrollRef.current
        if (c && orgScrollDir.current) c.scrollLeft += orgScrollDir.current * 20
      }, 16)
    } else if (dir === 0) {
      stopOrgAutoScroll()
    }
  }
  useEffect(() => {
    const onEnd = () => stopOrgAutoScroll()
    window.addEventListener('dragend', onEnd)
    window.addEventListener('drop', onEnd)
    return () => { window.removeEventListener('dragend', onEnd); window.removeEventListener('drop', onEnd) }
  }, [])
  const [pendingOrgRemove, setPendingOrgRemove] = useState<string | null>(null)
  const [milestoneDrafts, setMilestoneDrafts] = useState<Record<string, { title: string; priority: string; due_date: string }>>({})
  const [hideEmptyMilestoneProjects, setHideEmptyMilestoneProjects] = useState(false)
  const [showCompletedMilestones, setShowCompletedMilestones] = useState(false)
  const [msSummaryView, setMsSummaryView] = useState(false)
  // Consolidated Manager
  const [cmView, setCmView] = useState<'cards' | 'table'>('cards')
  const [cmFilter, setCmFilter] = useState<'all' | 'attention' | 'flagged' | 'understaffed' | 'nopp' | 'overdue' | 'p0' | 'noteam'>('all')
  const [cmSearch, setCmSearch] = useState('')
  const [cmCollapsed, setCmCollapsed] = useState<Record<string, boolean>>({})
  const [cmMenu, setCmMenu] = useState<null | 'export' | 'clear' | 'slack'>(null)
  const [cmShowAssign, setCmShowAssign] = useState(false)
  const [cmHighlight, setCmHighlight] = useState<string | null>(null)
  const [cmGroupCollapsed, setCmGroupCollapsed] = useState<Record<string, boolean>>({})
  const [weeklyDrafts, setWeeklyDrafts] = useState<Record<string, string>>({})
  const [slackReqStatus, setSlackReqStatus] = useState<Record<string, 'sending' | 'sent' | 'error'>>({})
  const [slackPermalinks, setSlackPermalinks] = useState<Record<string, string>>({})
  const [confirmEmptyUpdate, setConfirmEmptyUpdate] = useState<Project | null>(null)
  const [leadIdPrompt, setLeadIdPrompt] = useState<{ id: string; name: string }[] | null>(null)
  const [leadIdInputs, setLeadIdInputs] = useState<Record<string, string>>({})
  const [leadPreview, setLeadPreview] = useState<{ leads: { id: string; name: string; projects: string[] }[]; override: Record<string, string> } | null>(null)
  const [projectCsvPrompt, setProjectCsvPrompt] = useState(false)
  const [leadStatus, setLeadStatus] = useState<'idle' | 'sending' | 'sent' | 'error' | 'empty'>('idle')
  const [leadPermalink, setLeadPermalink] = useState<string | null>(null)
  const [pendingSlackIds, setPendingSlackIds] = useState<{ project: Project; people: { id: string; name: string; role: string }[] } | null>(null)
  const [slackIdInputs, setSlackIdInputs] = useState<Record<string, string>>({})
  const [overviewStatus, setOverviewStatus] = useState<'idle' | 'sending' | 'sent' | 'error' | 'empty'>('idle')
  const [overviewPermalink, setOverviewPermalink] = useState<string | null>(null)
  const [overviewOpen, setOverviewOpen] = useState(false)
  const [overviewIdInputs, setOverviewIdInputs] = useState<Record<string, string>>({})
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

  const [newProject, setNewProject] = useState({ name: '', customer_codename: '', status: 'active', duration_weeks: '', is_pilot: false, is_internal: false, project_type: '' })
  const [newStaff, setNewStaff] = useState({ name: '', position: '', ooo: false, ooo_return_date: '' })
  const [newAssignment, setNewAssignment] = useState({ project_id: '', staff_id: '', assignment_role: '' })

  const [editingProjectId, setEditingProjectId] = useState<string | null>(null)
  const [editProject, setEditProject] = useState({ name: '', customer_codename: '', status: 'active', duration_weeks: '', is_pilot: false, is_internal: false })
  const [renameProjectId, setRenameProjectId] = useState<string | null>(null)
  const [renameValue, setRenameValue] = useState('')

  const [editingStaffId, setEditingStaffId] = useState<string | null>(null)
  const [editStaff, setEditStaff] = useState({ name: '', position: '', ooo: false, ooo_return_date: '', flex_notes: '', slack_user_id: '' })
  const [inlineAssignment, setInlineAssignment] = useState({ project_id: '', assignment_role: '' })
  const [csvErrors, setCsvErrors] = useState<string[]>([])
  const [staffSort, setStaffSort] = useState<'default' | 'az' | 'za'>('az')
  const [positionFilter, setPositionFilter] = useState('')
  const [hideAssigned, setHideAssigned] = useState(false)
  const [hideAssignedInDropdown, setHideAssignedInDropdown] = useState(false)
  const [projectSort, setProjectSort] = useState<{ col: string; dir: 'az' | 'za' }>({ col: 'status', dir: 'az' })
  const [projectSearch, setProjectSearch] = useState('')
  const [assignmentSearch, setAssignmentSearch] = useState('')
  const [showSupervisorsInChart, setShowSupervisorsInChart] = useState(false)
  const [customerActiveOnly, setCustomerActiveOnly] = useState(false)
  const [excludePilots, setExcludePilots] = useState(false)
  const [excludeInternal, setExcludeInternal] = useState(false)
  const [visibleStatuses, setVisibleStatuses] = useState<Record<string, boolean>>({
    'active': true, 'starting-soon': true, 'paused': true, 'completed': false,
  })
  const [statusFilterOpen, setStatusFilterOpen] = useState(false)
  const [addingToProjectId, setAddingToProjectId] = useState<string | null>(null)
  const [showAddProjectInline, setShowAddProjectInline] = useState(false)
  const [supervisorFilter, setSupervisorFilter] = useState('')
  const [showProjectAddBetween, setShowProjectAddBetween] = useState(false)
  const [groupBySupervisor, setGroupBySupervisor] = useState(false)
  const [editingStaffAssignmentId, setEditingStaffAssignmentId] = useState<string | null>(null)
  const [confirmClearInactive, setConfirmClearInactive] = useState(false)
  const [clearMsPrompt, setClearMsPrompt] = useState(false)
  const [confirmClearAll, setConfirmClearAll] = useState(false)
  const [logoSpins, setLogoSpins] = useState(0)
  const [emojiPickerProjectId, setEmojiPickerProjectId] = useState<string | null>(null)
  const [customEmoji, setCustomEmoji] = useState('')
  const [emojiPickerStaffId, setEmojiPickerStaffId] = useState<string | null>(null)
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
  const [pendingFlex, setPendingFlex] = useState<{ staffId: string; assignmentId: string | null } | null>(null)
  const [flexNote, setFlexNote] = useState('')
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
        supabase.from('org_chart').select('*'),
        supabase.from('project_types').select('*').order('name', { ascending: true }),
        supabase.from('utilization_teams').select('*').order('sort_order', { ascending: true }),
      ])
      const [p, s, a, sc, sa, ms, oc, pt, ut] = await Promise.race([queries, timeout])
      // Core tables must load; scenario tables are optional (may not exist yet)
      if (p.error || s.error || a.error) throw (p.error || s.error || a.error)
      if (p.data) {
        // "On Hold" was retired in favour of "Paused" — migrate any lingering rows
        const onHold = p.data.filter((pr: Project) => pr.status === 'on-hold')
        if (onHold.length > 0) {
          await supabase.from('projects').update({ status: 'paused' }).in('id', onHold.map((pr: Project) => pr.id))
          p.data.forEach((pr: Project) => { if (pr.status === 'on-hold') pr.status = 'paused' })
        }
        setProjects(p.data)
      }
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
      if (oc.data) setOrgChart(oc.data)
      if (pt.data) setProjectTypes(pt.data)
      if (ut.error) {
        // Table not created yet → show the static sheet (edits stay in this browser session)
        setUtilLocalOnly(true)
        setUtilTeams(utilSeed().map((t, i) => ({ ...t, id: `local-${i}` })))
      } else if (ut.data && ut.data.length === 0 && !utilSeeding.current) {
        utilSeeding.current = true
        const { data: seeded } = await supabase.from('utilization_teams').insert(utilSeed()).select()
        setUtilLocalOnly(false)
        setUtilTeams([...(seeded ?? [])].sort((a, b) => a.sort_order - b.sort_order))
      } else if (ut.data && !(ut.data.length === 0 && utilSeeding.current)) {
        setUtilLocalOnly(false)
        setUtilTeams(ut.data)
      }
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
      is_pilot: newProject.is_pilot,
      is_internal: newProject.is_internal,
      ...(newProject.project_type ? { project_type: newProject.project_type } : {}),
    }]).select().single()
    if (data) {
      setProjects([data, ...projects])
      setNewProject({ name: '', customer_codename: '', status: 'active', duration_weeks: '', is_pilot: false, is_internal: false, project_type: '' })
    }
  }

  function startEditProject(p: Project) {
    const weeksRemaining = p.end_date
      ? String(Math.round((new Date(p.end_date).getTime() - new Date(p.start_date ?? '').getTime()) / (7 * 24 * 60 * 60 * 1000)))
      : ''
    setEditProject({ name: p.name, customer_codename: p.customer_codename ?? '', status: p.status, duration_weeks: weeksRemaining, is_pilot: p.is_pilot, is_internal: p.is_internal })
    setEditingProjectId(p.id)
  }

  async function saveProjectName(id: string) {
    const name = renameValue.trim()
    setRenameProjectId(null)
    const current = projects.find(p => p.id === id)
    if (!name || !current || current.name === name) return
    const { data } = await supabase.from('projects').update({ name }).eq('id', id).select().single()
    if (data) setProjects(projects.map(p => p.id === id ? data : p))
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
      is_pilot: editProject.is_pilot,
      is_internal: editProject.is_internal,
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

  const INACTIVE_STATUSES = ['paused', 'completed']
  function inactiveAssignmentIds() {
    const inactiveProjectIds = new Set(projects.filter(p => INACTIVE_STATUSES.includes(p.status)).map(p => p.id))
    return assignments.filter(a => inactiveProjectIds.has(a.project_id)).map(a => a.id)
  }
  async function clearAllAssignments() {
    setConfirmClearAll(false)
    if (assignments.length === 0) return
    const { error } = await supabase.from('assignments').delete().neq('id', '00000000-0000-0000-0000-000000000000')
    if (error) { console.error('clearAllAssignments failed', error); return }
    setAssignments([])
  }

  async function clearInactiveStaffing() {
    setConfirmClearInactive(false)
    const inactiveProjectIds = projects.filter(p => INACTIVE_STATUSES.includes(p.status)).map(p => p.id)
    if (inactiveProjectIds.length === 0) return
    const { error } = await supabase.from('assignments').delete().in('project_id', inactiveProjectIds)
    if (error) { console.error('clearInactiveStaffing failed', error); return }
    setAssignments(prev => prev.filter(a => !inactiveProjectIds.includes(a.project_id)))
  }

  async function toggleProjectFlag(id: string, flagged: boolean) {
    const { data } = await supabase.from('projects').update({ flagged }).eq('id', id).select().single()
    if (data) setProjects(projects.map(p => p.id === id ? data : p))
  }

  async function updateProjectPilot(id: string, is_pilot: boolean) {
    const { data } = await supabase.from('projects').update({ is_pilot }).eq('id', id).select().single()
    if (data) setProjects(projects.map(p => p.id === id ? data : p))
  }

  async function updateProjectInternal(id: string, is_internal: boolean) {
    const { data } = await supabase.from('projects').update({ is_internal }).eq('id', id).select().single()
    if (data) setProjects(projects.map(p => p.id === id ? data : p))
  }

  // --- Project types (managed dropdown list, like Sheets data validation) ---
  // Options = saved list + any value already on a project (so nothing goes missing)
  function projectTypeOptions() {
    const names = new Set(projectTypes.map(t => t.name))
    projects.forEach(p => { if (p.project_type) names.add(p.project_type) })
    return [...names].sort((a, b) => a.localeCompare(b))
  }
  async function updateProjectType(id: string, project_type: string | null) {
    const { data } = await supabase.from('projects').update({ project_type }).eq('id', id).select().single()
    if (data) setProjects(prev => prev.map(p => p.id === id ? data : p))
  }
  // Creates the type if new (case-insensitive match reuses the existing one); returns the canonical name
  async function addProjectType(raw: string): Promise<string | null> {
    const name = raw.trim()
    if (!name) return null
    const existing = projectTypeOptions().find(n => n.toLowerCase() === name.toLowerCase())
    if (existing) {
      if (!projectTypes.some(t => t.name === existing)) {
        const { data } = await supabase.from('project_types').insert([{ name: existing }]).select().single()
        if (data) setProjectTypes(prev => [...prev, data])
      }
      return existing
    }
    const { data, error } = await supabase.from('project_types').insert([{ name }]).select().single()
    if (error) { alert(`Couldn't save project type: ${error.message}`); return null }
    if (data) setProjectTypes(prev => [...prev, data])
    return name
  }
  async function renameProjectType(oldName: string, raw: string) {
    const name = raw.trim()
    if (!name || name === oldName) return
    const t = projectTypes.find(x => x.name === oldName)
    if (t) {
      const { data } = await supabase.from('project_types').update({ name }).eq('id', t.id).select().single()
      if (data) setProjectTypes(prev => prev.map(x => x.id === t.id ? data : x))
    } else {
      const { data } = await supabase.from('project_types').insert([{ name }]).select().single()
      if (data) setProjectTypes(prev => [...prev, data])
    }
    await supabase.from('projects').update({ project_type: name }).eq('project_type', oldName)
    setProjects(prev => prev.map(p => p.project_type === oldName ? { ...p, project_type: name } : p))
  }
  async function deleteProjectType(name: string) {
    const inUse = projects.filter(p => p.project_type === name).length
    if (inUse > 0 && !confirm(`"${name}" is set on ${inUse} project${inUse === 1 ? '' : 's'}. Delete it and clear it from ${inUse === 1 ? 'that project' : 'those projects'}?`)) return
    await supabase.from('project_types').delete().eq('name', name)
    setProjectTypes(prev => prev.filter(t => t.name !== name))
    if (inUse > 0) {
      await supabase.from('projects').update({ project_type: null }).eq('project_type', name)
      setProjects(prev => prev.map(p => p.project_type === name ? { ...p, project_type: null } : p))
    }
  }
  // Handles the dropdown's special "add new" / "manage" entries
  async function handleProjectTypeSelect(value: string, apply: (v: string | null) => void) {
    if (value === '__add__') {
      const name = await addProjectType(prompt('New project type:') ?? '')
      if (name) apply(name)
    } else if (value === '__manage__') {
      setManageTypesOpen(true)
    } else {
      apply(value || null)
    }
  }

  // --- Utilization (teams × queues) ---
  function utilIsFree(t: UtilTeam) {
    const n = (t.assignment_name ?? '').trim().toLowerCase()
    return !n || n === 'free'
  }
  // A free team counts as C ("not assigned"); an assigned team with no score yet is unscored (null)
  function utilTier(t: UtilTeam): string | null {
    return utilIsFree(t) ? 'C' : (t.tier && UTIL_POINTS[t.tier] ? t.tier : null)
  }
  // Utilized / under-utilized / not utilized, plus the "why" (a typed reason overrides the automatic one)
  function utilStatus(t: UtilTeam): { status: 'utilized' | 'under' | 'not'; reason: string } {
    const custom = (t.reason ?? '').trim()
    if (utilIsFree(t)) {
      return { status: 'not', reason: custom || `No current assignment${t.assign_priority ? ` · priority ${t.assign_priority} to staff` : ''}` }
    }
    if (t.tier === 'C') return { status: 'under', reason: custom || 'Very low revenue result (C tier)' }
    const qc = COMPLEXITY_RANK[t.queue_complexity ?? ''] ?? 0
    const ac = COMPLEXITY_RANK[t.assignment_complexity ?? ''] ?? 0
    if (qc && ac && qc > ac) {
      return { status: 'under', reason: custom || `${t.queue_complexity}-complexity team on a ${t.assignment_complexity}-complexity queue` }
    }
    return { status: 'utilized', reason: custom || (t.tier ? '' : 'Score not set yet') }
  }
  // 1–10 rating: average tier points (S=10, A=7, B=4, C=1) across scored teams
  function utilRating(list: UtilTeam[]) {
    const pts = list.map(utilTier).filter((x): x is string => !!x).map(x => UTIL_POINTS[x])
    return pts.length ? pts.reduce((a, b) => a + b, 0) / pts.length : null
  }
  async function updateUtilTeam(id: string, patch: Partial<UtilTeam>) {
    setUtilTeams(prev => prev.map(t => t.id === id ? { ...t, ...patch } : t))
    if (utilLocalOnly) return
    const { error } = await supabase.from('utilization_teams').update(patch).eq('id', id)
    if (error) alert(`Couldn't save: ${error.message}`)
  }
  // Edit popup: open for an existing team, or a blank one (prefilled with the lead filter)
  function openUtilEdit(t?: UtilTeam) {
    if (t) { setUtilEdit({ ...t }); return }
    setUtilEdit({
      id: null, lead: utilLead && utilLead !== '__none__' ? utilLead : null, supervisor: null, sto: null, ops: null,
      queue_complexity: null, assignment_name: 'Free', assignment_complexity: null, tier: null, assign_priority: null,
      notes: null, reason: null, sort_order: utilTeams.reduce((m, x) => Math.max(m, x.sort_order), 0) + 1,
    })
  }
  async function saveUtilEdit() {
    if (!utilEdit) return
    const clean = (v: string | null) => (v ?? '').trim() || null
    const { id, ...rest } = utilEdit
    const free = !clean(rest.assignment_name) || clean(rest.assignment_name)!.toLowerCase() === 'free'
    const row = {
      ...rest,
      lead: clean(rest.lead), supervisor: clean(rest.supervisor), sto: clean(rest.sto), ops: clean(rest.ops),
      assignment_name: free ? 'Free' : clean(rest.assignment_name),
      assignment_complexity: free ? null : rest.assignment_complexity,
      tier: free ? null : rest.tier,
      assign_priority: free ? rest.assign_priority : null,
      notes: clean(rest.notes), reason: clean(rest.reason),
    }
    if (id) {
      await updateUtilTeam(id, row)
    } else if (utilLocalOnly) {
      setUtilTeams(prev => [...prev, { ...row, id: `local-${Date.now()}` }])
    } else {
      const { data, error } = await supabase.from('utilization_teams').insert([row]).select().single()
      if (error) { alert(`Couldn't add team: ${error.message}`); return }
      if (data) setUtilTeams(prev => [...prev, data])
    }
    setUtilEdit(null)
  }
  async function deleteUtilTeam(t: { id: string | null; sto: string | null; lead: string | null }) {
    if (!t.id) return
    const who = [t.sto, t.lead].filter(Boolean).join(' / ') || 'this team'
    if (!confirm(`Remove ${who} from the utilization sheet?`)) return
    if (!utilLocalOnly) await supabase.from('utilization_teams').delete().eq('id', t.id)
    setUtilTeams(prev => prev.filter(x => x.id !== t.id))
    setUtilEdit(null)
  }
  async function resetUtilToSheet() {
    if (!confirm('Reset to the original sheet? This replaces all teams (and any edits) with the original 32 teams.')) return
    if (utilLocalOnly) { setUtilTeams(utilSeed().map((t, i) => ({ ...t, id: `local-${i}` }))); return }
    const { error } = await supabase.from('utilization_teams').delete().neq('id', '00000000-0000-0000-0000-000000000000')
    if (error) { alert(`Couldn't reset: ${error.message}`); return }
    const { data } = await supabase.from('utilization_teams').insert(utilSeed()).select()
    setUtilTeams([...(data ?? [])].sort((a, b) => a.sort_order - b.sort_order))
  }
  function exportUtilCSV() {
    const rows: (string | number | boolean | null)[][] = [[
      'Lead', 'Supervisor', 'STO', 'Ops', 'Queue Complexity', 'Current Assignment Name', 'Current Assignment Complexity',
      'Current Assignment Tier', 'Prio (1-3) on giving an assignment', 'Notes', 'Utilization Status', 'Why',
    ]]
    utilTeams.forEach(t => {
      const st = utilStatus(t)
      rows.push([
        t.lead, t.supervisor, t.sto, t.ops, t.queue_complexity, t.assignment_name, t.assignment_complexity,
        utilIsFree(t) ? '' : t.tier, t.assign_priority, t.notes,
        st.status === 'utilized' ? 'Utilized' : st.status === 'under' ? 'Under-utilized' : 'Not utilized', st.reason,
      ])
    })
    downloadCSV(`utilization-${new Date().toISOString().split('T')[0]}.csv`, rows)
  }

  async function updateProjectEmoji(id: string, emoji: string | null) {
    const { data } = await supabase.from('projects').update({ emoji }).eq('id', id).select().single()
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
    setEditStaff({ name: s.name, position: s.position ?? '', ooo: s.ooo, ooo_return_date: s.ooo_return_date ?? '', flex_notes: s.flex_notes ?? '', slack_user_id: s.slack_user_id ?? '' })
    setEditingStaffId(s.id)
  }

  async function saveStaff(id: string) {
    if (!editStaff.name.trim()) return
    const updates = { name: editStaff.name.trim(), position: editStaff.position || null, ooo: editStaff.ooo, ooo_return_date: editStaff.ooo_return_date || null, flex_notes: editStaff.flex_notes.trim() || null, slack_user_id: editStaff.slack_user_id.trim() || null }
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

  function csvCell(v: unknown): string {
    if (v === null || v === undefined) return ''
    const s = String(v)
    return /[",\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s
  }
  function downloadCSV(filename: string, rows: (string | number | boolean | null)[][]) {
    const csv = rows.map(r => r.map(csvCell).join(',')).join('\n')
    const blob = new Blob(['﻿' + csv], { type: 'text/csv;charset=utf-8;' })
    const url = URL.createObjectURL(blob)
    const a = document.createElement('a')
    a.href = url
    a.download = filename
    a.click()
    URL.revokeObjectURL(url)
  }
  function exportProjectsCSV(includeCompleted: boolean) {
    const header = [
      'id', 'name', 'customer_codename', 'status', 'start_date', 'end_date',
      'flagged', 'emoji', 'is_pilot', 'is_internal', 'project_type', 'weekly_target', 'staff_assigned',
    ]
    const rows: (string | number | boolean | null)[][] = [header]
    ;[...projects]
      .filter(p => includeCompleted || p.status !== 'completed')
      .sort((a, b) => a.name.localeCompare(b.name))
      .forEach(p => {
        const staffAssigned = assignments
          .filter(a => a.project_id === p.id)
          .map(a => {
            const s = staff.find(x => x.id === a.staff_id)
            if (!s) return null
            return a.assignment_role ? `${s.name} (${a.assignment_role})` : s.name
          })
          .filter(Boolean)
          .sort((a, b) => a!.localeCompare(b!))
          .join('; ')
        rows.push([
          p.id, p.name, p.customer_codename, p.status, p.start_date, p.end_date,
          p.flagged, p.emoji, p.is_pilot, p.is_internal, p.project_type, p.weekly_target, staffAssigned,
        ])
      })
    downloadCSV(`projects-${new Date().toISOString().split('T')[0]}.csv`, rows)
    setProjectCsvPrompt(false)
  }
  function exportStaffCSV() {
    const header = [
      'id', 'name', 'position', 'ooo', 'ooo_return_date', 'flexed',
      'onboarding', 'flex_notes', 'emoji', 'slack_user_id',
    ]
    const rows: (string | number | boolean | null)[][] = [header]
    ;[...staff]
      .sort((a, b) => a.name.localeCompare(b.name))
      .forEach(s => {
        rows.push([
          s.id, s.name, s.position, s.ooo, s.ooo_return_date, s.flexed,
          s.onboarding, s.flex_notes, s.emoji, s.slack_user_id,
        ])
      })
    downloadCSV(`staff-${new Date().toISOString().split('T')[0]}.csv`, rows)
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

    const prColors: Record<string, [string, string]> = {
      P00: ['#f8d0cc', '#a01808'],
      P0: ['#fdecea', '#c0392b'],
      P1: ['#fef3c7', '#b45309'],
      P2: ['#e0f2fe', '#0369a1'],
    }
    const prPill = (pr: string) => {
      const [bg, fg] = prColors[pr] ?? ['#eeeeee', '#555555']
      const label = pr === 'P00' ? '⚠ P00' : pr
      return `<span style="background:${bg};color:${fg};padding:2px 8px;border-radius:999px;font-weight:600;font-size:12px">${label}</span>`
    }

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
  <td>${prPill(m.priority)}</td>
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

  async function updateAssignmentProject(id: string, projectId: string) {
    const a = assignments.find(x => x.id === id)
    if (!a || a.project_id === projectId) return
    const dup = assignments.some(x => x.id !== id && x.project_id === projectId && x.staff_id === a.staff_id)
    if (dup) return
    const { data } = await supabase.from('assignments').update({ project_id: projectId }).eq('id', id).select().single()
    if (data) setAssignments(prev => prev.map(x => x.id === id ? data : x))
  }

  async function updateStaffEmoji(id: string, emoji: string | null) {
    const { data } = await supabase.from('staff').update({ emoji }).eq('id', id).select().single()
    if (data) setStaff(prev => prev.map(s => s.id === id ? data : s))
  }

  async function clearStaffOOO(staffId: string) {
    const { data } = await supabase.from('staff').update({ ooo: false, ooo_return_date: null }).eq('id', staffId).select().single()
    if (data) setStaff(prev => prev.map(s => s.id === staffId ? data : s))
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

  async function updateWeeklyTarget(id: string, raw: string) {
    const n = raw.trim() === '' ? null : Math.max(0, Math.floor(Number(raw)) || 0)
    const { data } = await supabase.from('projects').update({ weekly_target: n }).eq('id', id).select().single()
    if (data) setProjects(prev => prev.map(p => p.id === id ? data : p))
    setWeeklyDrafts(prev => { const c = { ...prev }; delete c[id]; return c })
  }
  // Clear all milestones and/or production-plan targets from active projects
  async function clearMilestonesAndTargets(opts: { milestones: boolean; targets: boolean }) {
    const activeIds = projects.filter(p => p.status === 'active').map(p => p.id)
    setClearMsPrompt(false)
    if (activeIds.length === 0) return
    if (opts.milestones) {
      await supabase.from('milestones').delete().in('project_id', activeIds)
      setMilestones(prev => prev.filter(m => !activeIds.includes(m.project_id)))
    }
    if (opts.targets) {
      await supabase.from('projects').update({ weekly_target: null }).in('id', activeIds)
      setProjects(prev => prev.map(p => activeIds.includes(p.id) ? { ...p, weekly_target: null } : p))
      setWeeklyDrafts({})
    }
  }
  async function addToOrg(staffId: string, parentStaffId: string | null) {
    if (orgChart.some(n => n.staff_id === staffId)) return
    const { data } = await supabase.from('org_chart').insert([{ staff_id: staffId, parent_staff_id: parentStaffId }]).select().single()
    if (data) setOrgChart(prev => [...prev, data])
  }
  async function moveOrgNode(staffId: string, parentStaffId: string | null) {
    const { data } = await supabase.from('org_chart').update({ parent_staff_id: parentStaffId }).eq('staff_id', staffId).select().single()
    if (data) setOrgChart(prev => prev.map(n => n.staff_id === staffId ? data : n))
  }
  async function removeFromOrg(staffId: string) {
    // Collect the node and all its descendants
    const toRemove = new Set([staffId])
    let grew = true
    while (grew) {
      grew = false
      for (const n of orgChart) {
        if (n.parent_staff_id && toRemove.has(n.parent_staff_id) && !toRemove.has(n.staff_id)) { toRemove.add(n.staff_id); grew = true }
      }
    }
    await supabase.from('org_chart').delete().in('staff_id', Array.from(toRemove))
    setOrgChart(prev => prev.filter(n => !toRemove.has(n.staff_id)))
  }
  async function removeOrgReattach(staffId: string) {
    const newParent = orgChart.find(n => n.staff_id === staffId)?.parent_staff_id ?? null
    const kids = orgChart.filter(n => n.parent_staff_id === staffId)
    for (const k of kids) {
      await supabase.from('org_chart').update({ parent_staff_id: newParent }).eq('staff_id', k.staff_id)
    }
    await supabase.from('org_chart').delete().eq('staff_id', staffId)
    setOrgChart(prev => prev
      .filter(n => n.staff_id !== staffId)
      .map(n => n.parent_staff_id === staffId ? { ...n, parent_staff_id: newParent } : n))
  }

  async function exportOrgChart(mode: 'download' | 'print') {
    const el = document.getElementById('org-chart-capture')
    if (!el) return
    const { toPng } = await import('html-to-image')
    const bg = document.documentElement.classList.contains('light') ? '#ffffff' : '#030712'
    // Scale wide charts down so the exported image isn't enormous
    const maxWidth = 2400
    const k = Math.min(1, maxWidth / (el.scrollWidth || maxWidth))
    const dataUrl = await toPng(el, {
      backgroundColor: bg,
      pixelRatio: 1.5,
      cacheBust: true,
      width: Math.ceil(el.scrollWidth * k),
      height: Math.ceil(el.scrollHeight * k),
      style: { transform: `scale(${k})`, transformOrigin: 'top left' },
    })
    if (mode === 'download') {
      const a = document.createElement('a')
      a.href = dataUrl
      a.download = `org-chart-${new Date().toISOString().split('T')[0]}.png`
      a.click()
    } else {
      const w = window.open('')
      if (w) {
        w.document.write(`<img src="${dataUrl}" style="max-width:100%" onload="window.print()" />`)
        w.document.close()
      }
    }
  }

  function stepWeeklyTarget(p: Project, delta: number) {
    const cur = weeklyDrafts[p.id] != null ? (Number(weeklyDrafts[p.id]) || 0) : (p.weekly_target ?? 0)
    updateWeeklyTarget(p.id, String(Math.max(0, cur + delta)))
  }

  async function requestMilestoneUpdate(p: Project, slackIdOverride?: Record<string, string>): Promise<boolean> {
    setSlackReqStatus(prev => ({ ...prev, [p.id]: 'sending' }))
    const openMs = [...milestones.filter(m => m.project_id === p.id && !m.done)]
      .sort((a, b) => priorityRank(a.priority) - priorityRank(b.priority) || a.created_at.localeCompare(b.created_at))
      .map(m => `${m.priority} — ${m.title}${m.due_date ? ` (due ${m.due_date})` : ''}`)
    const mentionFor = (role: string) => assignments
      .filter(a => a.project_id === p.id && a.assignment_role === role)
      .map(a => staff.find(s => s.id === a.staff_id))
      .filter((s): s is Staff => !!s)
      .map(s => {
        const sid = slackIdOverride?.[s.id] ?? s.slack_user_id
        return sid ? `<@${sid}>` : s.name
      })
      .join(', ')
    const stoMention = mentionFor('STO')
    const supMention = mentionFor('Supervisor')
    try {
      const res = await fetch('/api/slack/request-update', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ projectName: p.name, milestones: openMs, weeklyTarget: p.weekly_target, stoMention, supMention, requesterMention: REQUESTER_SLACK_MENTION }),
      })
      const data = await res.json().catch(() => ({}))
      const ok = res.ok && data.ok
      setSlackReqStatus(prev => ({ ...prev, [p.id]: ok ? 'sent' : 'error' }))
      if (ok && data.permalink) setSlackPermalinks(prev => ({ ...prev, [p.id]: data.permalink }))
      return ok
    } catch {
      setSlackReqStatus(prev => ({ ...prev, [p.id]: 'error' }))
      return false
    }
  }

  function handleRequestUpdate(p: Project) {
    const hasOpenMs = milestones.some(m => m.project_id === p.id && !m.done)
    const hasPP = p.weekly_target != null
    if (!hasOpenMs && !hasPP) { setConfirmEmptyUpdate(p); return }
    proceedRequest(p)
  }

  // Prompt for missing STO/Supervisor Slack IDs before sending; otherwise send directly
  function proceedRequest(p: Project) {
    const missing = assignments
      .filter(a => a.project_id === p.id && (a.assignment_role === 'STO' || a.assignment_role === 'Supervisor'))
      .map(a => ({ role: a.assignment_role as string, staff: staff.find(s => s.id === a.staff_id) }))
      .filter((x): x is { role: string; staff: Staff } => !!x.staff && !x.staff.slack_user_id)
    // de-dupe by staff id (keep first role seen)
    const seen = new Map<string, { id: string; name: string; role: string }>()
    for (const x of missing) {
      if (!seen.has(x.staff.id)) seen.set(x.staff.id, { id: x.staff.id, name: x.staff.name, role: x.role })
    }
    const people = Array.from(seen.values())
    if (people.length > 0) {
      setSlackIdInputs({})
      setPendingSlackIds({ project: p, people })
    } else {
      requestMilestoneUpdate(p)
    }
  }

  async function submitSlackIdsAndRequest() {
    if (!pendingSlackIds) return
    const p = pendingSlackIds.project
    const override: Record<string, string> = {}
    for (const person of pendingSlackIds.people) {
      const val = (slackIdInputs[person.id] ?? '').trim()
      if (val) {
        override[person.id] = val
        const { data } = await supabase.from('staff').update({ slack_user_id: val }).eq('id', person.id).select().single()
        if (data) setStaff(prev => prev.map(s => s.id === person.id ? data : s))
      }
    }
    setPendingSlackIds(null)
    setSlackIdInputs({})
    requestMilestoneUpdate(p, override)
  }

  // Build the overview draft. `mode` controls whether unresolved STOs show as
  // their name (preview) or a raw mention (send). Also returns the STOs missing a Slack ID.
  function buildOverview(mode: 'preview' | 'send', override: Record<string, string> = {}) {
    const active = [...projects].filter(p => p.status === 'active').sort((a, b) => a.name.localeCompare(b.name))
    const sections: string[] = []
    const missingMap = new Map<string, { id: string; name: string; role: string }>()
    // Mention (send) or name (preview); people without a Slack ID fall back to their name
    const tag = (s: Staff, role: string) => {
      const sid = override[s.id] ?? s.slack_user_id
      if (!sid) { if (!missingMap.has(s.id)) missingMap.set(s.id, { id: s.id, name: s.name, role }); return s.name }
      return mode === 'send' ? `<@${sid}>` : s.name
    }
    // Supervisors & leads on the projects included in the message → cc'd at the bottom
    const ccStaff = new Map<string, { staff: Staff; role: string }>()
    for (const p of active) {
      const openMs = [...milestones.filter(m => m.project_id === p.id && !m.done)]
        .sort((a, b) => priorityRank(a.priority) - priorityRank(b.priority) || a.created_at.localeCompare(b.created_at))
      const hasPP = p.weekly_target != null
      // Skip projects with nothing to report: no open milestones and no PP (unset or 0)
      if (openMs.length === 0 && !p.weekly_target) continue
      const stoTag = assignments
        .filter(a => a.project_id === p.id && a.assignment_role === 'STO')
        .map(a => staff.find(s => s.id === a.staff_id))
        .filter((s): s is Staff => !!s)
        .map(s => tag(s, 'STO'))
        .join(', ')
      assignments
        .filter(a => a.project_id === p.id && (a.assignment_role === 'Lead' || a.assignment_role === 'Supervisor'))
        .forEach(a => {
          const s = staff.find(x => x.id === a.staff_id)
          if (s && !ccStaff.has(s.id)) ccStaff.set(s.id, { staff: s, role: a.assignment_role as string })
        })
      let sec = `*${p.emoji || '📁'} ${p.name}*${stoTag ? `  (STO: ${stoTag})` : ''}`
      if (openMs.length > 0) {
        sec += '\n' + openMs.map(m => `  • ${m.priority} — ${m.title}${m.due_date ? ` (due ${m.due_date})` : ''}`).join('\n')
      } else {
        sec += '\n  • _No open milestones_'
      }
      sec += `\n  _Production plan this week: ${hasPP ? `${p.weekly_target} task${p.weekly_target === 1 ? '' : 's'}` : 'not defined'}_`
      sections.push(sec)
    }
    if (sections.length === 0) return null
    // Always cc the requester (skip them in the list if they're also a lead/supervisor)
    const ccLine = [
      ...[...ccStaff.values()]
        .filter(({ staff: s }) => `<@${s.slack_user_id}>` !== REQUESTER_SLACK_MENTION)
        .sort((a, b) => a.staff.name.localeCompare(b.staff.name))
        .map(({ staff: s, role }) => tag(s, role)),
      mode === 'send' ? REQUESTER_SLACK_MENTION : 'you',
    ].join(', ')
    // Weeks start on Tuesday → most recent Tuesday on or before today
    const weekStart = new Date()
    weekStart.setDate(weekStart.getDate() - ((weekStart.getDay() - 2 + 7) % 7))
    const text = [
      `*Week of ${weekStart.getMonth() + 1}/${weekStart.getDate()}*`,
      ':mega: *Code Pod — Milestones & Production Plan overview*',
      '',
      sections.join('\n\n'),
      '',
      'Any questions? Drop them in this thread. :point_down:',
      ...(ccLine ? ['', `cc: ${ccLine}`] : []),
    ].join('\n')
    return { text, missing: Array.from(missingMap.values()) }
  }

  function giveTeamOverview() {
    const built = buildOverview('preview')
    if (!built) { setOverviewStatus('empty'); return }
    setOverviewIdInputs({})
    setOverviewStatus('idle')
    setOverviewOpen(true)
  }

  async function sendOverview() {
    // Persist any Slack IDs the user filled in, then post with real mentions
    const override: Record<string, string> = {}
    for (const [id, raw] of Object.entries(overviewIdInputs)) {
      const val = raw.trim()
      if (val) {
        override[id] = val
        const { data } = await supabase.from('staff').update({ slack_user_id: val }).eq('id', id).select().single()
        if (data) setStaff(prev => prev.map(s => s.id === id ? data : s))
      }
    }
    const built = buildOverview('send', override)
    if (!built) { setOverviewOpen(false); setOverviewStatus('empty'); return }
    setOverviewOpen(false)
    setOverviewStatus('sending')
    setOverviewPermalink(null)
    try {
      const res = await fetch('/api/slack/overview', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ text: built.text }),
      })
      const data = await res.json().catch(() => ({}))
      if (res.ok && data.ok) {
        setOverviewStatus('sent')
        if (data.permalink) setOverviewPermalink(data.permalink)
      } else {
        setOverviewStatus('error')
      }
    } catch {
      setOverviewStatus('error')
    }
  }

  // --- Confirm assignments with leads ---
  // Each active project's owner = Lead → else Supervisor → else STO (acting as lead)
  function leadsForConfirmation() {
    const activeProjects = projects.filter(p => p.status === 'active')
    const map = new Map<string, { id: string; name: string; projects: string[] }>()
    const add = (id: string, name: string, project: string) => {
      if (!map.has(id)) map.set(id, { id, name, projects: [] })
      map.get(id)!.projects.push(project)
    }
    activeProjects.forEach(p => {
      // Hard-coded leads for specific projects (replace the usual Lead → Supervisor → STO pick)
      const forced = LEAD_CONFIRMATION_OVERRIDES[p.name.trim().toLowerCase()]
      if (forced) {
        forced.forEach(name => {
          const s = staff.find(x => x.name.trim().toLowerCase() === name.toLowerCase())
          add(s ? s.id : `override:${name}`, s ? s.name : name, p.name)
        })
        return
      }
      const pa = assignments.filter(a => a.project_id === p.id)
      const tierRole = pa.some(a => a.assignment_role === 'Lead') ? 'Lead'
        : pa.some(a => a.assignment_role === 'Supervisor') ? 'Supervisor'
        : pa.some(a => a.assignment_role === 'STO') ? 'STO'
        : null
      if (!tierRole) return
      pa.filter(a => a.assignment_role === tierRole).forEach(h => {
        const s = staff.find(x => x.id === h.staff_id)
        if (!s) return
        add(s.id, s.name, p.name)
      })
    })
    return Array.from(map.values())
      .map(l => ({ ...l, projects: [...new Set(l.projects)].sort((a, b) => a.localeCompare(b)) }))
      .sort((a, b) => a.name.localeCompare(b.name))
  }
  type LeadEntry = { id: string; name: string; projects: string[] }
  function buildLeadText(leads: LeadEntry[], override: Record<string, string>, mode: 'preview' | 'send') {
    const lines: string[] = ['Hello team — please confirm these are all the active queues your team is working on!', '']
    leads.forEach(l => {
      const sid = override[l.id] ?? staff.find(s => s.id === l.id)?.slack_user_id
      const label = mode === 'send'
        ? (sid ? `<@${sid}>` : `*${l.name}*`)
        : `${l.name}${sid ? '' : '  (no Slack ID)'}`
      lines.push(`${label}:`)
      l.projects.forEach(p => lines.push(`  • ${p}`))
      lines.push('')
    })
    // FYI summary of projects that aren't active yet / anymore (names only — nobody is tagged)
    const ownerNames = (p: Project) => {
      const forced = LEAD_CONFIRMATION_OVERRIDES[p.name.trim().toLowerCase()]
      if (forced) return forced
      const pa = assignments.filter(a => a.project_id === p.id)
      const tier = ['Lead', 'Supervisor', 'STO'].find(r => pa.some(a => a.assignment_role === r))
      return tier ? pa.filter(a => a.assignment_role === tier).map(a => staff.find(s => s.id === a.staff_id)?.name).filter((n): n is string => !!n) : []
    }
    const fyi: [string, string, string][] = [
      ['starting-soon', ':hourglass_flowing_sand:', 'Starting soon'],
      ['paused', ':double_vertical_bar:', 'Paused'],
    ]
    fyi.forEach(([status, emoji, title]) => {
      const list = projects.filter(p => p.status === status).sort((a, b) => a.name.localeCompare(b.name))
      if (list.length === 0) return
      lines.push(`${emoji} *${title}* (${list.length})`)
      list.forEach(p => { const o = ownerNames(p); lines.push(`  • ${p.name}${o.length ? ` — ${o.join(', ')}` : ''}`) })
      lines.push('')
    })
    lines.push('If it looks accurate then please confirm! If a project or assignment is missing please comment on the thread.')
    lines.push('')
    lines.push(`cc: ${mode === 'send' ? REQUESTER_SLACK_MENTION : 'you'}`)
    return lines.join('\n')
  }
  function handleConfirmLeads() {
    const leads = leadsForConfirmation()
    if (leads.length === 0) { setLeadStatus('empty'); return }
    const missing = leads.filter(l => !l.id.startsWith('override:') && !staff.find(s => s.id === l.id)?.slack_user_id)
    if (missing.length > 0) {
      setLeadIdInputs({})
      setLeadIdPrompt(missing.map(l => ({ id: l.id, name: l.name })))
    } else {
      setLeadPreview({ leads, override: {} })
    }
  }
  async function submitLeadIds(skip: boolean) {
    const override: Record<string, string> = {}
    if (!skip && leadIdPrompt) {
      for (const l of leadIdPrompt) {
        const v = (leadIdInputs[l.id] ?? '').trim()
        if (v) {
          override[l.id] = v
          const { data } = await supabase.from('staff').update({ slack_user_id: v }).eq('id', l.id).select().single()
          if (data) setStaff(prev => prev.map(s => s.id === l.id ? data : s))
        }
      }
    }
    setLeadIdPrompt(null)
    setLeadIdInputs({})
    setLeadPreview({ leads: leadsForConfirmation(), override })
  }
  async function sendLeadConfirmation() {
    if (!leadPreview) return
    setLeadStatus('sending')
    try {
      const res = await fetch('/api/slack/confirm-leads', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ text: buildLeadText(leadPreview.leads, leadPreview.override, 'send') }),
      })
      const data = await res.json().catch(() => ({}))
      if (res.ok && data.ok) {
        setLeadStatus('sent')
        setLeadPermalink(data.permalink ?? null)
        setLeadPreview(null)
      } else {
        setLeadStatus('error')
      }
    } catch {
      setLeadStatus('error')
    }
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
      else if (col === 'type')     { valA = a.project_type ?? '\uffff'; valB = b.project_type ?? '\uffff' }
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

  function toggleSwitch(on: boolean, onToggle: () => void, label: string) {
    return (
      <button onClick={onToggle} className="flex items-center gap-2 group">
        <span className={`text-xs transition-colors ${on ? 'text-emerald-400' : 'text-gray-500 group-hover:text-gray-300'}`}>{label}</span>
        <span className={`relative w-9 h-5 rounded-full transition-colors shrink-0 ${on ? 'bg-emerald-500' : 'bg-gray-700'}`}>
          <span className={`absolute top-0.5 left-0.5 w-4 h-4 rounded-full bg-white shadow transition-transform ${on ? 'translate-x-4' : ''}`} />
        </span>
      </button>
    )
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
    if (role === 'Lead') return 'bg-blue-500/10 text-blue-400'
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

  // Project Type dropdown cell — pick, add a new value, or manage the list
  function projectTypeSelect(p: Project) {
    return (
      <select
        value={p.project_type ?? ''}
        onChange={e => handleProjectTypeSelect(e.target.value, v => updateProjectType(p.id, v))}
        title="Project type"
        className={`text-xs px-2 py-1 rounded-md border bg-gray-900 cursor-pointer focus:outline-none focus:ring-1 focus:ring-gray-500 max-w-[160px] ${
          p.project_type ? 'border-gray-700 text-gray-200' : 'border-dashed border-gray-700 text-gray-600'
        }`}
      >
        <option value="">— None —</option>
        {projectTypeOptions().map(t => <option key={t} value={t}>{t}</option>)}
        <option disabled>──────────</option>
        <option value="__add__">＋ Add new type…</option>
        <option value="__manage__">⚙ Manage types…</option>
      </select>
    )
  }

  // Project emoji picker (Milestones summarized view + Consolidated Manager)
  function renderProjectEmojiPicker(p: Project) {
    return (
            <div className="relative">
              <button onClick={() => setEmojiPickerProjectId(emojiPickerProjectId === p.id ? null : p.id)} title="Set project emoji" className="text-lg leading-none hover:scale-110 transition-transform">
                {p.emoji || '📁'}
              </button>
              {emojiPickerProjectId === p.id && (
                <>
                  <div className="fixed inset-0 z-10" onClick={() => setEmojiPickerProjectId(null)} />
                  <div className="absolute left-0 top-8 z-20 bg-gray-900 border border-gray-700 rounded-lg p-2 shadow-xl w-60">
                    <div className="grid grid-cols-6 gap-1 max-h-48 overflow-y-auto">
                      {PROJECT_EMOJIS.map(em => (
                        <button key={em} onClick={() => { updateProjectEmoji(p.id, em); setEmojiPickerProjectId(null) }} className="text-lg hover:bg-gray-800 rounded p-1">{em}</button>
                      ))}
                    </div>
                    <div className="flex gap-1 mt-2 pt-2 border-t border-gray-800">
                      <input
                        className={inputSmClass + ' flex-1'}
                        placeholder="Paste any emoji…"
                        value={customEmoji}
                        onChange={e => setCustomEmoji(e.target.value)}
                        onKeyDown={e => { if (e.key === 'Enter' && customEmoji.trim()) { updateProjectEmoji(p.id, customEmoji.trim()); setCustomEmoji(''); setEmojiPickerProjectId(null) } }}
                      />
                      <button className="text-xs px-2 py-1 rounded bg-gray-700 text-gray-300 hover:text-white transition-colors" onClick={() => { if (customEmoji.trim()) { updateProjectEmoji(p.id, customEmoji.trim()); setCustomEmoji(''); setEmojiPickerProjectId(null) } }}>Set</button>
                    </div>
                    <button onClick={() => { updateProjectEmoji(p.id, null); setEmojiPickerProjectId(null) }} className="w-full text-xs text-gray-500 hover:text-gray-300 mt-2">Reset to default</button>
                  </div>
                </>
              )}
            </div>
          )
  }

  // Production-plan stepper (Milestones tab + Consolidated Manager)
  function ppStepper(p: Project) {
    return (
            <div className="inline-flex items-center border border-gray-700 rounded-lg overflow-hidden">
              <button onClick={() => stepWeeklyTarget(p, -1)} className="px-2.5 py-1 text-gray-400 hover:text-white hover:bg-gray-800 transition-colors leading-none">−</button>
              <input
                type="number"
                min="0"
                className="no-spin w-14 bg-transparent text-center text-sm text-gray-100 focus:outline-none py-1"
                placeholder="—"
                value={weeklyDrafts[p.id] ?? (p.weekly_target != null ? String(p.weekly_target) : '')}
                onChange={e => setWeeklyDrafts(prev => ({ ...prev, [p.id]: e.target.value }))}
                onBlur={e => updateWeeklyTarget(p.id, e.target.value)}
                onKeyDown={e => { if (e.key === 'Enter') (e.target as HTMLInputElement).blur() }}
              />
              <button onClick={() => stepWeeklyTarget(p, 1)} className="px-2.5 py-1 text-gray-400 hover:text-white hover:bg-gray-800 transition-colors leading-none">+</button>
            </div>
          )
  }

  // Milestone checklist + "add milestone" row for one project (Milestones tab + Consolidated Manager)
  function renderMilestoneList(p: Project) {
    const today = fmt(new Date())
    const ms = [...milestones.filter(m => m.project_id === p.id)].sort((a, b) =>
      (a.done === b.done ? 0 : a.done ? 1 : -1) || priorityRank(a.priority) - priorityRank(b.priority) || a.created_at.localeCompare(b.created_at)
    )
    const visibleMs = showCompletedMilestones ? ms : ms.filter(m => !m.done)
    const draft = getMilestoneDraft(p.id)
    return (
      <>
                      {visibleMs.length > 0 && (
                        <div className="space-y-1.5 mb-3">
                          {visibleMs.map(m => {
                            const overdue = !m.done && m.due_date && m.due_date < today
                            return (
                              <div key={m.id} className="flex items-center gap-3 bg-gray-800/40 rounded-lg px-3 py-2 text-sm">
                                {editingMilestoneId === m.id ? (
                                  <>
                                    <input
                                      className={inputSmClass + ' flex-1'}
                                      value={editMilestone.title}
                                      placeholder="Milestone name"
                                      autoFocus
                                      onChange={e => setEditMilestone({ ...editMilestone, title: e.target.value })}
                                      onKeyDown={e => e.key === 'Enter' && saveMilestone(m.id)}
                                    />
                                    <input type="date" className={inputSmClass + ' shrink-0'} style={{ width: 150 }} value={editMilestone.due_date} onChange={e => setEditMilestone({ ...editMilestone, due_date: e.target.value })} />
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
                                      {PRIORITIES.map(pr => <option key={pr} value={pr} className="bg-gray-900 text-gray-100">{priorityLabel(pr)}</option>)}
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
                            {PRIORITIES.map(pr => <option key={pr} value={pr}>{priorityLabel(pr)}</option>)}
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
      </>
    )
  }

  // Drag-and-drop staff pools: Unassigned / Flexed / OOO / Onboarding (Assignments + Consolidated Manager)
  function renderStaffPools(widthClass = 'w-52') {
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
                <div className={`${widthClass} shrink-0 sticky top-6 flex flex-col gap-5`}>

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
                    onDrop={e => {
                      e.preventDefault(); setDragOverFlexed(false)
                      const type = e.dataTransfer.getData('type')
                      const id = e.dataTransfer.getData('id')
                      if (type === 'staff' && id) {
                        const member = staff.find(s => s.id === id)
                        setFlexNote(member?.flex_notes ?? '')
                        setPendingFlex({ staffId: id, assignmentId: null })
                        setDraggedStaffId(null)
                      } else if (type === 'assignment' && id) {
                        const assignment = assignments.find(a => a.id === id)
                        if (assignment) {
                          const member = staff.find(s => s.id === assignment.staff_id)
                          setFlexNote(member?.flex_notes ?? '')
                          setPendingFlex({ staffId: assignment.staff_id, assignmentId: id })
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
  }

  // Role pickers shown after dropping a person / moving an assignment onto a project
  function renderDropModals() {
    return (
      <>
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
                    <option>Lead</option>
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
                        if (member?.flexed || member?.onboarding) {
                          const { data } = await supabase.from('staff').update({ flexed: false, onboarding: false }).eq('id', pendingDrop.staffId).select().single()
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
                      <option>Lead</option>
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

      </>
    )
  }

  // Inline rename of a project name (used in the Assignments tab, both views)
  function renderProjectName(p: Project, textClass: string) {
    if (renameProjectId === p.id) {
      return (
        <input
          autoFocus
          className={inputSmClass}
          style={{ minWidth: 180, width: 'auto' }}
          value={renameValue}
          onChange={e => setRenameValue(e.target.value)}
          onClick={e => e.stopPropagation()}
          onKeyDown={e => {
            if (e.key === 'Enter') saveProjectName(p.id)
            else if (e.key === 'Escape') setRenameProjectId(null)
          }}
          onBlur={() => saveProjectName(p.id)}
        />
      )
    }
    return (
      <span
        className={textClass + ' cursor-text hover:underline decoration-dotted decoration-gray-600 underline-offset-2'}
        title="Click to rename"
        onClick={() => { setRenameProjectId(p.id); setRenameValue(p.name) }}
      >
        {p.name}
      </span>
    )
  }

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
      <div style={{ transform: `rotate(${logoSpins * 360}deg)`, transition: 'transform 1s ease-in-out' }}>
      <div className="border-b px-8 py-5 flex items-center justify-between" style={{ borderColor: '#193a29' }}>
        <div className="flex items-center gap-3">
          {/* eslint-disable-next-line @next/next/no-img-element */}
          <img
            src="/logo.png"
            alt="Code Pod logo"
            onClick={() => setLogoSpins(n => n + 1)}
            className="h-10 w-auto shrink-0 cursor-pointer hover:scale-110 active:scale-95 transition-transform select-none"
          />
          <h1 className="text-xl font-semibold text-gray-100 tracking-tight">Code Pod Project Manager</h1>
        </div>
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

      <div className={`${tab === 'orgchart' ? 'max-w-none' : tab === 'consolidated' || tab === 'utilization' ? 'max-w-[1720px]' : 'max-w-4xl'} mx-auto px-8 py-8`}>
        {/* Tabs */}
        <div className="flex flex-wrap gap-1 mb-8 bg-gray-900 rounded-xl p-1 w-fit max-w-full">
          {([
            { key: 'dashboard', label: 'Dashboard' },
            { key: 'projects', label: 'Projects' },
            { key: 'staff', label: 'Staff' },
            { key: 'assignments', label: 'Assignments' },
            { key: 'scenarios', label: 'Staffing Scenarios' },
            { key: 'milestones', label: 'Milestones' },
            { key: 'orgchart', label: 'Org Chart' },
            { key: 'consolidated', label: '✨ Consolidated Manager' },
            { key: 'utilization', label: 'Utilization' },
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
                <select
                  className={selectClass}
                  value={newProject.project_type}
                  onChange={e => handleProjectTypeSelect(e.target.value, v => setNewProject(np => ({ ...np, project_type: v ?? '' })))}
                >
                  <option value="">Project type…</option>
                  {projectTypeOptions().map(t => <option key={t} value={t}>{t}</option>)}
                  <option disabled>──────────</option>
                  <option value="__add__">＋ Add new type…</option>
                  <option value="__manage__">⚙ Manage types…</option>
                </select>
                <label className="flex items-center gap-2 text-sm text-gray-400 cursor-pointer select-none">
                  <input type="checkbox" checked={newProject.is_pilot} onChange={e => setNewProject({ ...newProject, is_pilot: e.target.checked })} className="accent-[#193a29] w-4 h-4" />
                  Pilot
                </label>
                <label className="flex items-center gap-2 text-sm text-gray-400 cursor-pointer select-none">
                  <input type="checkbox" checked={newProject.is_internal} onChange={e => setNewProject({ ...newProject, is_internal: e.target.checked })} className="accent-[#193a29] w-4 h-4" />
                  Internal
                </label>
                <button className={btnPrimary} style={{ backgroundColor: '#193a29' }} onClick={addProject}>Add</button>
              </div>
            </div>

            <div className="mb-4 flex items-center gap-2">
              <div className="relative flex-1">
                <input
                  className={inputClass + ' w-full pl-9'}
                  placeholder="Search project names…"
                  value={projectSearch}
                  onChange={e => setProjectSearch(e.target.value)}
                />
                <span className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-500 text-sm pointer-events-none">🔍</span>
                {projectSearch && (
                  <button onClick={() => setProjectSearch('')} className="absolute right-3 top-1/2 -translate-y-1/2 text-gray-500 hover:text-gray-300 text-sm">✕</button>
                )}
              </div>
              <button
                onClick={() => setProjectCsvPrompt(true)}
                className="text-xs px-3 py-2 rounded-lg border border-gray-700 text-gray-400 hover:text-gray-200 transition-colors whitespace-nowrap"
              >
                ⬇ Export CSV
              </button>
            </div>

            {projects.length === 0 ? (
              <p className="text-gray-600 text-sm">No projects yet.</p>
            ) : (
              <>
              <table className="w-full text-sm">
                <thead>
                  <tr className="text-left text-gray-500 border-b border-gray-800">
                    <th className="pb-3 pr-3 font-medium text-xs uppercase tracking-wider w-8">#</th>
                    <th className="pb-3 pr-3 font-medium text-xs uppercase tracking-wider w-12">Pilot</th>
                    <th className="pb-3 pr-3 font-medium text-xs uppercase tracking-wider w-14">Internal</th>
                    {[
                      { key: 'name', label: 'Name' },
                      { key: 'customer', label: 'Customer' },
                      { key: 'type', label: 'Project Type' },
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
                  {sortedProjects(projects.filter(p => visibleStatuses[p.status] && p.name.toLowerCase().includes(projectSearch.trim().toLowerCase()))).map((p, idx) => {
                    const count = assignments.filter(a => a.project_id === p.id).length
                    const isEditing = editingProjectId === p.id
                    return (
                      <tr key={p.id} className="border-b border-gray-800/60 hover:bg-gray-900/60 transition-colors">
                        <td className="py-3.5 pr-3 text-gray-600 text-xs tabular-nums align-top">{idx + 1}</td>
                        <td className="py-3.5 pr-3 align-top">
                          <input
                            type="checkbox"
                            checked={p.is_pilot}
                            onChange={e => updateProjectPilot(p.id, e.target.checked)}
                            title="Pilot project"
                            className="accent-[#193a29] w-4 h-4 cursor-pointer"
                          />
                        </td>
                        <td className="py-3.5 pr-3 align-top">
                          <input
                            type="checkbox"
                            checked={p.is_internal}
                            onChange={e => updateProjectInternal(p.id, e.target.checked)}
                            title="Internal project (no external customer)"
                            className="accent-[#193a29] w-4 h-4 cursor-pointer"
                          />
                        </td>
                        {isEditing ? (
                          <>
                            <td className="py-2 pr-2"><input className={inputSmClass} value={editProject.name} onChange={e => setEditProject({ ...editProject, name: e.target.value })} /></td>
                            <td className="py-2 pr-2"><input className={inputSmClass} value={editProject.customer_codename} onChange={e => setEditProject({ ...editProject, customer_codename: e.target.value })} /></td>
                            <td className="py-2 pr-2">{projectTypeSelect(p)}</td>
                            <td className="py-2 pr-2">
                              <select className={selectSmClass} value={editProject.status} onChange={e => setEditProject({ ...editProject, status: e.target.value })}>
                                <option value="active">Active</option>
                                <option value="starting-soon">Starting Soon</option>
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
                            <td className="py-3.5 pr-3">{projectTypeSelect(p)}</td>
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
              <p className="text-xs text-gray-500 mt-4">Total projects in list: <span className="text-gray-300 font-medium">{projects.filter(p => visibleStatuses[p.status] && p.name.toLowerCase().includes(projectSearch.trim().toLowerCase())).length}</span></p>
              </>
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
                  <select
                    value={positionFilter}
                    onChange={e => setPositionFilter(e.target.value)}
                    title="Filter by position"
                    className={`text-xs px-3 py-1.5 rounded-lg border transition-colors cursor-pointer ${
                      positionFilter ? 'text-emerald-400' : 'border-gray-700 text-gray-400 hover:text-gray-200'
                    }`}
                    style={positionFilter ? { borderColor: '#193a29' } : {}}
                  >
                    <option value="">All positions</option>
                    <option value="__none__">No position</option>
                    {VALID_POSITIONS.map(pos => <option key={pos} value={pos}>{pos}</option>)}
                  </select>
                  {toggleSwitch(hideAssigned, () => setHideAssigned(!hideAssigned), 'Unassigned only')}
                  <label className="cursor-pointer text-xs text-gray-400 hover:text-gray-200 transition-colors border border-gray-700 rounded-lg px-3 py-1.5">
                    Import CSV
                    <input type="file" accept=".csv" className="hidden" onChange={importStaffCSV} />
                  </label>
                  <button
                    onClick={exportStaffCSV}
                    className="text-xs text-gray-400 hover:text-gray-200 transition-colors border border-gray-700 rounded-lg px-3 py-1.5"
                  >
                    ⬇ Export CSV
                  </button>
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
                  {sortedList(staff.filter(s => (!hideAssigned || !assignments.some(a => a.staff_id === s.id)) && (!positionFilter || (positionFilter === '__none__' ? !s.position : s.position === positionFilter))), staffSort).map(s => {
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
                                {VALID_POSITIONS.map(pos => <option key={pos}>{pos}</option>)}
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
                              <input
                                type="text"
                                className={inputSmClass + ' mt-1'}
                                placeholder="Slack user ID (e.g. U0123ABCD)"
                                value={editStaff.slack_user_id}
                                onChange={e => setEditStaff({ ...editStaff, slack_user_id: e.target.value })}
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
                                  <option>Lead</option>
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
                                    if (!project) return null
                                    if (editingStaffAssignmentId === a.id) {
                                      return (
                                        <select
                                          key={a.id}
                                          autoFocus
                                          value={a.project_id}
                                          onChange={e => { updateAssignmentProject(a.id, e.target.value); setEditingStaffAssignmentId(null) }}
                                          onBlur={() => setEditingStaffAssignmentId(null)}
                                          className={selectSmClass}
                                          style={{ width: 'auto' }}
                                        >
                                          {[...projects].sort((p1, p2) => p1.name.localeCompare(p2.name)).map(p => <option key={p.id} value={p.id}>{p.name}</option>)}
                                        </select>
                                      )
                                    }
                                    return (
                                      <button
                                        key={a.id}
                                        onClick={() => setEditingStaffAssignmentId(a.id)}
                                        title="Click to move to another project"
                                        className="text-xs px-2 py-0.5 rounded-full bg-gray-800 text-gray-300 flex items-center gap-1 hover:bg-gray-700 hover:ring-1 hover:ring-gray-600 transition-colors cursor-pointer"
                                      >
                                        {project.name}
                                        {a.assignment_role && (
                                          <span className={`px-1.5 py-0.5 rounded-full text-[10px] font-medium ${roleColor(a.assignment_role)}`}>{a.assignment_role}</span>
                                        )}
                                      </button>
                                    )
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
              <div className="flex items-start justify-between gap-4 mb-4">
                <h2 className="text-xs font-semibold text-gray-400 uppercase tracking-wider shrink-0 pt-1.5">Assign Staff to Project</h2>
                <div className="flex flex-wrap items-center justify-end gap-2">
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
                {toggleSwitch(hideAssignedInDropdown, () => setHideAssignedInDropdown(!hideAssignedInDropdown), 'Unassigned only')}
                <select
                  value={supervisorFilter}
                  onChange={e => setSupervisorFilter(e.target.value)}
                  title="Filter projects by supervisor"
                  className={`text-xs px-3 py-1.5 rounded-lg border transition-colors cursor-pointer ${
                    supervisorFilter ? 'text-emerald-400' : 'border-gray-700 text-gray-400 hover:text-gray-200'
                  }`}
                  style={supervisorFilter ? { borderColor: '#193a29' } : {}}
                >
                  <option value="">All supervisors</option>
                  <option value="__none__">No supervisor</option>
                  {staff
                    .filter(s => assignments.some(a => a.staff_id === s.id && a.assignment_role === 'Supervisor'))
                    .sort((a, b) => a.name.localeCompare(b.name))
                    .map(s => <option key={s.id} value={s.id}>{s.name}</option>)}
                </select>
                {toggleSwitch(groupBySupervisor, () => setGroupBySupervisor(v => !v), 'Group by lead')}
                <button
                  onClick={handleConfirmLeads}
                  disabled={leadStatus === 'sending'}
                  className={`text-xs px-3 py-1.5 rounded-lg border transition-colors whitespace-nowrap ${
                    leadStatus === 'sent' ? 'border-emerald-500/40 text-emerald-400' :
                    leadStatus === 'error' || leadStatus === 'empty' ? 'border-red-500/40 text-red-400' :
                    'border-gray-700 text-gray-300 hover:text-white hover:border-gray-500'
                  }`}
                >
                  {leadStatus === 'sending' ? 'Posting…' :
                   leadStatus === 'sent' ? '✓ Sent to #code_leads_plus' :
                   leadStatus === 'error' ? '⚠ Failed — retry' :
                   leadStatus === 'empty' ? 'No leads on active projects' :
                   '✅ Confirm assignments with leads'}
                </button>
                {leadStatus === 'sent' && leadPermalink && (
                  <a href={leadPermalink} target="_blank" rel="noopener noreferrer" className="text-xs text-sky-400 hover:text-sky-300 whitespace-nowrap">View ↗</a>
                )}
                {(() => {
                  const n = inactiveAssignmentIds().length
                  return n > 0 ? (
                    <button
                      onClick={() => setConfirmClearInactive(true)}
                      className="text-xs px-3 py-1.5 rounded-lg border border-red-500/30 text-red-400 hover:bg-red-500/10 transition-colors whitespace-nowrap"
                    >
                      Clear staffing from inactive ({n})
                    </button>
                  ) : null
                })()}
                {assignments.length > 0 && (
                  <button
                    onClick={() => setConfirmClearAll(true)}
                    className="text-xs px-3 py-1.5 rounded-lg border border-red-500/30 text-red-400 hover:bg-red-500/10 transition-colors whitespace-nowrap"
                  >
                    Clear all assignments ({assignments.length})
                  </button>
                )}
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
                  <option>Lead</option>
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

            {/* Quick add-project between the form and the project list */}
            <div className="mb-4">
              {showProjectAddBetween ? (
                <div className="bg-gray-900 rounded-xl p-4 border border-gray-800 flex flex-wrap gap-2 items-center">
                  <input className={inputClass} placeholder="Project name *" value={newProject.name} autoFocus onChange={e => setNewProject({ ...newProject, name: e.target.value })} onKeyDown={e => { if (e.key === 'Enter') { addProject(); setShowProjectAddBetween(false) } }} />
                  <input className={inputClass} placeholder="Customer Codename" value={newProject.customer_codename} onChange={e => setNewProject({ ...newProject, customer_codename: e.target.value })} />
                  <select className={selectClass} value={newProject.status} onChange={e => setNewProject({ ...newProject, status: e.target.value })}>
                    {STATUS_OPTIONS.map(o => <option key={o.value} value={o.value}>{o.label}</option>)}
                  </select>
                  <input className={inputClass} type="number" placeholder="Expected Duration (Weeks)" value={newProject.duration_weeks} onChange={e => setNewProject({ ...newProject, duration_weeks: e.target.value })} style={{ width: 200 }} />
                  <button className={btnPrimary} style={{ backgroundColor: '#193a29' }} onClick={() => { addProject(); setShowProjectAddBetween(false) }}>Add</button>
                  <button className={btnCancel} onClick={() => setShowProjectAddBetween(false)}>Cancel</button>
                </div>
              ) : (
                <button
                  onClick={() => setShowProjectAddBetween(true)}
                  className="w-full border border-dashed border-gray-700 rounded-xl py-2.5 text-gray-500 hover:text-gray-300 hover:border-gray-500 text-sm transition-colors"
                >
                  + Add a project
                </button>
              )}
            </div>

            <div className="mb-4 relative">
              <input
                className={inputClass + ' w-full pl-9'}
                placeholder="Search project names…"
                value={assignmentSearch}
                onChange={e => setAssignmentSearch(e.target.value)}
              />
              <span className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-500 text-sm pointer-events-none">🔍</span>
              {assignmentSearch && (
                <button onClick={() => setAssignmentSearch('')} className="absolute right-3 top-1/2 -translate-y-1/2 text-gray-500 hover:text-gray-300 text-sm">✕</button>
              )}
            </div>

            {/* Two-column layout */}
            <div className="flex gap-6 items-start">

            {/* Left: project list */}
            <div className="flex-1 min-w-0">
            {groupBySupervisor ? (() => {
              const visibleProjects = [...projects].filter(p => visibleStatuses[p.status] && p.name.toLowerCase().includes(assignmentSearch.trim().toLowerCase()))
              const groupsMap = new Map<string, { key: string; name: string; items: { project: Project; members: Assignment[]; viaRole: string | null }[] }>()
              const ensure = (key: string, name: string) => {
                if (!groupsMap.has(key)) groupsMap.set(key, { key, name, items: [] })
                return groupsMap.get(key)!
              }
              visibleProjects.forEach(p => {
                const projAssignments = assignments.filter(a => a.project_id === p.id)
                // Grouping tier: Lead → Supervisor → STO
                const tierRole = projAssignments.some(a => a.assignment_role === 'Lead') ? 'Lead'
                  : projAssignments.some(a => a.assignment_role === 'Supervisor') ? 'Supervisor'
                  : projAssignments.some(a => a.assignment_role === 'STO') ? 'STO'
                  : null
                if (tierRole) {
                  const headers = projAssignments.filter(a => a.assignment_role === tierRole)
                  headers.forEach(h => {
                    const st = staff.find(s => s.id === h.staff_id)
                    const members = projAssignments.filter(a => a.id !== h.id && a.assignment_role !== tierRole)
                    ensure(h.staff_id, st?.name ?? 'Unknown').items.push({ project: p, members, viaRole: tierRole === 'Lead' ? null : tierRole })
                  })
                } else {
                  ensure('__none__', 'No lead / supervisor / STO').items.push({ project: p, members: projAssignments, viaRole: null })
                }
              })
              let groups = Array.from(groupsMap.values())
              if (supervisorFilter) groups = groups.filter(g => g.key === supervisorFilter)
              groups.sort((a, b) => a.key === '__none__' ? 1 : b.key === '__none__' ? -1 : a.name.localeCompare(b.name))
              if (groups.length === 0) return <p className="text-gray-600 text-sm">No leads to show.</p>
              return (
                <div className="space-y-3">
                  {groups.map(g => {
                    const totalMembers = g.items.reduce((n, it) => n + it.members.length, 0)
                    return (
                      <div key={g.key} className="border border-gray-800 rounded-xl p-5 bg-gray-900/40">
                        <div className="flex items-center justify-between mb-4">
                          {g.key === '__none__' ? (
                            <h3 className="font-semibold text-gray-100">⚠️ No lead / supervisor / STO</h3>
                          ) : (
                            <h3 className="font-semibold text-gray-100 flex items-center gap-2">
                              <span className="relative">
                                <button
                                  onClick={() => setEmojiPickerStaffId(emojiPickerStaffId === g.key ? null : g.key)}
                                  title="Set supervisor emoji"
                                  className="text-lg leading-none hover:scale-110 transition-transform"
                                >
                                  {staff.find(s => s.id === g.key)?.emoji || '👤'}
                                </button>
                                {emojiPickerStaffId === g.key && (
                                  <>
                                    <div className="fixed inset-0 z-10" onClick={() => setEmojiPickerStaffId(null)} />
                                    <div className="absolute left-0 top-8 z-20 bg-gray-900 border border-gray-700 rounded-lg p-2 shadow-xl w-60">
                                      <div className="grid grid-cols-6 gap-1 max-h-48 overflow-y-auto">
                                        {PROJECT_EMOJIS.map(em => (
                                          <button key={em} onClick={() => { updateStaffEmoji(g.key, em); setEmojiPickerStaffId(null) }} className="text-lg hover:bg-gray-800 rounded p-1">{em}</button>
                                        ))}
                                      </div>
                                      <div className="flex gap-1 mt-2 pt-2 border-t border-gray-800">
                                        <input
                                          className={inputSmClass + ' flex-1'}
                                          placeholder="Paste any emoji…"
                                          value={customEmoji}
                                          onChange={e => setCustomEmoji(e.target.value)}
                                          onKeyDown={e => { if (e.key === 'Enter' && customEmoji.trim()) { updateStaffEmoji(g.key, customEmoji.trim()); setCustomEmoji(''); setEmojiPickerStaffId(null) } }}
                                        />
                                        <button className="text-xs px-2 py-1 rounded bg-gray-700 text-gray-300 hover:text-white transition-colors" onClick={() => { if (customEmoji.trim()) { updateStaffEmoji(g.key, customEmoji.trim()); setCustomEmoji(''); setEmojiPickerStaffId(null) } }}>Set</button>
                                      </div>
                                      <button onClick={() => { updateStaffEmoji(g.key, null); setEmojiPickerStaffId(null) }} className="w-full text-xs text-gray-500 hover:text-gray-300 mt-2">Reset to default</button>
                                    </div>
                                  </>
                                )}
                              </span>
                              {g.name}
                            </h3>
                          )}
                          <span className="text-xs text-gray-500">{g.items.length} project{g.items.length !== 1 ? 's' : ''} · {totalMembers} team</span>
                        </div>
                        <div className="space-y-3">
                          {[...g.items].sort((a, b) => a.project.name.localeCompare(b.project.name)).map(it => {
                            const isOver = dragOverProjectId === it.project.id
                            return (
                            <div
                              key={it.project.id}
                              className={`border rounded-lg p-3 transition-all ${isOver ? 'border-[#193a29] bg-[#193a29]/10' : 'border-gray-800 bg-gray-900/40'}`}
                              onDragOver={e => { e.preventDefault(); setDragOverProjectId(it.project.id) }}
                              onDragLeave={() => setDragOverProjectId(null)}
                              onDrop={e => {
                                e.preventDefault(); setDragOverProjectId(null)
                                const type = e.dataTransfer.getData('type')
                                const id = e.dataTransfer.getData('id')
                                if (type === 'staff' && id) {
                                  if (!assignments.some(x => x.project_id === it.project.id && x.staff_id === id)) setPendingDrop({ staffId: id, projectId: it.project.id })
                                  setDraggedStaffId(null)
                                } else if (type === 'assignment' && id) {
                                  const assignment = assignments.find(x => x.id === id)
                                  if (assignment && assignment.project_id !== it.project.id && !assignments.some(x => x.project_id === it.project.id && x.staff_id === assignment.staff_id)) {
                                    setPendingMove({ assignmentId: id, projectId: it.project.id })
                                    setDropRole(assignment.assignment_role ?? '')
                                  }
                                  setDraggedAssignmentId(null)
                                }
                              }}
                            >
                              <div className="flex items-center gap-2 mb-2">
                                <p className="text-sm font-medium text-gray-200 flex items-center gap-2">
                                  <span>{it.project.emoji || '📁'}</span>
                                  {renderProjectName(it.project, 'text-sm font-medium text-gray-200')}
                                </p>
                                <button
                                  onClick={() => { setAddingToProjectId(addingToProjectId === it.project.id ? null : it.project.id); setQuickAdd({ staff_id: '', assignment_role: '' }) }}
                                  title="Add a person to this project"
                                  className="w-5 h-5 flex items-center justify-center rounded-full border border-gray-600 text-gray-400 hover:text-white hover:border-gray-400 transition-colors text-sm leading-none"
                                >
                                  +
                                </button>
                              </div>
                              {addingToProjectId === it.project.id && (
                                <div className="flex flex-wrap gap-2 mb-2 p-2 rounded-lg bg-gray-800/40 border border-gray-700">
                                  <select
                                    className={selectSmClass + ' flex-1 min-w-[150px]'}
                                    value={quickAdd.staff_id}
                                    onChange={e => setQuickAdd({ ...quickAdd, staff_id: e.target.value })}
                                  >
                                    <option value="">Select staff…</option>
                                    {[...staff]
                                      .filter(s => !assignments.some(a => a.project_id === it.project.id && a.staff_id === s.id))
                                      .sort((a, b) => a.name.localeCompare(b.name))
                                      .map(s => <option key={s.id} value={s.id}>{s.name}{s.position ? ` — ${s.position}` : ''}</option>)}
                                  </select>
                                  <select
                                    className={selectSmClass + ' w-28'}
                                    value={quickAdd.assignment_role}
                                    onChange={e => setQuickAdd({ ...quickAdd, assignment_role: e.target.value })}
                                  >
                                    <option value="">Role</option>
                                    <option>Lead</option>
                                    <option>Supervisor</option>
                                    <option>STO</option>
                                    <option>Ops Support</option>
                                  </select>
                                  <button
                                    className="text-xs px-2 py-1 rounded bg-gray-700 text-gray-300 hover:text-white transition-colors whitespace-nowrap"
                                    onClick={async () => {
                                      if (!quickAdd.staff_id) return
                                      const { data } = await supabase.from('assignments').insert([{ project_id: it.project.id, staff_id: quickAdd.staff_id, assignment_role: quickAdd.assignment_role || null }]).select().single()
                                      if (data) setAssignments(prev => [...prev, data])
                                      setQuickAdd({ staff_id: '', assignment_role: '' })
                                      setAddingToProjectId(null)
                                    }}
                                  >+ Add</button>
                                </div>
                              )}
                              {it.members.length === 0 ? (
                                <p className="text-xs text-gray-600">{isOver ? 'Drop to add' : 'No non-supervisor team.'}</p>
                              ) : (
                                <div className="space-y-1.5">
                                  {[...it.members].sort((a, b) => {
                                    const order = ['Lead', 'Supervisor', 'STO', 'Ops Support']
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
                                        className={`flex items-center justify-between bg-gray-800/60 rounded-lg px-3 py-2 text-sm cursor-grab active:cursor-grabbing transition-opacity ${draggedAssignmentId === a.id ? 'opacity-40' : ''}`}
                                      >
                                        <div className="flex items-center gap-3">
                                          <span className="font-medium text-gray-200">{member?.name ?? 'Unknown'}</span>
                                          <select
                                            value={a.assignment_role ?? ''}
                                            onChange={e => updateAssignmentRole(a.id, e.target.value)}
                                            onClick={e => e.stopPropagation()}
                                            title="Change role"
                                            className={`text-xs font-medium px-2 py-0.5 pr-5 rounded-full cursor-pointer appearance-none focus:outline-none focus:ring-1 focus:ring-gray-500 ${a.assignment_role ? roleColor(a.assignment_role) : 'bg-gray-700/60 text-gray-400'}`}
                                            style={{ backgroundImage: 'none' }}
                                          >
                                            <option value="" className="bg-gray-900 text-gray-100">No role</option>
                                            <option value="Lead" className="bg-gray-900 text-gray-100">Lead</option>
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
                      </div>
                    )
                  })}
                </div>
              )
            })() : projects.length === 0 ? (
              <p className="text-gray-600 text-sm">Add projects first.</p>
            ) : (
              <div className="space-y-3">
                {[...projects].filter(p => visibleStatuses[p.status] && p.name.toLowerCase().includes(assignmentSearch.trim().toLowerCase())).filter(p => {
                  if (!supervisorFilter) return true
                  const hasSupervisor = assignments.some(a => a.project_id === p.id && a.assignment_role === 'Supervisor')
                  if (supervisorFilter === '__none__') return !hasSupervisor
                  return assignments.some(a => a.project_id === p.id && a.staff_id === supervisorFilter && a.assignment_role === 'Supervisor')
                }).sort((a, b) => statusRank(a.status) - statusRank(b.status) || a.name.localeCompare(b.name)).map(p => {
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
                          <div className="relative">
                            <button
                              onClick={() => setEmojiPickerProjectId(emojiPickerProjectId === p.id ? null : p.id)}
                              title="Set project emoji"
                              className="text-lg leading-none hover:scale-110 transition-transform"
                            >
                              {p.emoji || '📁'}
                            </button>
                            {emojiPickerProjectId === p.id && (
                              <>
                                <div className="fixed inset-0 z-10" onClick={() => setEmojiPickerProjectId(null)} />
                                <div className="absolute left-0 top-8 z-20 bg-gray-900 border border-gray-700 rounded-lg p-2 shadow-xl w-60">
                                  <div className="grid grid-cols-6 gap-1 max-h-48 overflow-y-auto">
                                    {PROJECT_EMOJIS.map(em => (
                                      <button key={em} onClick={() => { updateProjectEmoji(p.id, em); setEmojiPickerProjectId(null) }} className="text-lg hover:bg-gray-800 rounded p-1">{em}</button>
                                    ))}
                                  </div>
                                  <div className="flex gap-1 mt-2 pt-2 border-t border-gray-800">
                                    <input
                                      className={inputSmClass + ' flex-1'}
                                      placeholder="Paste any emoji…"
                                      value={customEmoji}
                                      onChange={e => setCustomEmoji(e.target.value)}
                                      onKeyDown={e => { if (e.key === 'Enter' && customEmoji.trim()) { updateProjectEmoji(p.id, customEmoji.trim()); setCustomEmoji(''); setEmojiPickerProjectId(null) } }}
                                    />
                                    <button
                                      className="text-xs px-2 py-1 rounded bg-gray-700 text-gray-300 hover:text-white transition-colors"
                                      onClick={() => { if (customEmoji.trim()) { updateProjectEmoji(p.id, customEmoji.trim()); setCustomEmoji(''); setEmojiPickerProjectId(null) } }}
                                    >Set</button>
                                  </div>
                                  <button onClick={() => { updateProjectEmoji(p.id, null); setEmojiPickerProjectId(null) }} className="w-full text-xs text-gray-500 hover:text-gray-300 mt-2">Reset to default</button>
                                </div>
                              </>
                            )}
                          </div>
                          {renderProjectName(p, 'font-semibold text-gray-100')}
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
                            <option>Lead</option>
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
                            const order = ['Lead', 'Supervisor', 'STO', 'Ops Support']
                            return (order.indexOf(a.assignment_role ?? '') === -1 ? 99 : order.indexOf(a.assignment_role ?? '')) -
                                   (order.indexOf(b.assignment_role ?? '') === -1 ? 99 : order.indexOf(b.assignment_role ?? ''))
                          }).map(a => {
                            const member = staff.find(s => s.id === a.staff_id)
                            const nonSupProjectCount = new Set(
                              assignments.filter(x => x.staff_id === a.staff_id && x.assignment_role !== 'Supervisor' && x.assignment_role !== 'Lead').map(x => x.project_id)
                            ).size
                            const overAllocated = a.assignment_role !== 'Supervisor' && a.assignment_role !== 'Lead' && nonSupProjectCount >= 2
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
                                      <button onClick={() => clearStaffOOO(member.id)} title="Clear OOO" className="ml-0.5 text-amber-400/70 hover:text-amber-300 leading-none">×</button>
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
                                    <option value="Lead" className="bg-gray-900 text-gray-100">Lead</option>
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
            {renderStaffPools()}

            </div>{/* end two-column layout */}
          </div>
        )}

        {/* Dashboard Tab */}
        {tab === 'dashboard' && ((projects: Project[], assignments: Assignment[]) => {
          const assignedStaffIds = new Set(assignments.map(a => a.staff_id))
          const availableStaff = staff.filter(s => !assignedStaffIds.has(s.id) && !s.flexed && !s.ooo && !s.onboarding)
          const oooCount = staff.filter(s => s.ooo).length
          const flexedCount = staff.filter(s => s.flexed && !assignedStaffIds.has(s.id)).length
          const onboardingCount = staff.filter(s => s.onboarding && !assignedStaffIds.has(s.id)).length

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

          const customerCounts = projects
            .filter(p => !customerActiveOnly || p.status === 'active')
            .reduce((acc, p) => {
              const key = p.customer_codename || 'No Codename'
              acc[key] = (acc[key] || 0) + 1
              return acc
            }, {} as Record<string, number>)
          const maxCustomer = Math.max(...Object.values(customerCounts), 1)

          // Active projects aggregated by project type
          const activeForType = projects.filter(p => p.status === 'active')
          const typeAgg = new Map<string, { projects: number; people: number }>()
          activeForType.forEach(p => {
            const key = p.project_type || 'No type'
            const cur = typeAgg.get(key) ?? { projects: 0, people: 0 }
            cur.projects += 1
            cur.people += assignments.filter(a => a.project_id === p.id).length
            typeAgg.set(key, cur)
          })
          const typeRows = [...typeAgg.entries()]
            .sort((a, b) => a[0] === 'No type' ? 1 : b[0] === 'No type' ? -1 : b[1].projects - a[1].projects || a[0].localeCompare(b[0]))
          const maxType = Math.max(...typeRows.map(([, v]) => v.projects), 1)

          const headcounts = projects.filter(p => p.status === 'active').map(p => {
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
              <div className="flex justify-end gap-4">
                {toggleSwitch(excludePilots, () => setExcludePilots(v => !v), 'Exclude pilots')}
                {toggleSwitch(excludeInternal, () => setExcludeInternal(v => !v), 'Exclude internal')}
              </div>
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
                <div className="bg-gray-900 border border-gray-800 rounded-xl p-5">
                  <div className="flex items-center justify-between mb-4">
                    <p className="text-xs text-gray-500 uppercase tracking-wider">Projects by Customer</p>
                    {toggleSwitch(customerActiveOnly, () => setCustomerActiveOnly(v => !v), 'Active only')}
                  </div>
                  <div className="space-y-3">
                    {Object.entries(customerCounts).sort((a, b) => b[1] - a[1]).map(([label, count]) => (
                      <div key={label}>
                        <div className="flex justify-between text-xs mb-1">
                          <span className="text-gray-300">{label}</span>
                          <span className="text-gray-500">{count}</span>
                        </div>
                        <div className="w-full bg-gray-800 rounded-full h-2">
                          <div className="h-2 rounded-full bg-violet-500" style={{ width: `${(count / maxCustomer) * 100}%` }} />
                        </div>
                      </div>
                    ))}
                  </div>
                </div>
              </div>

              {/* Active projects by type */}
              <div className="bg-gray-900 border border-gray-800 rounded-xl p-5">
                <div className="flex items-center justify-between mb-4">
                  <p className="text-xs text-gray-500 uppercase tracking-wider">Active Projects by Type</p>
                  <span className="text-xs text-gray-500">{activeForType.length} active · {typeRows.filter(([k]) => k !== 'No type').length} type{typeRows.filter(([k]) => k !== 'No type').length === 1 ? '' : 's'}</span>
                </div>
                {typeRows.length === 0 ? (
                  <p className="text-gray-600 text-sm">No active projects.</p>
                ) : (
                  <div className="space-y-3">
                    {typeRows.map(([label, v]) => (
                      <div key={label}>
                        <div className="flex justify-between text-xs mb-1">
                          <span className={label === 'No type' ? 'text-gray-500 italic' : 'text-gray-300'}>{label}</span>
                          <div className="flex items-center gap-3">
                            <span className="text-gray-600">{v.people} {v.people === 1 ? 'person' : 'people'}</span>
                            <span className="text-gray-400">{v.projects} project{v.projects === 1 ? '' : 's'}</span>
                            <span className="text-gray-600 w-10 text-right">{Math.round((v.projects / activeForType.length) * 100)}%</span>
                          </div>
                        </div>
                        <div className="w-full bg-gray-800 rounded-full h-2">
                          <div className={`h-2 rounded-full ${label === 'No type' ? 'bg-gray-600' : 'bg-sky-500'}`} style={{ width: `${(v.projects / maxType) * 100}%` }} />
                        </div>
                      </div>
                    ))}
                  </div>
                )}
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
                  <p className="text-xs text-gray-500 uppercase tracking-wider">Total Headcount per Project <span className="normal-case text-gray-600 ml-1">— active</span></p>
                  {toggleSwitch(showSupervisorsInChart, () => setShowSupervisorsInChart(v => !v), 'Supervisors')}
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
        })(
          projects.filter(p => (!excludePilots || !p.is_pilot) && (!excludeInternal || !p.is_internal)),
          (() => {
            const excludedIds = new Set(projects.filter(p => (excludePilots && p.is_pilot) || (excludeInternal && p.is_internal)).map(p => p.id))
            return excludedIds.size ? assignments.filter(a => !excludedIds.has(a.project_id)) : assignments
          })(),
        )}

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
                        <option>Lead</option>
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
                                <option>Lead</option>
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
                                const order = ['Lead', 'Supervisor', 'STO', 'Ops Support']
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
                                        <option value="Lead" className="bg-gray-900 text-gray-100">Lead</option>
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
          const openP0 = allMs.filter(m => !m.done && (m.priority === 'P00' || m.priority === 'P0'))
          const openCounts = {
            P00: allMs.filter(m => !m.done && m.priority === 'P00').length,
            P0: allMs.filter(m => !m.done && m.priority === 'P0').length,
            P1: allMs.filter(m => !m.done && m.priority === 'P1').length,
            P2: allMs.filter(m => !m.done && m.priority === 'P2').length,
          }
          return (
            <div className="flex gap-6 items-start">
              {/* Left: per-project milestone checklists */}
              <div className="flex-1 min-w-0 space-y-3">
                <div className="flex flex-wrap justify-end gap-2 items-center">
                  {overviewStatus === 'sent' && overviewPermalink && (
                    <a href={overviewPermalink} target="_blank" rel="noopener noreferrer" className="text-xs text-sky-400 hover:text-sky-300 whitespace-nowrap">View post ↗</a>
                  )}
                  <button
                    onClick={giveTeamOverview}
                    disabled={overviewStatus === 'sending'}
                    className={`text-xs px-3 py-1.5 rounded-lg border transition-colors ${
                      overviewStatus === 'sent' ? 'border-emerald-500/40 text-emerald-400' :
                      overviewStatus === 'error' || overviewStatus === 'empty' ? 'border-red-500/40 text-red-400' :
                      'border-gray-700 text-gray-300 hover:text-white hover:border-gray-500'
                    }`}
                    title="Post a full overview of milestones & production plans to #code_pod"
                  >
                    {overviewStatus === 'sending' ? 'Posting…' :
                     overviewStatus === 'sent' ? '✓ Overview posted' :
                     overviewStatus === 'error' ? '⚠ Failed — retry' :
                     overviewStatus === 'empty' ? 'Nothing to share' :
                     '📣 Give milestones & PP update to the whole team'}
                  </button>
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
                    onClick={() => setClearMsPrompt(true)}
                    className="text-xs px-3 py-1.5 rounded-lg border border-gray-700 text-gray-400 hover:text-red-400 hover:border-red-500/40 transition-colors"
                    title="Clear all milestones and/or production-plan targets from active projects"
                  >
                    🗑 Clear…
                  </button>
                  <div className="pl-1">{toggleSwitch(msSummaryView, () => setMsSummaryView(v => !v), 'Summarized view')}</div>
                  {!msSummaryView && <div className="pl-1">{toggleSwitch(showCompletedMilestones, () => setShowCompletedMilestones(v => !v), 'Show completed')}</div>}
                  {!msSummaryView && <div className="pl-1">{toggleSwitch(hideEmptyMilestoneProjects, () => setHideEmptyMilestoneProjects(v => !v), 'With milestones only')}</div>}
                </div>
                {msSummaryView ? (
                  activeProjects.length === 0 ? (
                    <p className="text-gray-600 text-sm">No active projects.</p>
                  ) : (
                    <div className="border border-gray-800 rounded-xl bg-gray-900/40">
                      <table className="w-full text-sm">
                        <thead>
                          <tr className="text-left text-gray-500 border-b border-gray-800">
                            <th className="py-3 pl-5 pr-3 font-medium text-xs uppercase tracking-wider w-10">#</th>
                            <th className="py-3 pr-3 font-medium text-xs uppercase tracking-wider">Project</th>
                            <th className="py-3 pr-5 font-medium text-xs uppercase tracking-wider text-right">Production plan (this week)</th>
                          </tr>
                        </thead>
                        <tbody>
                          {activeProjects.map((p, i) => (
                            <tr key={p.id} className="border-b border-gray-800/60 last:border-0 hover:bg-gray-800/20">
                              <td className="py-2.5 pl-5 pr-3 text-gray-600 text-xs">{i + 1}</td>
                              <td className="py-2.5 pr-3 text-gray-100">
                                <div className="flex items-center gap-2">{renderProjectEmojiPicker(p)}<span>{p.name}</span></div>
                              </td>
                              <td className="py-2.5 pr-5 text-right">{ppStepper(p)}</td>
                            </tr>
                          ))}
                        </tbody>
                        <tfoot>
                          <tr className="border-t border-gray-800 text-xs">
                            <td className="py-3 pl-5 pr-3" />
                            <td className="py-3 pr-3 text-gray-500">
                              {activeProjects.length} active project{activeProjects.length === 1 ? '' : 's'} · {activeProjects.filter(p => p.weekly_target != null).length} with a plan
                            </td>
                            <td className="py-3 pr-5 text-right text-gray-300 font-medium">
                              Total: {activeProjects.reduce((n, p) => n + (p.weekly_target ?? 0), 0)} tasks
                            </td>
                          </tr>
                        </tfoot>
                      </table>
                    </div>
                  )
                ) : (() => {
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
                  return (
                    <div key={p.id} className="border border-gray-800 rounded-xl p-5 bg-gray-900/40">
                      <div className="flex items-center justify-between mb-4">
                        <h3 className="font-semibold text-gray-100">{p.name}</h3>
                        <div className="flex items-center gap-3">
                          {ms.length > 0 && <span className="text-xs text-gray-500">{pDone}/{ms.length} done</span>}
                          <button
                            onClick={() => handleRequestUpdate(p)}
                            disabled={slackReqStatus[p.id] === 'sending'}
                            className={`text-xs px-3 py-1.5 rounded-lg border transition-colors ${
                              slackReqStatus[p.id] === 'sent' ? 'border-emerald-500/40 text-emerald-400' :
                              slackReqStatus[p.id] === 'error' ? 'border-red-500/40 text-red-400' :
                              'border-gray-700 text-gray-400 hover:text-gray-200'
                            }`}
                          >
                            {slackReqStatus[p.id] === 'sending' ? 'Requesting…' :
                             slackReqStatus[p.id] === 'sent' ? '✓ Requested in Slack' :
                             slackReqStatus[p.id] === 'error' ? '⚠ Failed — retry' :
                             '💬 Request update'}
                          </button>
                          {slackPermalinks[p.id] && (
                            <a href={slackPermalinks[p.id]} target="_blank" rel="noopener noreferrer" className="text-xs text-sky-400 hover:text-sky-300 whitespace-nowrap">View thread ↗</a>
                          )}
                        </div>
                      </div>
                      <div className="flex items-center gap-2 mb-4">
                        <span className="text-xs text-gray-500">Production plan (this week):</span>
                        {ppStepper(p)}
                        {p.weekly_target != null && <span className="text-xs text-gray-600">task{p.weekly_target === 1 ? '' : 's'} planned</span>}
                      </div>
                      {renderMilestoneList(p)}
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
                  <div className="flex gap-3 mt-4 text-xs flex-wrap">
                    {openCounts.P00 > 0 && <span className="text-red-300">⚠️ {openCounts.P00} P00</span>}
                    <span className="text-red-400">{openCounts.P0} P0</span>
                    <span className="text-amber-400">{openCounts.P1} P1</span>
                    <span className="text-sky-400">{openCounts.P2} P2</span>
                    <span className="text-gray-600">open</span>
                  </div>
                </div>

                <div className="bg-gray-900 border border-gray-800 rounded-xl p-5">
                  <div className="flex items-center justify-between mb-3">
                    <p className="text-xs text-gray-500 uppercase tracking-wider">P00 / P0 — needs attention</p>
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

        {/* Consolidated Manager Tab — Assignments + Milestones in one workspace */}
        {tab === 'consolidated' && (() => {
          const today = fmt(new Date())
          const ROLE_ORDER = ['Lead', 'Supervisor', 'STO', 'Ops Support']
          const roleRank = (r: string | null) => { const i = ROLE_ORDER.indexOf(r ?? ''); return i === -1 ? 99 : i }

          // ---- Per-project signals (staffing + milestones + PP) ----
          const info = (p: Project) => {
            const pa = assignments.filter(a => a.project_id === p.id)
            const sto = pa.filter(a => a.assignment_role === 'STO').length
            const ops = pa.filter(a => a.assignment_role === 'Ops Support').length
            const staffing = sto === 0 || (sto >= 1 && ops === 0) ? 'understaffed' : sto >= 2 || ops > 2 ? 'overstaffed' : 'good'
            const ms = milestones.filter(m => m.project_id === p.id)
            const open = ms.filter(m => !m.done)
            const overdue = open.filter(m => m.due_date && m.due_date < today)
            const p0 = open.filter(m => m.priority === 'P00' || m.priority === 'P0')
            const nextDue = [...open].filter(m => m.due_date).sort((a, b) => (a.due_date ?? '').localeCompare(b.due_date ?? ''))[0]
            const tierRole = pa.some(a => a.assignment_role === 'Lead') ? 'Lead'
              : pa.some(a => a.assignment_role === 'Supervisor') ? 'Supervisor'
              : pa.some(a => a.assignment_role === 'STO') ? 'STO' : null
            const owners = tierRole ? pa.filter(a => a.assignment_role === tierRole).map(a => staff.find(s => s.id === a.staff_id)).filter((s): s is Staff => !!s) : []
            const active = p.status === 'active'
            const flags = {
              flagged: p.flagged,
              understaffed: active && staffing === 'understaffed',
              nopp: active && p.weekly_target == null,
              overdue: active && overdue.length > 0,
              p0: active && p0.length > 0,
              noteam: active && pa.length === 0,
            }
            const attention = Object.values(flags).some(Boolean)
            const health = !active ? 'idle' : (flags.flagged || flags.overdue || flags.noteam) ? 'red' : (flags.understaffed || flags.nopp || flags.p0) ? 'amber' : 'green'
            return { pa, staffing, ms, open, overdue, p0, nextDue, tierRole, owners, flags, attention, health }
          }

          // ---- Visible project set ----
          const q = cmSearch.trim().toLowerCase()
          const base = [...projects]
            .filter(p => visibleStatuses[p.status])
            .filter(p => !q || p.name.toLowerCase().includes(q) || (p.customer_codename ?? '').toLowerCase().includes(q))
            .filter(p => !hideEmptyMilestoneProjects || milestones.some(m => m.project_id === p.id))
            .sort((a, b) => statusRank(a.status) - statusRank(b.status) || a.name.localeCompare(b.name))
          const infoMap = new Map(base.map(p => [p.id, info(p)]))
          const countFor = (k: typeof cmFilter) => k === 'all' ? base.length : k === 'attention'
            ? base.filter(p => infoMap.get(p.id)!.attention).length
            : base.filter(p => infoMap.get(p.id)!.flags[k]).length
          const filtered = base.filter(p => {
            const i = infoMap.get(p.id)!
            if (cmFilter === 'all') return true
            if (cmFilter === 'attention') return i.attention
            return i.flags[cmFilter]
          })
          // Supervisor filter (same semantics as the Assignments tab's ungrouped view)
          const shown = groupBySupervisor ? filtered : filtered.filter(p => {
            if (!supervisorFilter) return true
            const hasSup = assignments.some(a => a.project_id === p.id && a.assignment_role === 'Supervisor')
            if (supervisorFilter === '__none__') return !hasSup
            return assignments.some(a => a.project_id === p.id && a.staff_id === supervisorFilter && a.assignment_role === 'Supervisor')
          })

          // ---- Grouping by lead (Lead → Supervisor → STO), same as Assignments ----
          const groups = (() => {
            const map = new Map<string, { key: string; name: string; projects: Project[] }>()
            const ensure = (key: string, name: string) => { if (!map.has(key)) map.set(key, { key, name, projects: [] }); return map.get(key)! }
            shown.forEach(p => {
              const i = infoMap.get(p.id)!
              if (i.owners.length === 0) ensure('__none__', 'No lead / supervisor / STO').projects.push(p)
              else i.owners.forEach(o => ensure(o.id, o.name).projects.push(p))
            })
            let g = [...map.values()]
            if (supervisorFilter) g = g.filter(x => x.key === supervisorFilter)
            return g.sort((a, b) => a.key === '__none__' ? 1 : b.key === '__none__' ? -1 : a.name.localeCompare(b.name))
          })()

          // ---- Pulse (right rail) numbers — active projects ----
          const activeAll = projects.filter(p => p.status === 'active')
          const activeInfo = activeAll.map(p => ({ p, i: info(p) }))
          const ppTotal = activeAll.reduce((n, p) => n + (p.weekly_target ?? 0), 0)
          const ppSet = activeAll.filter(p => p.weekly_target != null).length
          const activeMs = milestones.filter(m => activeAll.some(p => p.id === m.project_id))
          const msDone = activeMs.filter(m => m.done).length
          const msPct = activeMs.length ? Math.round((msDone / activeMs.length) * 100) : 0
          const peopleOnActive = new Set(assignments.filter(a => activeAll.some(p => p.id === a.project_id)).map(a => a.staff_id)).size
          const openP0All = activeMs.filter(m => !m.done && (m.priority === 'P00' || m.priority === 'P0'))
            .sort((a, b) => (a.due_date ?? '9999').localeCompare(b.due_date ?? '9999'))

          const jumpTo = (pid: string) => {
            if (!shown.some(p => p.id === pid)) { setCmFilter('all'); setCmSearch('') }
            setCmView('cards')
            setCmCollapsed(prev => ({ ...prev, [pid]: false }))
            setCmHighlight(pid)
            setTimeout(() => document.getElementById(`cm-proj-${pid}`)?.scrollIntoView({ behavior: 'smooth', block: 'center' }), 60)
            setTimeout(() => setCmHighlight(h => (h === pid ? null : h)), 1800)
          }

          // ---- Small building blocks ----
          const healthDot = (h: string) => (
            <span
              title={h === 'red' ? 'Needs attention now' : h === 'amber' ? 'Worth a look' : h === 'green' ? 'Healthy' : 'Not active'}
              className={`inline-block w-2.5 h-2.5 rounded-full shrink-0 ${h === 'red' ? 'bg-red-500' : h === 'amber' ? 'bg-amber-400' : h === 'green' ? 'bg-emerald-500' : 'bg-gray-600'}`}
            />
          )
          const statusSelect = (p: Project) => (
            <select
              value={p.status}
              onChange={e => updateProjectStatus(p.id, e.target.value)}
              title="Change project status"
              className={`px-2.5 py-1 pr-6 rounded-full text-xs font-medium cursor-pointer appearance-none focus:outline-none focus:ring-1 focus:ring-gray-500 ${
                p.status === 'active' ? 'bg-emerald-500/10 text-emerald-400' :
                p.status === 'starting-soon' ? 'bg-sky-500/10 text-sky-400' :
                p.status === 'paused' ? 'bg-rose-500/10 text-rose-400' :
                'bg-gray-700 text-gray-400'
              }`}
              style={{ backgroundImage: 'none' }}
            >
              {STATUS_OPTIONS.map(o => <option key={o.value} value={o.value} className="bg-gray-900 text-gray-100">{o.label}</option>)}
            </select>
          )
          const staffingBadge = (st: string) => (
            <span className={`px-2 py-0.5 rounded-full text-[11px] font-medium whitespace-nowrap ${
              st === 'understaffed' ? 'bg-amber-500/10 text-amber-400' : st === 'overstaffed' ? 'bg-blue-500/10 text-blue-400' : 'bg-emerald-500/10 text-emerald-400'
            }`}>
              {st === 'understaffed' ? 'Potentially understaffed' : st === 'overstaffed' ? 'Potentially overstaffed' : 'Well staffed'}
            </span>
          )
          const emojiPicker = renderProjectEmojiPicker
          const leadEmojiPicker = (staffId: string) => (
            <span className="relative">
              <button onClick={() => setEmojiPickerStaffId(emojiPickerStaffId === staffId ? null : staffId)} title="Set lead emoji" className="text-lg leading-none hover:scale-110 transition-transform">
                {staff.find(s => s.id === staffId)?.emoji || '👤'}
              </button>
              {emojiPickerStaffId === staffId && (
                <>
                  <div className="fixed inset-0 z-10" onClick={() => setEmojiPickerStaffId(null)} />
                  <div className="absolute left-0 top-8 z-20 bg-gray-900 border border-gray-700 rounded-lg p-2 shadow-xl w-60">
                    <div className="grid grid-cols-6 gap-1 max-h-48 overflow-y-auto">
                      {PROJECT_EMOJIS.map(em => (
                        <button key={em} onClick={() => { updateStaffEmoji(staffId, em); setEmojiPickerStaffId(null) }} className="text-lg hover:bg-gray-800 rounded p-1">{em}</button>
                      ))}
                    </div>
                    <div className="flex gap-1 mt-2 pt-2 border-t border-gray-800">
                      <input
                        className={inputSmClass + ' flex-1'}
                        placeholder="Paste any emoji…"
                        value={customEmoji}
                        onChange={e => setCustomEmoji(e.target.value)}
                        onKeyDown={e => { if (e.key === 'Enter' && customEmoji.trim()) { updateStaffEmoji(staffId, customEmoji.trim()); setCustomEmoji(''); setEmojiPickerStaffId(null) } }}
                      />
                      <button className="text-xs px-2 py-1 rounded bg-gray-700 text-gray-300 hover:text-white transition-colors" onClick={() => { if (customEmoji.trim()) { updateStaffEmoji(staffId, customEmoji.trim()); setCustomEmoji(''); setEmojiPickerStaffId(null) } }}>Set</button>
                    </div>
                    <button onClick={() => { updateStaffEmoji(staffId, null); setEmojiPickerStaffId(null) }} className="w-full text-xs text-gray-500 hover:text-gray-300 mt-2">Reset to default</button>
                  </div>
                </>
              )}
            </span>
          )
          const requestUpdateBtn = (p: Project) => (
            <>
              <button
                onClick={() => handleRequestUpdate(p)}
                disabled={slackReqStatus[p.id] === 'sending'}
                className={`text-xs px-2.5 py-1 rounded-lg border transition-colors whitespace-nowrap ${
                  slackReqStatus[p.id] === 'sent' ? 'border-emerald-500/40 text-emerald-400' :
                  slackReqStatus[p.id] === 'error' ? 'border-red-500/40 text-red-400' :
                  'border-gray-700 text-gray-400 hover:text-gray-200'
                }`}
              >
                {slackReqStatus[p.id] === 'sending' ? 'Requesting…' :
                 slackReqStatus[p.id] === 'sent' ? '✓ Requested' :
                 slackReqStatus[p.id] === 'error' ? '⚠ Retry' :
                 '💬 Request update'}
              </button>
              {slackPermalinks[p.id] && (
                <a href={slackPermalinks[p.id]} target="_blank" rel="noopener noreferrer" className="text-xs text-sky-400 hover:text-sky-300 whitespace-nowrap">Thread ↗</a>
              )}
            </>
          )
          const dropProps = (pid: string) => ({
            onDragOver: (e: React.DragEvent) => { e.preventDefault(); setDragOverProjectId(pid) },
            onDragLeave: () => setDragOverProjectId(null),
            onDrop: (e: React.DragEvent) => {
              e.preventDefault(); setDragOverProjectId(null)
              const type = e.dataTransfer.getData('type')
              const id = e.dataTransfer.getData('id')
              if (type === 'staff' && id) {
                if (!assignments.some(a => a.project_id === pid && a.staff_id === id)) setPendingDrop({ staffId: id, projectId: pid })
                setDraggedStaffId(null)
              } else if (type === 'assignment' && id) {
                const a = assignments.find(x => x.id === id)
                if (a && a.project_id !== pid && !assignments.some(x => x.project_id === pid && x.staff_id === a.staff_id)) {
                  setPendingMove({ assignmentId: id, projectId: pid })
                  setDropRole(a.assignment_role ?? '')
                }
                setDraggedAssignmentId(null)
              }
            },
          })
          const assignmentRow = (a: Assignment, groupOwnerId?: string) => {
            const member = staff.find(s => s.id === a.staff_id)
            const isGroupOwner = !!groupOwnerId && a.staff_id === groupOwnerId
            const nonSupCount = new Set(assignments.filter(x => x.staff_id === a.staff_id && x.assignment_role !== 'Supervisor' && x.assignment_role !== 'Lead').map(x => x.project_id)).size
            const overAllocated = a.assignment_role !== 'Supervisor' && a.assignment_role !== 'Lead' && nonSupCount >= 2
            return (
              <div
                key={a.id}
                draggable
                onDragStart={e => { e.dataTransfer.setData('type', 'assignment'); e.dataTransfer.setData('id', a.id); setDraggedAssignmentId(a.id); setDraggedStaffId(null) }}
                onDragEnd={() => setDraggedAssignmentId(null)}
                className={`flex items-center justify-between gap-2 rounded-lg px-3 py-2 text-sm cursor-grab active:cursor-grabbing transition-opacity ${isGroupOwner ? 'bg-gray-800/20 border border-gray-700/70' : 'bg-gray-800/50'} ${draggedAssignmentId === a.id ? 'opacity-40' : ''}`}
              >
                <div className="flex items-center gap-2 flex-wrap min-w-0">
                  <span className="font-medium text-gray-200 truncate">{member?.name ?? 'Unknown'}</span>
                  {isGroupOwner && <span className="text-[10px] uppercase tracking-wider text-gray-500">group lead</span>}
                  <select
                    value={a.assignment_role ?? ''}
                    onChange={e => updateAssignmentRole(a.id, e.target.value)}
                    onClick={e => e.stopPropagation()}
                    title="Change role"
                    className={`text-xs font-medium px-2 py-0.5 pr-5 rounded-full cursor-pointer appearance-none focus:outline-none focus:ring-1 focus:ring-gray-500 ${a.assignment_role ? roleColor(a.assignment_role) : 'bg-gray-700/60 text-gray-400'}`}
                    style={{ backgroundImage: 'none' }}
                  >
                    <option value="" className="bg-gray-900 text-gray-100">No role</option>
                    {ROLE_ORDER.map(r => <option key={r} value={r} className="bg-gray-900 text-gray-100">{r}</option>)}
                  </select>
                  {member?.ooo && (
                    <span className="flex items-center gap-1 text-xs font-medium px-2 py-0.5 rounded-full bg-amber-500/10 text-amber-400" title={member.ooo_return_date ? `OOO — back ${member.ooo_return_date}` : 'Out of office'}>
                      ⚠️ OOO{member.ooo_return_date ? ` · back ${member.ooo_return_date}` : ''}
                      <button onClick={() => clearStaffOOO(member.id)} title="Clear OOO" className="ml-0.5 text-amber-400/70 hover:text-amber-300 leading-none">×</button>
                    </span>
                  )}
                  {overAllocated && (
                    <span className="text-xs font-medium px-2 py-0.5 rounded-full bg-red-500/10 text-red-400" title={`Assigned to ${nonSupCount} projects in a non-Supervisor role`}>⚠️ In {nonSupCount} projects</span>
                  )}
                  {member?.position && <span className="text-xs text-gray-500">{member.position}</span>}
                </div>
                <button className={btnDanger} onClick={() => removeAssignment(a.id)} title="Remove from project">×</button>
              </div>
            )
          }
          const quickAddRow = (p: Project, pa: Assignment[]) => (
            <div className="flex flex-wrap gap-2 mb-2 p-2 rounded-lg bg-gray-800/40 border border-gray-700">
              <select className={selectSmClass + ' flex-1 min-w-[140px]'} value={quickAdd.staff_id} onChange={e => setQuickAdd({ ...quickAdd, staff_id: e.target.value })}>
                <option value="">Select staff…</option>
                {[...staff].filter(s => !pa.some(a => a.staff_id === s.id)).sort((a, b) => a.name.localeCompare(b.name))
                  .map(s => <option key={s.id} value={s.id}>{s.name}{s.position ? ` — ${s.position}` : ''}</option>)}
              </select>
              <select className={selectSmClass + ' w-28'} value={quickAdd.assignment_role} onChange={e => setQuickAdd({ ...quickAdd, assignment_role: e.target.value })}>
                <option value="">Role</option>
                {ROLE_ORDER.map(r => <option key={r}>{r}</option>)}
              </select>
              <button
                className="text-xs px-2 py-1 rounded bg-gray-700 text-gray-300 hover:text-white transition-colors whitespace-nowrap"
                onClick={async () => {
                  if (!quickAdd.staff_id) return
                  const { data } = await supabase.from('assignments').insert([{ project_id: p.id, staff_id: quickAdd.staff_id, assignment_role: quickAdd.assignment_role || null }]).select().single()
                  if (data) setAssignments(prev => [...prev, data])
                  setQuickAdd({ staff_id: '', assignment_role: '' })
                  setAddingToProjectId(null)
                }}
              >+ Add</button>
              <button className={btnCancel} onClick={() => { setAddingToProjectId(null); setQuickAdd({ staff_id: '', assignment_role: '' }) }}>Cancel</button>
            </div>
          )

          // ---- The project card: header · PP & Slack strip · Team | Milestones ----
          const card = (p: Project, groupOwnerId?: string) => {
            const i = infoMap.get(p.id) ?? info(p)
            const coOwners = groupOwnerId ? i.owners.filter(o => o.id !== groupOwnerId) : i.owners
            const collapsed = !!cmCollapsed[p.id]
            const isOver = dragOverProjectId === p.id
            const done = i.ms.length - i.open.length
            const pct = i.ms.length ? Math.round((done / i.ms.length) * 100) : 0
            return (
              <div
                key={p.id}
                id={`cm-proj-${p.id}`}
                {...dropProps(p.id)}
                className={`border rounded-xl p-4 transition-all ${
                  isOver ? 'border-[#193a29] bg-[#193a29]/10 scale-[1.005]' :
                  p.flagged ? 'border-red-500/40 bg-red-500/5' : 'border-gray-800 bg-gray-900/40'
                } ${cmHighlight === p.id ? 'ring-2 ring-sky-500/70' : ''}`}
              >
                <div className="flex items-start justify-between gap-3">
                  <div className="flex items-center gap-2 flex-wrap min-w-0">
                    {healthDot(i.health)}
                    {emojiPicker(p)}
                    {renderProjectName(p, 'font-semibold text-gray-100')}
                    {p.customer_codename && <span className="text-xs text-gray-500">{p.customer_codename}</span>}
                    <button
                      onClick={() => toggleProjectFlag(p.id, !p.flagged)}
                      title={p.flagged ? 'Flagged: needs more support — click to clear' : 'Flag: needs more support'}
                      className={`text-sm leading-none transition-colors ${p.flagged ? '' : 'opacity-40 hover:opacity-100 grayscale'}`}
                    >🚩</button>
                    {p.flagged && <span className="px-2 py-0.5 rounded-full text-[11px] font-medium bg-red-500/10 text-red-400">Needs more support</span>}
                    {p.status === 'active' && staffingBadge(i.staffing)}
                    {i.overdue.length > 0 && <span className="px-2 py-0.5 rounded-full text-[11px] font-medium bg-red-500/10 text-red-400">⏰ {i.overdue.length} overdue</span>}
                    {i.p0.length > 0 && <span className="px-2 py-0.5 rounded-full text-[11px] font-medium bg-red-500/10 text-red-300">🔥 {i.p0.length} open P0</span>}
                  </div>
                  <div className="flex items-center gap-2 shrink-0">
                    {statusSelect(p)}
                    <button
                      onClick={() => setCmCollapsed(prev => ({ ...prev, [p.id]: !collapsed }))}
                      title={collapsed ? 'Expand' : 'Collapse'}
                      className="w-6 h-6 flex items-center justify-center rounded-md text-gray-500 hover:text-gray-200 hover:bg-gray-800 transition-colors text-xs"
                    >{collapsed ? '▸' : '▾'}</button>
                  </div>
                </div>

                <div className="flex flex-wrap items-center gap-x-4 gap-y-2 mt-3">
                  <div className="flex items-center gap-2">
                    <span className="text-xs text-gray-500">PP this week</span>
                    {ppStepper(p)}
                  </div>
                  <div className="flex items-center gap-2">{requestUpdateBtn(p)}</div>
                  {coOwners.length > 0 && (
                    <span className="text-xs text-gray-500">
                      {groupOwnerId ? 'Also led by' : i.tierRole === 'Lead' ? 'Lead' : `${i.tierRole} (acting lead)`}: <span className="text-gray-300">{coOwners.map(o => o.name).join(', ')}</span>
                    </span>
                  )}
                  {i.ms.length > 0 && (
                    <div className="flex items-center gap-2 ml-auto" title={`${done} of ${i.ms.length} milestones done`}>
                      <div className="w-24 bg-gray-800 rounded-full h-1.5"><div className="h-1.5 rounded-full bg-emerald-600" style={{ width: `${pct}%` }} /></div>
                      <span className="text-xs text-gray-500">{done}/{i.ms.length}</span>
                    </div>
                  )}
                </div>

                {collapsed ? (
                  <p className="text-xs text-gray-500 mt-3">
                    👥 {i.pa.length} {i.pa.length === 1 ? 'person' : 'people'} · ✅ {i.open.length} open milestone{i.open.length === 1 ? '' : 's'}
                    {i.nextDue ? <> · next due <span className={i.nextDue.due_date! < today ? 'text-red-400' : 'text-gray-300'}>{i.nextDue.due_date}</span></> : null}
                  </p>
                ) : (
                  <div className="grid grid-cols-1 xl:grid-cols-2 gap-4 mt-4">
                    {/* Team */}
                    <div className="min-w-0">
                      <div className="flex items-center gap-2 mb-2">
                        <p className="text-[11px] text-gray-500 uppercase tracking-wider">Team · {i.pa.length}</p>
                        <button
                          onClick={() => { setAddingToProjectId(addingToProjectId === p.id ? null : p.id); setQuickAdd({ staff_id: '', assignment_role: '' }) }}
                          title="Add a person to this project"
                          className="w-5 h-5 flex items-center justify-center rounded-full border border-gray-600 text-gray-400 hover:text-white hover:border-gray-400 transition-colors text-sm leading-none"
                        >+</button>
                      </div>
                      {addingToProjectId === p.id && quickAddRow(p, i.pa)}
                      {i.pa.length === 0 ? (
                        <p className={`text-xs border border-dashed rounded-lg px-3 py-3 text-center ${isOver ? 'border-[#193a29] text-gray-300' : 'border-gray-800 text-gray-600'}`}>
                          {isOver ? 'Drop to assign' : 'No staff yet — drag someone here from the pools'}
                        </p>
                      ) : (
                        <div className="space-y-1.5">{[...i.pa].sort((a, b) => roleRank(a.assignment_role) - roleRank(b.assignment_role)).map(a => assignmentRow(a, groupOwnerId))}</div>
                      )}
                    </div>
                    {/* Milestones */}
                    <div className="min-w-0 xl:border-l xl:border-gray-800 xl:pl-4">
                      <p className="text-[11px] text-gray-500 uppercase tracking-wider mb-2">Milestones · {done}/{i.ms.length}</p>
                      {renderMilestoneList(p)}
                    </div>
                  </div>
                )}
              </div>
            )
          }

          // ---- Summary table (one row per project) ----
          const tableRow = (p: Project, idx: number) => {
            const i = infoMap.get(p.id) ?? info(p)
            const done = i.ms.length - i.open.length
            return (
              <tr key={p.id} {...dropProps(p.id)} className={`border-b border-gray-800/60 last:border-0 transition-colors ${dragOverProjectId === p.id ? 'bg-[#193a29]/20' : 'hover:bg-gray-800/20'}`}>
                <td className="py-2.5 pl-4 pr-2 text-gray-600 text-xs">{idx + 1}</td>
                <td className="py-2.5 pr-3">
                  <div className="flex items-center gap-2">
                    {healthDot(i.health)}
                    <button onClick={() => jumpTo(p.id)} className="text-gray-100 hover:underline decoration-dotted underline-offset-4 text-left" title="Open in card view">{p.emoji || '📁'} {p.name}</button>
                    {p.flagged && <span title="Needs more support">🚩</span>}
                  </div>
                </td>
                <td className="py-2.5 pr-3">{statusSelect(p)}</td>
                <td className="py-2.5 pr-3 text-xs text-gray-300">{i.owners.map(o => o.name).join(', ') || <span className="text-gray-600">—</span>}</td>
                <td className="py-2.5 pr-3 text-xs text-gray-400" title={i.pa.map(a => staff.find(s => s.id === a.staff_id)?.name).filter(Boolean).join(', ')}>{i.pa.length}</td>
                <td className="py-2.5 pr-3">{p.status === 'active' ? staffingBadge(i.staffing) : <span className="text-xs text-gray-600">—</span>}</td>
                <td className="py-2.5 pr-3 text-xs">
                  <span className="text-gray-300">{done}/{i.ms.length}</span>
                  {i.overdue.length > 0 && <span className="ml-2 text-red-400">⏰{i.overdue.length}</span>}
                  {i.p0.length > 0 && <span className="ml-2 text-red-300">🔥{i.p0.length}</span>}
                </td>
                <td className={`py-2.5 pr-3 text-xs ${i.nextDue && i.nextDue.due_date! < today ? 'text-red-400' : 'text-gray-400'}`}>{i.nextDue?.due_date ?? '—'}</td>
                <td className="py-2.5 pr-4 text-right">{ppStepper(p)}</td>
              </tr>
            )
          }
          const table = (rowsFor: { key: string; title: React.ReactNode; projects: Project[] }[]) => (
            <div className="border border-gray-800 rounded-xl bg-gray-900/40 overflow-x-auto">
              <table className="w-full text-sm">
                <thead>
                  <tr className="text-left text-gray-500 border-b border-gray-800 text-xs uppercase tracking-wider">
                    <th className="py-3 pl-4 pr-2 font-medium w-8">#</th>
                    <th className="py-3 pr-3 font-medium">Project</th>
                    <th className="py-3 pr-3 font-medium">Status</th>
                    <th className="py-3 pr-3 font-medium">Lead</th>
                    <th className="py-3 pr-3 font-medium">Team</th>
                    <th className="py-3 pr-3 font-medium">Staffing</th>
                    <th className="py-3 pr-3 font-medium">Milestones</th>
                    <th className="py-3 pr-3 font-medium">Next due</th>
                    <th className="py-3 pr-4 font-medium text-right">PP this week</th>
                  </tr>
                </thead>
                <tbody>
                  {rowsFor.map(g => (
                    <Fragment key={g.key}>
                      {g.title && (
                        <tr className="border-y border-gray-800" style={{ background: `linear-gradient(90deg, ${accentOf(g.key)}1a, transparent 70%)` }}>
                          <td colSpan={9} className="py-2.5 pl-3 pr-4" style={{ boxShadow: `inset 4px 0 0 ${accentOf(g.key)}` }}>{g.title}</td>
                        </tr>
                      )}
                      {g.projects.map((p, idx) => tableRow(p, idx))}
                    </Fragment>
                  ))}
                </tbody>
                <tfoot>
                  <tr className="border-t border-gray-800 text-xs">
                    <td />
                    <td className="py-3 pr-3 text-gray-500" colSpan={3}>{shown.length} project{shown.length === 1 ? '' : 's'} shown</td>
                    <td className="py-3 pr-3 text-gray-400">{shown.reduce((n, p) => n + (infoMap.get(p.id)?.pa.length ?? 0), 0)}</td>
                    <td />
                    <td className="py-3 pr-3 text-gray-400">{shown.reduce((n, p) => n + (infoMap.get(p.id)?.open.length ?? 0), 0)} open</td>
                    <td />
                    <td className="py-3 pr-4 text-right text-gray-300 font-medium">Total: {shown.reduce((n, p) => n + (p.weekly_target ?? 0), 0)} tasks</td>
                  </tr>
                </tfoot>
              </table>
            </div>
          )

          // ---- Lead groups: accent colour, role label and roll-up stats ----
          type CmGroup = { key: string; name: string; projects: Project[] }
          const ACCENTS = ['#3b82f6', '#8b5cf6', '#10b981', '#f59e0b', '#ec4899', '#06b6d4', '#ef4444', '#84cc16', '#f97316', '#14b8a6']
          const accentOf = (key: string) => {
            if (key === '__none__') return '#f59e0b'
            let h = 0
            for (const c of key) h = (h * 31 + c.charCodeAt(0)) >>> 0
            return ACCENTS[h % ACCENTS.length]
          }
          const groupStats = (g: CmGroup) => {
            const infos = g.projects.map(p => infoMap.get(p.id) ?? info(p))
            const people = new Set(infos.flatMap(x => x.pa.map(a => a.staff_id)))
            if (g.key !== '__none__') people.delete(g.key)
            const leadCount = infos.filter(x => x.tierRole === 'Lead').length
            const actingRoles = [...new Set(infos.filter(x => x.tierRole && x.tierRole !== 'Lead').map(x => x.tierRole as string))]
            const roleLabel = g.key === '__none__' ? null
              : leadCount === infos.length ? 'Lead'
              : leadCount === 0 ? `${actingRoles.join(' / ')} · acting lead`
              : `Lead · acting lead on ${infos.length - leadCount}`
            return {
              people: people.size,
              pp: g.projects.reduce((n, p) => n + (p.weekly_target ?? 0), 0),
              open: infos.reduce((n, x) => n + x.open.length, 0),
              attention: infos.filter(x => x.attention).length,
              health: (['red', 'amber', 'green', 'idle'] as const).map(h => ({ h, n: infos.filter(x => x.health === h).length })),
              roleLabel,
              actingOnly: leadCount === 0,
            }
          }
          const statPill = (label: React.ReactNode, title?: string) => (
            <span title={title} className="text-[11px] px-2 py-0.5 rounded-full bg-gray-800/80 text-gray-300 whitespace-nowrap">{label}</span>
          )
          const groupHeader = (g: CmGroup, opts: { collapsible: boolean }) => {
            const st = groupStats(g)
            const accent = accentOf(g.key)
            const collapsed = !!cmGroupCollapsed[g.key]
            const lead = staff.find(x => x.id === g.key)
            return (
              <div className="flex flex-wrap items-center justify-between gap-3">
                <div className="flex items-center gap-3 min-w-0">
                  {opts.collapsible && (
                    <button
                      onClick={() => setCmGroupCollapsed(prev => ({ ...prev, [g.key]: !collapsed }))}
                      title={collapsed ? 'Expand group' : 'Collapse group'}
                      className="w-6 h-6 flex items-center justify-center rounded-md text-gray-500 hover:text-gray-200 hover:bg-gray-800 transition-colors text-xs"
                    >{collapsed ? '▸' : '▾'}</button>
                  )}
                  {g.key === '__none__' ? (
                    <span className="w-9 h-9 rounded-full flex items-center justify-center text-lg" style={{ backgroundColor: `${accent}22` }}>⚠️</span>
                  ) : (
                    <span className="w-9 h-9 rounded-full flex items-center justify-center" style={{ backgroundColor: `${accent}22`, boxShadow: `inset 0 0 0 1px ${accent}55` }}>{leadEmojiPicker(g.key)}</span>
                  )}
                  <div className="min-w-0">
                    <div className="flex items-center gap-2 flex-wrap">
                      <span className="text-base font-semibold text-gray-100">{g.key === '__none__' ? 'No lead, supervisor or STO' : g.name}</span>
                      {st.roleLabel && (
                        <span className="text-[11px] font-medium px-2 py-0.5 rounded-full" style={{ backgroundColor: `${accent}1f`, color: accent }}>{st.roleLabel}</span>
                      )}
                    </div>
                    <p className="text-xs text-gray-500 mt-0.5">
                      {g.key === '__none__'
                        ? 'These projects have nobody to group under — assign a Lead (or Supervisor / STO).'
                        : <>{lead?.position ?? 'No position'}{st.actingOnly ? ' · no Lead assigned on these projects' : ''}</>}
                    </p>
                  </div>
                </div>
                <div className="flex items-center gap-1.5 flex-wrap">
                  {statPill(`${g.projects.length} project${g.projects.length === 1 ? '' : 's'}`)}
                  {statPill(`👥 ${st.people}`, 'People on these projects (excluding the lead)')}
                  {statPill(`PP ${st.pp}`, 'Production plan this week, summed')}
                  {statPill(`✅ ${st.open} open`, 'Open milestones')}
                  {st.attention > 0 && <span className="text-[11px] px-2 py-0.5 rounded-full bg-red-500/10 text-red-400 whitespace-nowrap">⚠ {st.attention} need attention</span>}
                  <div className="flex h-1.5 w-16 rounded-full overflow-hidden bg-gray-800 ml-1" title="Health mix">
                    {st.health.map(({ h, n }) => n > 0 ? (
                      <div key={h} className={h === 'red' ? 'bg-red-500' : h === 'amber' ? 'bg-amber-400' : h === 'green' ? 'bg-emerald-500' : 'bg-gray-600'} style={{ width: `${(n / g.projects.length) * 100}%` }} />
                    ) : null)}
                  </div>
                </div>
              </div>
            )
          }
          const groupSection = (g: CmGroup) => {
            const accent = accentOf(g.key)
            const collapsed = !!cmGroupCollapsed[g.key]
            const none = g.key === '__none__'
            return (
              <section
                key={g.key}
                className={`rounded-2xl border overflow-hidden ${none ? 'border-dashed border-amber-500/40 bg-amber-500/[0.03]' : 'border-gray-800 bg-gray-900/20'}`}
                style={{ borderLeft: `4px solid ${accent}` }}
              >
                <div className={`px-4 py-3 ${collapsed ? '' : 'border-b border-gray-800/80'}`} style={{ background: `linear-gradient(90deg, ${accent}14, transparent 60%)` }}>
                  {groupHeader(g, { collapsible: true })}
                </div>
                {!collapsed && (
                  <div className="p-3 space-y-3">{g.projects.map(p => card(p, none ? undefined : g.key))}</div>
                )}
              </section>
            )
          }

          const chip = (k: typeof cmFilter, label: string) => {
            const n = countFor(k)
            const on = cmFilter === k
            return (
              <button
                key={k}
                onClick={() => setCmFilter(k)}
                className={`text-xs px-2.5 py-1 rounded-full border transition-colors whitespace-nowrap ${
                  on ? 'text-white border-transparent' : n === 0 && k !== 'all' ? 'border-gray-800 text-gray-600' : 'border-gray-700 text-gray-400 hover:text-gray-200'
                }`}
                style={on ? { backgroundColor: '#193a29' } : {}}
              >{label} <span className={on ? 'text-white/70' : 'text-gray-500'}>{n}</span></button>
            )
          }
          const menuBtn = (key: 'export' | 'clear' | 'slack', label: string, items: { label: React.ReactNode; onClick: () => void; danger?: boolean; hidden?: boolean }[]) => (
            <div className="relative">
              <button
                onClick={() => setCmMenu(cmMenu === key ? null : key)}
                className={`text-xs px-3 py-1.5 rounded-lg border transition-colors whitespace-nowrap ${cmMenu === key ? 'border-gray-500 text-gray-100' : 'border-gray-700 text-gray-300 hover:text-white hover:border-gray-500'}`}
              >{label} ▾</button>
              {cmMenu === key && (
                <>
                  <div className="fixed inset-0 z-10" onClick={() => setCmMenu(null)} />
                  <div className="absolute right-0 top-9 z-20 bg-gray-900 border border-gray-700 rounded-lg shadow-xl py-1 min-w-[260px]">
                    {items.filter(it => !it.hidden).map((it, n) => (
                      <button
                        key={n}
                        onClick={() => { setCmMenu(null); it.onClick() }}
                        className={`w-full text-left text-xs px-3 py-2 transition-colors ${it.danger ? 'text-red-400 hover:bg-red-500/10' : 'text-gray-300 hover:bg-gray-800 hover:text-white'}`}
                      >{it.label}</button>
                    ))}
                  </div>
                </>
              )}
            </div>
          )

          return (
            <div className="space-y-4">
              {/* ===== Command bar ===== */}
              <div className="bg-gray-900 border border-gray-800 rounded-xl p-4 space-y-3">
                <div className="flex flex-wrap items-center gap-2">
                  <div className="relative flex-1 min-w-[220px]">
                    <input className={inputClass + ' w-full pl-9'} placeholder="Search projects or customers…" value={cmSearch} onChange={e => setCmSearch(e.target.value)} />
                    <span className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-500 text-sm pointer-events-none">🔍</span>
                    {cmSearch && <button onClick={() => setCmSearch('')} className="absolute right-3 top-1/2 -translate-y-1/2 text-gray-500 hover:text-gray-300 text-sm">✕</button>}
                  </div>
                  {statusFilterDropdown()}
                  <select
                    value={supervisorFilter}
                    onChange={e => setSupervisorFilter(e.target.value)}
                    title="Filter projects by supervisor"
                    className={`text-xs px-3 py-1.5 rounded-lg border transition-colors cursor-pointer ${supervisorFilter ? 'text-emerald-400' : 'border-gray-700 text-gray-400 hover:text-gray-200'}`}
                    style={supervisorFilter ? { borderColor: '#193a29' } : {}}
                  >
                    <option value="">All supervisors</option>
                    <option value="__none__">No supervisor</option>
                    {staff.filter(s => assignments.some(a => a.staff_id === s.id && a.assignment_role === 'Supervisor')).sort((a, b) => a.name.localeCompare(b.name))
                      .map(s => <option key={s.id} value={s.id}>{s.name}</option>)}
                  </select>
                  {toggleSwitch(groupBySupervisor, () => setGroupBySupervisor(v => !v), 'Group by lead')}
                  <div className="flex items-center gap-1 bg-gray-800/60 rounded-lg p-0.5">
                    {(['cards', 'table'] as const).map(v => (
                      <button key={v} onClick={() => setCmView(v)} className={`text-xs px-3 py-1 rounded-md transition-colors ${cmView === v ? 'text-white' : 'text-gray-400 hover:text-gray-200'}`} style={cmView === v ? { backgroundColor: '#193a29' } : {}}>
                        {v === 'cards' ? '▦ Cards' : '☰ Table'}
                      </button>
                    ))}
                  </div>
                  {cmView === 'cards' && (
                    <>
                      <button onClick={() => { setCmCollapsed({}); setCmGroupCollapsed({}) }} className="text-xs px-2.5 py-1.5 rounded-lg border border-gray-700 text-gray-400 hover:text-gray-200 transition-colors">Expand all</button>
                      <button onClick={() => groupBySupervisor ? setCmGroupCollapsed(Object.fromEntries(groups.map(g => [g.key, true]))) : setCmCollapsed(Object.fromEntries(shown.map(p => [p.id, true])))} className="text-xs px-2.5 py-1.5 rounded-lg border border-gray-700 text-gray-400 hover:text-gray-200 transition-colors">Collapse all</button>
                    </>
                  )}
                </div>

                <div className="flex flex-wrap items-center gap-2">
                  <button
                    onClick={() => setShowAddProjectInline(v => !v)}
                    className={`text-xs px-3 py-1.5 rounded-lg border transition-colors ${showAddProjectInline ? 'text-emerald-400' : 'border-gray-700 text-gray-300 hover:text-white hover:border-gray-500'}`}
                    style={showAddProjectInline ? { borderColor: '#193a29' } : {}}
                  >+ New project</button>
                  <button
                    onClick={() => setCmShowAssign(v => !v)}
                    className={`text-xs px-3 py-1.5 rounded-lg border transition-colors ${cmShowAssign ? 'text-emerald-400' : 'border-gray-700 text-gray-300 hover:text-white hover:border-gray-500'}`}
                    style={cmShowAssign ? { borderColor: '#193a29' } : {}}
                  >⚡ Quick assign</button>
                  {menuBtn('slack', '💬 Slack', [
                    { label: leadStatus === 'sending' ? 'Posting…' : leadStatus === 'sent' ? '✓ Sent to #code_leads_plus — send again' : leadStatus === 'error' ? '⚠ Confirm with leads failed — retry' : leadStatus === 'empty' ? 'Confirm with leads (no leads on active projects)' : '✅ Confirm assignments with leads', onClick: handleConfirmLeads },
                    { label: overviewStatus === 'sending' ? 'Posting…' : overviewStatus === 'sent' ? '✓ Overview posted — post again' : overviewStatus === 'error' ? '⚠ Overview failed — retry' : overviewStatus === 'empty' ? 'Milestones & PP overview (nothing to share)' : '📣 Give milestones & PP update to the whole team', onClick: giveTeamOverview },
                  ])}
                  {leadStatus === 'sent' && leadPermalink && <a href={leadPermalink} target="_blank" rel="noopener noreferrer" className="text-xs text-sky-400 hover:text-sky-300 whitespace-nowrap">Leads post ↗</a>}
                  {overviewStatus === 'sent' && overviewPermalink && <a href={overviewPermalink} target="_blank" rel="noopener noreferrer" className="text-xs text-sky-400 hover:text-sky-300 whitespace-nowrap">Overview post ↗</a>}
                  {menuBtn('export', '⬇ Export', [
                    { label: 'Assignments — HTML', onClick: exportAssignmentsHTML },
                    { label: 'Assignments — DOCX', onClick: exportAssignmentsDOCX },
                    { label: 'Milestones — HTML', onClick: exportMilestonesHTML },
                    { label: 'Milestones — PDF', onClick: exportMilestonesPDF },
                  ])}
                  {menuBtn('clear', '🗑 Clear', [
                    { label: `Clear staffing from inactive projects (${inactiveAssignmentIds().length})`, onClick: () => setConfirmClearInactive(true), danger: true, hidden: inactiveAssignmentIds().length === 0 },
                    { label: `Clear all assignments (${assignments.length})`, onClick: () => setConfirmClearAll(true), danger: true, hidden: assignments.length === 0 },
                    { label: 'Clear milestones and/or PP targets…', onClick: () => setClearMsPrompt(true), danger: true },
                  ])}
                  <div className="flex items-center gap-4 ml-auto">
                    {toggleSwitch(showCompletedMilestones, () => setShowCompletedMilestones(v => !v), 'Show completed')}
                    {toggleSwitch(hideEmptyMilestoneProjects, () => setHideEmptyMilestoneProjects(v => !v), 'With milestones only')}
                  </div>
                </div>

                {cmShowAssign && (
                  <div className="flex flex-wrap items-center gap-2 pt-3 border-t border-gray-800">
                    <select className={selectClass} value={newAssignment.project_id} onChange={e => setNewAssignment({ ...newAssignment, project_id: e.target.value })}>
                      <option value="">Select project *</option>
                      {[...projects].filter(p => visibleStatuses[p.status]).sort((a, b) => statusRank(a.status) - statusRank(b.status) || a.name.localeCompare(b.name)).map(p => <option key={p.id} value={p.id}>{p.name}</option>)}
                    </select>
                    <select className={selectClass} value={newAssignment.staff_id} onChange={e => setNewAssignment({ ...newAssignment, staff_id: e.target.value })}>
                      <option value="">Select staff *</option>
                      {staff.filter(s => !hideAssignedInDropdown || !assignments.some(a => a.staff_id === s.id)).sort((a, b) => a.name.localeCompare(b.name)).map(s => <option key={s.id} value={s.id}>{s.name}</option>)}
                    </select>
                    <select className={selectClass} value={newAssignment.assignment_role} onChange={e => setNewAssignment({ ...newAssignment, assignment_role: e.target.value })}>
                      <option value="">Role</option>
                      {ROLE_ORDER.map(r => <option key={r}>{r}</option>)}
                    </select>
                    <button className={btnPrimary} style={{ backgroundColor: '#193a29' }} onClick={addAssignment}>Assign</button>
                    {toggleSwitch(hideAssignedInDropdown, () => setHideAssignedInDropdown(!hideAssignedInDropdown), 'Unassigned only')}
                  </div>
                )}

                {showAddProjectInline && (
                  <div className="flex flex-wrap items-center gap-2 pt-3 border-t border-gray-800">
                    <input className={inputClass} placeholder="Project name *" value={newProject.name} onChange={e => setNewProject({ ...newProject, name: e.target.value })} onKeyDown={e => e.key === 'Enter' && addProject()} />
                    <input className={inputClass} placeholder="Customer Codename" value={newProject.customer_codename} onChange={e => setNewProject({ ...newProject, customer_codename: e.target.value })} />
                    <select className={selectClass} value={newProject.status} onChange={e => setNewProject({ ...newProject, status: e.target.value })}>
                      {STATUS_OPTIONS.map(o => <option key={o.value} value={o.value}>{o.label}</option>)}
                    </select>
                    <input className={inputClass} type="number" placeholder="Expected Duration (Weeks)" value={newProject.duration_weeks} onChange={e => setNewProject({ ...newProject, duration_weeks: e.target.value })} style={{ width: 210 }} />
                    <button className={btnPrimary} style={{ backgroundColor: '#193a29' }} onClick={addProject}>Add</button>
                  </div>
                )}

                <div className="flex flex-wrap items-center gap-1.5 pt-1">
                  {chip('all', 'All')}
                  {chip('attention', '⚠ Needs attention')}
                  {chip('flagged', '🚩 Flagged')}
                  {chip('understaffed', '🟠 Understaffed')}
                  {chip('nopp', '📉 No PP')}
                  {chip('overdue', '⏰ Overdue')}
                  {chip('p0', '🔥 Open P0')}
                  {chip('noteam', '👻 No team')}
                </div>
              </div>

              {/* ===== Workspace: pools | projects | pulse ===== */}
              <div className="flex gap-5 items-start">
                {renderStaffPools('w-48')}

                <div className="flex-1 min-w-0">
                  {shown.length === 0 ? (
                    <div className="border border-dashed border-gray-800 rounded-xl p-10 text-center">
                      <p className="text-gray-400 text-sm">{cmFilter === 'all' ? 'No projects match.' : 'Nothing here — nice! 🎉'}</p>
                      {(cmFilter !== 'all' || cmSearch) && <button onClick={() => { setCmFilter('all'); setCmSearch('') }} className="text-xs text-sky-400 hover:text-sky-300 mt-2">Clear filters</button>}
                    </div>
                  ) : cmView === 'table' ? (
                    table(groupBySupervisor ? groups.map(g => ({ key: g.key, title: groupHeader(g, { collapsible: false }), projects: g.projects })) : [{ key: 'all', title: null, projects: shown }])
                  ) : groupBySupervisor ? (
                    <div className="space-y-5">
                      {groups.length === 0 && <p className="text-gray-600 text-sm">No leads to show.</p>}
                      {groups.map(groupSection)}
                    </div>
                  ) : (
                    <div className="space-y-3">{shown.map(p => card(p))}</div>
                  )}
                </div>

                {/* Pulse rail */}
                <div className="w-72 shrink-0 sticky top-6 flex flex-col gap-4">
                  <div className="grid grid-cols-2 gap-3">
                    {[
                      { label: 'Active projects', value: activeAll.length, sub: `${peopleOnActive} people on them` },
                      { label: 'PP this week', value: ppTotal, sub: `${ppSet}/${activeAll.length} with a plan` },
                      { label: 'Milestones done', value: `${msPct}%`, sub: `${msDone} of ${activeMs.length}` },
                      { label: 'Open P0', value: openP0All.length, sub: openP0All.length ? 'see below' : 'all clear' },
                    ].map(k => (
                      <div key={k.label} className="bg-gray-900 border border-gray-800 rounded-xl p-3">
                        <p className="text-[10px] text-gray-500 uppercase tracking-wider">{k.label}</p>
                        <p className="text-2xl font-bold text-gray-100 leading-tight mt-1">{k.value}</p>
                        <p className="text-[11px] text-gray-600">{k.sub}</p>
                      </div>
                    ))}
                  </div>

                  <div className="bg-gray-900 border border-gray-800 rounded-xl p-4">
                    <p className="text-xs text-gray-500 uppercase tracking-wider mb-3">Needs attention <span className="normal-case text-gray-600">— active</span></p>
                    <div className="space-y-1">
                      {([
                        ['flagged', '🚩 Flagged'], ['understaffed', '🟠 Potentially understaffed'], ['nopp', '📉 No production plan'],
                        ['overdue', '⏰ Overdue milestones'], ['p0', '🔥 Open P0 / P00'], ['noteam', '👻 No team'],
                      ] as const).map(([k, label]) => {
                        const n = activeInfo.filter(x => x.i.flags[k]).length
                        return (
                          <button key={k} onClick={() => setCmFilter(k)} className={`w-full flex items-center justify-between text-xs px-2 py-1.5 rounded-md transition-colors ${cmFilter === k ? 'bg-gray-800 text-white' : 'text-gray-400 hover:bg-gray-800/60 hover:text-gray-200'}`}>
                            <span>{label}</span>
                            <span className={`font-semibold ${n > 0 ? 'text-gray-100' : 'text-gray-600'}`}>{n}</span>
                          </button>
                        )
                      })}
                    </div>
                    <div className="flex h-2 rounded-full overflow-hidden mt-3 bg-gray-800">
                      {(['red', 'amber', 'green'] as const).map(h => {
                        const n = activeInfo.filter(x => x.i.health === h).length
                        return n > 0 ? <div key={h} title={`${n} ${h}`} className={h === 'red' ? 'bg-red-500' : h === 'amber' ? 'bg-amber-400' : 'bg-emerald-500'} style={{ width: `${(n / Math.max(activeAll.length, 1)) * 100}%` }} /> : null
                      })}
                    </div>
                    <p className="text-[11px] text-gray-600 mt-1.5">Health mix across {activeAll.length} active projects</p>
                  </div>

                  <div className="bg-gray-900 border border-gray-800 rounded-xl p-4">
                    <div className="flex items-center justify-between mb-3">
                      <p className="text-xs text-gray-500 uppercase tracking-wider">P00 / P0 — needs attention</p>
                      <span className={`text-xs font-semibold px-2 py-0.5 rounded-full ${openP0All.length > 0 ? 'bg-red-500/10 text-red-400' : 'bg-emerald-500/10 text-emerald-400'}`}>{openP0All.length}</span>
                    </div>
                    {openP0All.length === 0 ? (
                      <p className="text-xs text-gray-600">No open P0 milestones. 🎉</p>
                    ) : (
                      <div className="space-y-2 max-h-[40vh] overflow-y-auto">
                        {openP0All.map(m => {
                          const proj = projects.find(p => p.id === m.project_id)
                          const overdue = m.due_date && m.due_date < today
                          return (
                            <button key={m.id} onClick={() => jumpTo(m.project_id)} className="w-full text-left border border-gray-800 rounded-lg px-3 py-2 hover:border-gray-600 transition-colors">
                              <p className="text-sm text-gray-200 leading-tight">{m.title}</p>
                              <div className="flex items-center justify-between mt-1">
                                <span className="text-xs text-gray-500">{proj?.name}</span>
                                {m.due_date && <span className={`text-xs ${overdue ? 'text-red-400 font-medium' : 'text-gray-500'}`}>{overdue ? '⚠ overdue · ' : 'due '}{m.due_date}</span>}
                              </div>
                            </button>
                          )
                        })}
                      </div>
                    )}
                  </div>
                </div>
              </div>
            </div>
          )
        })()}

        {/* Utilization Tab — teams × queues */}
        {tab === 'utilization' && (() => {
          const all = utilTeams
          const statusOf = new Map(all.map(t => [t.id, utilStatus(t)]))
          const rating = utilRating(all)
          const ratingTone = (r: number | null) =>
            r == null ? 'text-gray-500' : r >= 8 ? 'text-violet-300' : r >= 6 ? 'text-emerald-400' : r >= 4 ? 'text-amber-400' : 'text-red-400'
          const ratingLabel = (r: number) => r >= 8.5 ? 'Excellent' : r >= 7 ? 'Strong' : r >= 5 ? 'Fair' : r >= 3 ? 'Needs work' : 'Poor'
          const counts = {
            utilized: all.filter(t => statusOf.get(t.id)!.status === 'utilized').length,
            under: all.filter(t => statusOf.get(t.id)!.status === 'under').length,
            not: all.filter(t => statusOf.get(t.id)!.status === 'not').length,
          }
          const tierCounts = Object.fromEntries(UTIL_TIERS.map(k => [k, all.filter(t => utilTier(t) === k).length])) as Record<string, number>
          const unscored = all.filter(t => utilTier(t) === null).length
          const scoredCount = all.length - unscored
          const free = all.filter(utilIsFree)
          // What-if: staffing every free team with B-tier work
          const whatIf = rating != null && free.length > 0 && scoredCount > 0
            ? (rating * scoredCount + free.length * (UTIL_POINTS.B - UTIL_POINTS.C)) / scoredCount
            : null

          // Leads in sheet order
          const leads: string[] = []
          all.forEach(t => { const l = t.lead ?? ''; if (!leads.includes(l)) leads.push(l) })

          const q = utilSearch.trim().toLowerCase()
          const visible = all.filter(t => {
            if (utilLead && (t.lead ?? '') !== (utilLead === '__none__' ? '' : utilLead)) return false
            if (utilFilter !== 'all' && statusOf.get(t.id)!.status !== utilFilter) return false
            if (!q) return true
            return [t.lead, t.supervisor, t.sto, t.ops, t.assignment_name, t.notes].some(v => (v ?? '').toLowerCase().includes(q))
          })
          const groups = leads.map(l => ({ lead: l, teams: visible.filter(t => (t.lead ?? '') === l) })).filter(g => g.teams.length > 0)

          // ---- Cell editors (save on blur / change) ----
          const textCell = (t: UtilTeam, field: 'lead' | 'supervisor' | 'sto' | 'ops' | 'assignment_name' | 'notes', opts: { placeholder?: string; muted?: boolean; list?: boolean } = {}) => (
            <input
              key={`${t.id}-${field}-${t[field] ?? ''}`}
              defaultValue={t[field] ?? ''}
              placeholder={opts.placeholder ?? '—'}
              list={opts.list ? 'util-people' : undefined}
              onBlur={e => { const v = e.target.value.trim() || null; if (v !== (t[field] ?? null)) updateUtilTeam(t.id, { [field]: v }) }}
              onKeyDown={e => { if (e.key === 'Enter') (e.target as HTMLInputElement).blur() }}
              className={`w-full min-w-0 bg-transparent px-1.5 py-1 rounded text-sm placeholder-gray-700 hover:bg-gray-800/60 focus:bg-gray-800 focus:outline-none focus:ring-1 focus:ring-[#193a29] ${opts.muted ? 'text-gray-500' : 'text-gray-200'}`}
            />
          )
          const complexityCell = (t: UtilTeam, field: 'queue_complexity' | 'assignment_complexity', disabled = false) => {
            const v = t[field]
            return disabled ? <span className="text-xs text-gray-700 px-1.5">—</span> : (
              <select
                value={v ?? ''}
                onChange={e => updateUtilTeam(t.id, { [field]: e.target.value || null })}
                className={`text-xs font-medium px-2 py-0.5 pr-5 rounded-full cursor-pointer appearance-none focus:outline-none focus:ring-1 focus:ring-gray-500 ${
                  v === 'High' ? 'bg-rose-500/10 text-rose-300' : v === 'Mid' ? 'bg-amber-500/10 text-amber-300' : v === 'Low' ? 'bg-teal-500/10 text-teal-300' : 'bg-gray-800 text-gray-500'
                }`}
                style={{ backgroundImage: 'none' }}
              >
                <option value="" className="bg-gray-900 text-gray-100">—</option>
                {['Low', 'Mid', 'High'].map(o => <option key={o} value={o} className="bg-gray-900 text-gray-100">{o}</option>)}
              </select>
            )
          }
          const tierPill = (k: string | null, extra = '') => (
            <span className={`inline-flex items-center justify-center w-7 h-7 rounded-lg text-sm font-bold ${k ? UTIL_TIER_INFO[k].cls : 'bg-gray-800 text-gray-600'} ${extra}`}>{k ?? '?'}</span>
          )
          const tierCell = (t: UtilTeam) => utilIsFree(t) ? (
            <span className="flex items-center gap-1.5" title="Not assigned counts as C">{tierPill('C', 'opacity-60')}<span className="text-[11px] text-gray-600">not assigned</span></span>
          ) : (
            <select
              value={t.tier ?? ''}
              onChange={e => updateUtilTeam(t.id, { tier: e.target.value || null })}
              title={t.tier ? `${UTIL_TIER_INFO[t.tier].label}: ${UTIL_TIER_INFO[t.tier].detail}` : 'Set a utilization score'}
              className={`w-12 h-7 text-sm font-bold text-center rounded-lg cursor-pointer appearance-none focus:outline-none focus:ring-1 focus:ring-gray-500 ${t.tier ? UTIL_TIER_INFO[t.tier].cls : 'bg-gray-800 text-gray-500 border border-dashed border-gray-600'}`}
              style={{ backgroundImage: 'none' }}
            >
              <option value="" className="bg-gray-900 text-gray-100">?</option>
              {UTIL_TIERS.map(k => <option key={k} value={k} className="bg-gray-900 text-gray-100">{k} — {UTIL_TIER_INFO[k].detail}</option>)}
            </select>
          )
          const prioCell = (t: UtilTeam) => utilIsFree(t) ? (
            <select
              value={t.assign_priority ?? ''}
              onChange={e => updateUtilTeam(t.id, { assign_priority: e.target.value ? Number(e.target.value) : null })}
              title="Priority on giving this team an assignment (1 = first)"
              className={`text-xs font-semibold px-2 py-0.5 pr-4 rounded-full cursor-pointer appearance-none focus:outline-none ${
                t.assign_priority === 1 ? 'bg-red-500/15 text-red-300' : t.assign_priority === 2 ? 'bg-amber-500/15 text-amber-300' : t.assign_priority === 3 ? 'bg-gray-700 text-gray-300' : 'bg-gray-800 text-gray-500'
              }`}
              style={{ backgroundImage: 'none' }}
            >
              <option value="" className="bg-gray-900 text-gray-100">—</option>
              {[1, 2, 3].map(n => <option key={n} value={n} className="bg-gray-900 text-gray-100">P{n}</option>)}
            </select>
          ) : <span className="text-xs text-gray-700">—</span>
          const statusCell = (t: UtilTeam) => {
            const st = statusOf.get(t.id)!
            return (
              <div className="min-w-0">
                <span className={`inline-flex items-center gap-1 text-[11px] font-semibold px-2 py-0.5 rounded-full whitespace-nowrap ${
                  st.status === 'utilized' ? 'bg-emerald-500/10 text-emerald-400' : st.status === 'under' ? 'bg-amber-500/10 text-amber-400' : 'bg-red-500/10 text-red-400'
                }`}>
                  {st.status === 'utilized' ? '● Utilized' : st.status === 'under' ? '◐ Under-utilized' : '○ Not utilized'}
                </span>
                {st.status !== 'utilized' || t.reason ? (
                  <input
                    key={`${t.id}-reason-${t.reason ?? ''}`}
                    defaultValue={t.reason ?? ''}
                    placeholder={st.reason}
                    title="Why? Type your own explanation, or leave blank to use the automatic one"
                    onBlur={e => { const v = e.target.value.trim() || null; if (v !== (t.reason ?? null)) updateUtilTeam(t.id, { reason: v }) }}
                    onKeyDown={e => { if (e.key === 'Enter') (e.target as HTMLInputElement).blur() }}
                    className="block w-full mt-1 bg-transparent px-1 py-0.5 rounded text-xs text-gray-300 placeholder-gray-500 hover:bg-gray-800/60 focus:bg-gray-800 focus:outline-none focus:ring-1 focus:ring-[#193a29]"
                  />
                ) : st.reason ? <p className="text-[11px] text-gray-600 mt-1 px-1">{st.reason}</p> : null}
              </div>
            )
          }

          const chip = (k: typeof utilFilter, label: string, n: number) => (
            <button
              key={k}
              onClick={() => setUtilFilter(k)}
              className={`text-xs px-3 py-1 rounded-full border transition-colors whitespace-nowrap ${utilFilter === k ? 'text-white border-transparent' : 'border-gray-700 text-gray-400 hover:text-gray-200'}`}
              style={utilFilter === k ? { backgroundColor: '#193a29' } : {}}
            >{label} <span className={utilFilter === k ? 'text-white/70' : 'text-gray-500'}>{n}</span></button>
          )

          if (all.length === 0) {
            return (
              <div className="border border-dashed border-gray-700 rounded-2xl p-12 text-center max-w-2xl mx-auto">
                <p className="text-3xl mb-3">📊</p>
                <p className="text-gray-100 font-semibold mb-1">No teams yet</p>
                <p className="text-sm text-gray-500 mb-6">All teams were removed. Add one, or restore the original sheet.</p>
                <div className="flex items-center justify-center gap-3">
                  <button onClick={() => openUtilEdit()} className={btnPrimary} style={{ backgroundColor: '#193a29' }}>+ Add a team</button>
                  <button onClick={resetUtilToSheet} className="text-sm px-4 py-2 rounded-lg border border-gray-700 text-gray-300 hover:text-white transition-colors">↺ Restore original sheet</button>
                </div>
              </div>
            )
          }

          const gaugePct = rating == null ? 0 : ((rating - 1) / 9) * 100
          return (
            <div className="space-y-5">
              {utilLocalOnly && (
                <div className="border border-amber-500/30 bg-amber-500/5 rounded-xl px-4 py-3 text-sm text-amber-300">
                  Showing the original sheet, but changes won’t be saved yet — the <code className="text-amber-200">utilization_teams</code> table doesn’t exist in Supabase. Run the setup SQL, then refresh.
                </div>
              )}
              <datalist id="util-people">
                {[...new Set([...staff.map(s => s.name), ...all.flatMap(t => [t.lead, t.supervisor, t.sto, t.ops]).filter((x): x is string => !!x)])]
                  .sort((a, b) => a.localeCompare(b)).map(n => <option key={n} value={n} />)}
              </datalist>

              {/* ===== Hero: Pod rating · tier mix · next to staff ===== */}
              <div className="grid grid-cols-1 lg:grid-cols-3 gap-4">
                <div className="bg-gray-900 border border-gray-800 rounded-2xl p-6">
                  <p className="text-xs text-gray-500 uppercase tracking-wider">Utilization Rating for the Pod</p>
                  <div className="flex items-end gap-3 mt-2">
                    <p className={`text-6xl font-bold leading-none ${ratingTone(rating)}`}>{rating == null ? '—' : rating.toFixed(1)}</p>
                    <div className="pb-1">
                      <p className="text-sm text-gray-500">/ 10</p>
                      {rating != null && <p className={`text-sm font-semibold ${ratingTone(rating)}`}>{ratingLabel(rating)}</p>}
                    </div>
                  </div>
                  <div className="relative mt-5">
                    <div className="h-2.5 rounded-full" style={{ background: 'linear-gradient(90deg, #ef4444, #f59e0b 35%, #10b981 70%, #8b5cf6)' }} />
                    {rating != null && <div className="absolute -top-1 w-1.5 rounded-full bg-white shadow" style={{ left: `calc(${gaugePct}% - 3px)`, height: 18 }} />}
                    <div className="flex justify-between text-[10px] text-gray-600 mt-1.5"><span>1 · all C</span><span>10 · all S</span></div>
                  </div>
                  <p className="text-xs text-gray-500 mt-4">
                    Average of {scoredCount} team{scoredCount === 1 ? '' : 's'} (S = 10, A = 7, B = 4, C = 1).
                    {unscored > 0 && <span className="text-amber-400"> {unscored} assigned team{unscored === 1 ? ' has' : 's have'} no score yet.</span>}
                  </p>
                  {whatIf != null && (
                    <p className="text-xs text-gray-400 mt-2 bg-gray-800/50 rounded-lg px-3 py-2">
                      💡 Staffing the <span className="text-gray-200 font-medium">{free.length} free team{free.length === 1 ? '' : 's'}</span> with B-tier work would lift the rating to <span className={`font-semibold ${ratingTone(whatIf)}`}>{whatIf.toFixed(1)}</span>.
                    </p>
                  )}
                </div>

                <div className="bg-gray-900 border border-gray-800 rounded-2xl p-6">
                  <div className="flex items-center justify-between">
                    <p className="text-xs text-gray-500 uppercase tracking-wider">Teams by tier</p>
                    <p className="text-xs text-gray-500">{all.length} teams</p>
                  </div>
                  <div className="flex h-3 rounded-full overflow-hidden mt-4 bg-gray-800">
                    {UTIL_TIERS.map(k => tierCounts[k] > 0 ? (
                      <div key={k} title={`${k}: ${tierCounts[k]}`} className={k === 'S' ? 'bg-violet-500' : k === 'A' ? 'bg-emerald-500' : k === 'B' ? 'bg-sky-500' : 'bg-gray-500'} style={{ width: `${(tierCounts[k] / all.length) * 100}%` }} />
                    ) : null)}
                  </div>
                  <div className="space-y-2.5 mt-4">
                    {UTIL_TIERS.map(k => (
                      <div key={k} className="flex items-center gap-3">
                        {tierPill(k)}
                        <div className="flex-1 min-w-0">
                          <p className="text-xs text-gray-300 leading-tight">{UTIL_TIER_INFO[k].detail}</p>
                        </div>
                        <span className="text-sm font-semibold text-gray-100 w-6 text-right">{tierCounts[k]}</span>
                      </div>
                    ))}
                  </div>
                  <div className="grid grid-cols-3 gap-2 mt-5 pt-4 border-t border-gray-800 text-center">
                    <div><p className="text-xl font-bold text-emerald-400">{counts.utilized}</p><p className="text-[11px] text-gray-500">Utilized</p></div>
                    <div><p className="text-xl font-bold text-amber-400">{counts.under}</p><p className="text-[11px] text-gray-500">Under-utilized</p></div>
                    <div><p className="text-xl font-bold text-red-400">{counts.not}</p><p className="text-[11px] text-gray-500">Not utilized</p></div>
                  </div>
                </div>

                <div className="bg-gray-900 border border-gray-800 rounded-2xl p-6">
                  <div className="flex items-center justify-between mb-3">
                    <p className="text-xs text-gray-500 uppercase tracking-wider">Next up to staff</p>
                    <span className={`text-xs font-semibold px-2 py-0.5 rounded-full ${free.length ? 'bg-red-500/10 text-red-400' : 'bg-emerald-500/10 text-emerald-400'}`}>{free.length} free</span>
                  </div>
                  {free.length === 0 ? (
                    <p className="text-xs text-gray-600">Every team has an assignment. 🎉</p>
                  ) : (
                    <div className="space-y-1.5 max-h-[300px] overflow-y-auto pr-1">
                      {[...free].sort((a, b) => (a.assign_priority ?? 9) - (b.assign_priority ?? 9) || (COMPLEXITY_RANK[b.queue_complexity ?? ''] ?? 0) - (COMPLEXITY_RANK[a.queue_complexity ?? ''] ?? 0)).map(t => (
                        <div key={t.id} className="flex items-center gap-2 bg-gray-800/40 rounded-lg px-3 py-2">
                          <span className={`text-[11px] font-bold w-7 text-center rounded ${t.assign_priority === 1 ? 'text-red-300' : t.assign_priority === 2 ? 'text-amber-300' : 'text-gray-400'}`}>{t.assign_priority ? `P${t.assign_priority}` : '—'}</span>
                          <div className="flex-1 min-w-0">
                            <p className="text-sm text-gray-200 truncate">{t.sto ?? 'No STO'}{t.ops ? <span className="text-gray-500"> + {t.ops}</span> : null}</p>
                            <p className="text-[11px] text-gray-500 truncate">{t.lead ?? 'No lead'}{t.supervisor ? ` · sup. ${t.supervisor}` : ''}</p>
                          </div>
                          {t.queue_complexity && <span className={`text-[10px] px-1.5 py-0.5 rounded-full ${t.queue_complexity === 'High' ? 'bg-rose-500/10 text-rose-300' : t.queue_complexity === 'Mid' ? 'bg-amber-500/10 text-amber-300' : 'bg-teal-500/10 text-teal-300'}`}>{t.queue_complexity}</span>}
                        </div>
                      ))}
                    </div>
                  )}
                </div>
              </div>

              {/* ===== By lead ===== */}
              <div className="flex gap-3 overflow-x-auto pb-1">
                {leads.map(l => {
                  const lt = all.filter(t => (t.lead ?? '') === l)
                  const r = utilRating(lt)
                  const used = lt.filter(t => statusOf.get(t.id)!.status === 'utilized').length
                  const key = l || '__none__'
                  const on = utilLead === key
                  return (
                    <button
                      key={key}
                      onClick={() => setUtilLead(on ? '' : key)}
                      className={`shrink-0 text-left rounded-xl border px-4 py-3 min-w-[170px] transition-colors ${on ? 'border-[#2d6a4a] bg-[#193a29]/30' : 'border-gray-800 bg-gray-900 hover:border-gray-600'}`}
                    >
                      <p className="text-sm font-semibold text-gray-100 truncate">{l || 'No lead'}</p>
                      <div className="flex items-end justify-between mt-1">
                        <p className={`text-2xl font-bold leading-none ${ratingTone(r)}`}>{r == null ? '—' : r.toFixed(1)}</p>
                        <p className="text-[11px] text-gray-500 text-right">{lt.length} team{lt.length === 1 ? '' : 's'}<br />{used} utilized</p>
                      </div>
                      <div className="flex h-1 rounded-full overflow-hidden mt-2 bg-gray-800">
                        {UTIL_TIERS.map(k => { const n = lt.filter(t => utilTier(t) === k).length; return n ? <div key={k} className={k === 'S' ? 'bg-violet-500' : k === 'A' ? 'bg-emerald-500' : k === 'B' ? 'bg-sky-500' : 'bg-gray-500'} style={{ width: `${(n / lt.length) * 100}%` }} /> : null })}
                      </div>
                    </button>
                  )
                })}
              </div>

              {/* ===== Toolbar ===== */}
              <div className="flex flex-wrap items-center gap-2">
                <div className="relative min-w-[240px] flex-1 max-w-md">
                  <input className={inputClass + ' w-full pl-9'} placeholder="Search people, queues, notes…" value={utilSearch} onChange={e => setUtilSearch(e.target.value)} />
                  <span className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-500 text-sm pointer-events-none">🔍</span>
                  {utilSearch && <button onClick={() => setUtilSearch('')} className="absolute right-3 top-1/2 -translate-y-1/2 text-gray-500 hover:text-gray-300 text-sm">✕</button>}
                </div>
                {chip('all', 'All', all.length)}
                {chip('utilized', '● Utilized', counts.utilized)}
                {chip('under', '◐ Under-utilized', counts.under)}
                {chip('not', '○ Not utilized', counts.not)}
                {utilLead && <button onClick={() => setUtilLead('')} className="text-xs text-sky-400 hover:text-sky-300">Lead: {utilLead === '__none__' ? 'No lead' : utilLead} ✕</button>}
                <div className="flex items-center gap-2 ml-auto">
                  <button onClick={() => openUtilEdit()} className="text-xs px-3 py-1.5 rounded-lg text-white transition-colors hover:brightness-125" style={{ backgroundColor: '#193a29' }}>+ Add team</button>
                  <button onClick={resetUtilToSheet} className="text-xs px-3 py-1.5 rounded-lg border border-gray-700 text-gray-400 hover:text-white hover:border-gray-500 transition-colors" title="Replace everything with the original 32-team sheet">↺ Reset to original sheet</button>
                  <button onClick={exportUtilCSV} className="text-xs px-3 py-1.5 rounded-lg border border-gray-700 text-gray-300 hover:text-white hover:border-gray-500 transition-colors">⬇ Export CSV</button>
                </div>
              </div>

              {/* ===== Teams table ===== */}
              <div className="border border-gray-800 rounded-xl bg-gray-900/40 overflow-x-auto">
                <table className="w-full text-sm">
                  <thead>
                    <tr className="text-left text-gray-500 border-b border-gray-800 text-[11px] uppercase tracking-wider">
                      <th className="py-3 pl-4 pr-2 font-medium w-8">#</th>
                      <th className="py-3 px-1 font-medium" colSpan={4}>
                        <span className="text-gray-600 normal-case tracking-normal">Configuration —</span> Lead · Supervisor · STO · Ops
                      </th>
                      <th className="py-3 px-1 font-medium whitespace-nowrap">Team complexity</th>
                      <th className="py-3 px-1 font-medium border-l border-gray-800 pl-3">Current assignment</th>
                      <th className="py-3 px-1 font-medium whitespace-nowrap">Queue complexity</th>
                      <th className="py-3 px-1 font-medium min-w-[220px]">Utilization status · why</th>
                      <th className="py-3 px-1 font-medium text-center" title="S / A / B / C">Score</th>
                      <th className="py-3 px-1 font-medium text-center whitespace-nowrap" title="Priority on giving this team an assignment">Assign prio</th>
                      <th className="py-3 px-1 font-medium min-w-[180px]">Notes</th>
                      <th className="py-3 pr-3 w-20" />
                    </tr>
                  </thead>
                  <tbody>
                    {groups.length === 0 && (
                      <tr><td colSpan={13} className="py-10 text-center text-sm text-gray-600">No teams match these filters.</td></tr>
                    )}
                    {groups.map(g => {
                      const r = utilRating(g.teams)
                      return (
                        <Fragment key={g.lead || '__none__'}>
                          <tr className="bg-gray-800/30 border-y border-gray-800">
                            <td colSpan={13} className="py-2 pl-4 pr-4">
                              <div className="flex items-center gap-3">
                                <span className="text-sm font-semibold text-gray-100">{g.lead || 'No lead'}</span>
                                <span className="text-xs text-gray-500">{g.teams.length} team{g.teams.length === 1 ? '' : 's'}</span>
                                <span className="text-xs text-gray-500">· rating <span className={`font-semibold ${ratingTone(r)}`}>{r == null ? '—' : r.toFixed(1)}</span></span>
                              </div>
                            </td>
                          </tr>
                          {g.teams.map(t => {
                            const st = statusOf.get(t.id)!
                            const isFree = utilIsFree(t)
                            return (
                              <tr key={t.id} className={`border-b border-gray-800/50 last:border-0 align-top hover:bg-gray-800/20 ${st.status === 'not' ? 'bg-red-500/[0.03]' : ''}`}>
                                <td className="py-2 pl-4 pr-2 text-xs text-gray-600 pt-3">{all.indexOf(t) + 1}</td>
                                <td className="py-1.5 px-1 w-[11%]">{textCell(t, 'lead', { muted: true, list: true })}</td>
                                <td className="py-1.5 px-1 w-[11%]">{textCell(t, 'supervisor', { list: true })}</td>
                                <td className="py-1.5 px-1 w-[12%]">{textCell(t, 'sto', { list: true })}</td>
                                <td className="py-1.5 px-1 w-[11%]">{textCell(t, 'ops', { list: true })}</td>
                                <td className="py-2 px-1 pt-2.5">{complexityCell(t, 'queue_complexity')}</td>
                                <td className="py-1.5 px-1 border-l border-gray-800 pl-3 w-[14%]">
                                  <input
                                    key={`${t.id}-an-${t.assignment_name ?? ''}`}
                                    defaultValue={isFree ? '' : (t.assignment_name ?? '')}
                                    placeholder="Free"
                                    onBlur={e => { const v = e.target.value.trim() || 'Free'; if (v !== (t.assignment_name ?? 'Free')) updateUtilTeam(t.id, { assignment_name: v }) }}
                                    onKeyDown={e => { if (e.key === 'Enter') (e.target as HTMLInputElement).blur() }}
                                    className={`w-full min-w-0 bg-transparent px-1.5 py-1 rounded text-sm hover:bg-gray-800/60 focus:bg-gray-800 focus:outline-none focus:ring-1 focus:ring-[#193a29] ${isFree ? 'placeholder-red-400/70 italic' : 'text-gray-100 font-medium'}`}
                                  />
                                </td>
                                <td className="py-2 px-1 pt-2.5">{complexityCell(t, 'assignment_complexity', isFree)}</td>
                                <td className="py-2 px-1">{statusCell(t)}</td>
                                <td className="py-2 px-1 text-center">{tierCell(t)}</td>
                                <td className="py-2 px-1 pt-2.5 text-center">{prioCell(t)}</td>
                                <td className="py-1.5 px-1">{textCell(t, 'notes', { placeholder: 'Add a note…', muted: true })}</td>
                                <td className="py-2 pr-3 pt-2 text-right whitespace-nowrap">
                                  <button onClick={() => openUtilEdit(t)} title="Edit team" className="text-xs px-2 py-1 rounded-md text-gray-400 hover:text-white hover:bg-gray-800 transition-colors">✎ Edit</button>
                                  <button className={btnDanger + ' ml-1'} onClick={() => deleteUtilTeam(t)} title="Remove team">×</button>
                                </td>
                              </tr>
                            )
                          })}
                        </Fragment>
                      )
                    })}
                  </tbody>
                </table>
              </div>
              <p className="text-xs text-gray-600">
                Status is automatic: <span className="text-red-400">Not utilized</span> = no assignment (Free) · <span className="text-amber-400">Under-utilized</span> = C-tier work, or a team assigned to a less complex queue than it can handle · <span className="text-emerald-400">Utilized</span> = everything else. Edit any cell in place, or use ✎ Edit for the full form. Type in the “why” line to override the explanation. Click a lead card to filter.
              </p>
            </div>
          )
        })()}

        {/* Org Chart Tab */}
        {tab === 'orgchart' && (() => {
          const inChart = new Set(orgChart.map(n => n.staff_id))
          const root = orgChart.find(n => !n.parent_staff_id)
          const childrenOf = (sid: string) => orgChart.filter(n => n.parent_staff_id === sid).map(n => n.staff_id)
          const parentOf = (sid: string) => orgChart.find(n => n.staff_id === sid)?.parent_staff_id ?? null
          const isConsultant = (sid: string) => staff.find(s => s.id === sid)?.position === 'GenAI Consultant'
          const descendantCount = (sid: string): number => {
            let count = 0
            const stack = [...childrenOf(sid)]
            while (stack.length) { const c = stack.pop()!; count++; stack.push(...childrenOf(c)) }
            return count
          }
          const orgColor = (sid: string) => {
            if (!orgColorsOn) return 'border-gray-700 bg-gray-900'                                // colors off → all neutral
            if (root && sid === root.staff_id) return 'border-amber-500/60 bg-amber-500/15'      // org leader
            if (isConsultant(sid)) return 'border-violet-500/50 bg-violet-500/15'                 // consultants
            if (root && parentOf(sid) === root.staff_id) return 'border-emerald-500/50 bg-emerald-500/15' // direct reports
            return 'border-sky-500/50 bg-sky-500/15'                                              // everyone else
          }
          const isAncestorOrSelf = (ancestor: string, node: string) => {
            let cur: string | null = node
            while (cur) { if (cur === ancestor) return true; cur = parentOf(cur) }
            return false
          }
          const leftList = [...staff]
            .filter(s => (orgUnassignedOnly ? !inChart.has(s.id) : true))
            .sort((a, b) => a.name.localeCompare(b.name))

          const onDropTarget = (targetStaffId: string | null) => (e: React.DragEvent) => {
            e.preventDefault(); e.stopPropagation(); setDragOverOrgId(null)
            const type = e.dataTransfer.getData('type')
            const id = e.dataTransfer.getData('id')
            if (!id) return
            if (type === 'org-add') {
              if (!inChart.has(id)) addToOrg(id, targetStaffId)
            } else if (type === 'org-move') {
              // Prevent dropping a node onto itself or one of its own descendants
              if (targetStaffId && (id === targetStaffId || isAncestorOrSelf(id, targetStaffId))) return
              moveOrgNode(id, targetStaffId)
            }
          }

          const sortKids = (ids: string[]) => [...ids].sort((a, b) => {
            const ra = isConsultant(a) ? 0 : 1
            const rb = isConsultant(b) ? 0 : 1
            if (ra !== rb) return ra - rb
            return (staff.find(s => s.id === a)?.name ?? '').localeCompare(staff.find(s => s.id === b)?.name ?? '')
          })
          const visibleKids = (staffId: string) => {
            const all = sortKids(childrenOf(staffId))
            return orgShowConsultants ? all : all.filter(k => !isConsultant(k))
          }
          const cardInner = (staffId: string) => {
            const s = staff.find(x => x.id === staffId)
            const allKids = childrenOf(staffId)
            const kids = visibleKids(staffId)
            const over = dragOverOrgId === staffId
            const collapsed = !!collapsedOrg[staffId]
            return (
              <div
                draggable
                onDragStart={e => { e.dataTransfer.setData('type', 'org-move'); e.dataTransfer.setData('id', staffId) }}
                onDragOver={e => { e.preventDefault(); setDragOverOrgId(staffId) }}
                onDragLeave={() => setDragOverOrgId(null)}
                onDrop={onDropTarget(staffId)}
                className={`org-node inline-block text-left rounded-lg border px-3 py-2 cursor-grab active:cursor-grabbing transition-all ${over ? 'border-[#193a29] bg-[#193a29]/30 scale-105' : orgColor(staffId)}`}
              >
                <div className="flex items-center gap-2">
                  <span className="text-sm font-medium text-gray-100 whitespace-nowrap">{s?.emoji ? `${s.emoji} ` : ''}{s?.name ?? 'Unknown'}</span>
                  {allKids.length > 0 && <span className="text-[10px] text-gray-500" title={`${descendantCount(staffId)} people below`}>({descendantCount(staffId)})</span>}
                  <button
                    onClick={() => { if (allKids.length > 0) setPendingOrgRemove(staffId); else removeFromOrg(staffId) }}
                    title="Remove from chart"
                    className="text-gray-500 hover:text-red-400 leading-none text-xs"
                  >✕</button>
                </div>
                {s?.position && <p className="text-[11px] text-gray-500 whitespace-nowrap">{s.position}</p>}
                {orgShowProjects && (() => {
                  const projs = assignments
                    .filter(a => a.staff_id === staffId)
                    .map(a => projects.find(pr => pr.id === a.project_id)?.name)
                    .filter((n): n is string => !!n)
                    .sort((a, b) => a.localeCompare(b))
                  return projs.length > 0 ? (
                    <div className="flex flex-wrap gap-1 mt-1.5 max-w-[220px]">
                      {projs.map((n, i) => <span key={i} className="text-[10px] px-1.5 py-0.5 rounded-full bg-gray-800 text-gray-300 whitespace-nowrap">{n}</span>)}
                    </div>
                  ) : <p className="text-[10px] text-gray-600 mt-1.5 italic">No projects</p>
                })()}
                {kids.length > 0 && (
                  <button
                    onClick={() => setCollapsedOrg(prev => ({ ...prev, [staffId]: !prev[staffId] }))}
                    title={collapsed ? 'Expand' : 'Collapse'}
                    className="mt-1 text-[10px] text-gray-500 hover:text-gray-200 transition-colors"
                  >
                    {collapsed ? `▸ ${kids.length} report${kids.length === 1 ? '' : 's'}` : '▾ collapse'}
                  </button>
                )}
              </div>
            )
          }
          const renderNode = (staffId: string) => {
            const kids = visibleKids(staffId)
            const collapsed = !!collapsedOrg[staffId]
            return (
              <li key={staffId}>
                {cardInner(staffId)}
                {kids.length > 0 && !collapsed && <ul>{kids.map(renderNode)}</ul>}
              </li>
            )
          }
          const renderNodeV = (staffId: string) => {
            const kids = visibleKids(staffId)
            const collapsed = !!collapsedOrg[staffId]
            return (
              <div key={staffId} className="org-vrow">
                {cardInner(staffId)}
                {kids.length > 0 && !collapsed && <div className="org-vchildren">{kids.map(renderNodeV)}</div>}
              </div>
            )
          }

          return (
            <div className="flex gap-6 items-start">
              {/* Left: staff list */}
              <div className="w-56 shrink-0 sticky top-6">
                <div className="mb-3">{toggleSwitch(orgUnassignedOnly, () => setOrgUnassignedOnly(v => !v), 'Unassigned only')}</div>
                <p className="text-xs text-gray-500 uppercase tracking-wider mb-2">Staff — drag into chart</p>
                <div className="flex flex-col gap-2 max-h-[70vh] overflow-y-auto pr-1">
                  {leftList.length === 0 ? (
                    <p className="text-xs text-gray-600">{orgUnassignedOnly ? 'Everyone is on the chart.' : 'No staff.'}</p>
                  ) : leftList.map(s => {
                    const placed = inChart.has(s.id)
                    return (
                      <div
                        key={s.id}
                        draggable={!placed}
                        onDragStart={e => { if (placed) return; e.dataTransfer.setData('type', 'org-add'); e.dataTransfer.setData('id', s.id) }}
                        className={`border rounded-lg px-3 py-2 text-sm transition-all ${placed ? 'border-gray-800 bg-gray-900/40 text-gray-600 opacity-60' : 'border-gray-700 bg-gray-900 text-gray-200 cursor-grab active:cursor-grabbing hover:border-gray-500'}`}
                      >
                        <p className="font-medium leading-tight">{s.emoji ? `${s.emoji} ` : ''}{s.name}{placed && <span className="text-[10px] ml-1">· on chart</span>}</p>
                        {s.position && <p className="text-xs text-gray-500 mt-0.5">{s.position}</p>}
                      </div>
                    )
                  })}
                </div>
              </div>

              {/* Right: the chart */}
              <div
                ref={orgScrollRef}
                className="flex-1 min-w-0 overflow-x-auto"
                onDragOver={orgAutoScroll}
                onDragLeave={stopOrgAutoScroll}
                onDrop={stopOrgAutoScroll}
              >
                {root && (
                  <div className="flex justify-between items-center gap-4 mb-3 flex-wrap">
                    <div className="flex items-center gap-4 flex-wrap text-[11px] text-gray-500">
                      {orgColorsOn && <>
                        <span className="flex items-center gap-1.5"><span className="w-2.5 h-2.5 rounded-full bg-amber-500/70 border border-amber-500" /> Org leader</span>
                        <span className="flex items-center gap-1.5"><span className="w-2.5 h-2.5 rounded-full bg-emerald-500/70 border border-emerald-500" /> Direct reports</span>
                        <span className="flex items-center gap-1.5"><span className="w-2.5 h-2.5 rounded-full bg-violet-500/70 border border-violet-500" /> Consultants</span>
                        <span className="flex items-center gap-1.5"><span className="w-2.5 h-2.5 rounded-full bg-sky-500/70 border border-sky-500" /> Everyone else</span>
                      </>}
                    </div>
                    <div className="flex items-center gap-3">
                      {toggleSwitch(orgColorsOn, () => setOrgColorsOn(v => !v), 'Colors')}
                      {toggleSwitch(orgShowConsultants, () => setOrgShowConsultants(v => !v), 'Show consultants')}
                      {toggleSwitch(orgShowProjects, () => setOrgShowProjects(v => !v), 'Show projects')}
                      {toggleSwitch(orgLayout === 'vertical', () => setOrgLayout(m => m === 'vertical' ? 'horizontal' : 'vertical'), 'Vertical')}
                      <button
                        onClick={() => { const c: Record<string, boolean> = {}; orgChart.forEach(n => { if (orgChart.some(x => x.parent_staff_id === n.staff_id)) c[n.staff_id] = true }); setCollapsedOrg(c) }}
                        className="text-xs px-3 py-1.5 rounded-lg border border-gray-700 text-gray-400 hover:text-gray-200 transition-colors"
                      >Collapse all</button>
                      <button
                        onClick={() => setCollapsedOrg({})}
                        className="text-xs px-3 py-1.5 rounded-lg border border-gray-700 text-gray-400 hover:text-gray-200 transition-colors"
                      >Expand all</button>
                      <div className="flex items-center border border-gray-700 rounded-lg overflow-hidden">
                        <button onClick={() => setOrgZoom(z => Math.max(0.4, Math.round((z - 0.1) * 10) / 10))} className="px-2.5 py-1 text-gray-400 hover:text-white hover:bg-gray-800 transition-colors leading-none" title="Zoom out">−</button>
                        <button onClick={() => setOrgZoom(1)} className="px-2 py-1 text-xs text-gray-400 hover:text-white tabular-nums leading-none" title="Reset zoom">{Math.round(orgZoom * 100)}%</button>
                        <button onClick={() => setOrgZoom(z => Math.min(2, Math.round((z + 0.1) * 10) / 10))} className="px-2.5 py-1 text-gray-400 hover:text-white hover:bg-gray-800 transition-colors leading-none" title="Zoom in">+</button>
                      </div>
                      <button onClick={() => exportOrgChart('download')} className="text-xs px-3 py-1.5 rounded-lg border border-gray-700 text-gray-400 hover:text-gray-200 transition-colors">Download PNG</button>
                      <button onClick={() => exportOrgChart('print')} className="text-xs px-3 py-1.5 rounded-lg border border-gray-700 text-gray-400 hover:text-gray-200 transition-colors">Print</button>
                    </div>
                  </div>
                )}
                {!root ? (
                  <div
                    onDragOver={e => { e.preventDefault(); setDragOverOrgId('root') }}
                    onDragLeave={() => setDragOverOrgId(null)}
                    onDrop={onDropTarget(null)}
                    className={`rounded-xl border-2 border-dashed min-h-[240px] flex items-center justify-center text-center transition-all ${dragOverOrgId === 'root' ? 'border-[#193a29] bg-[#193a29]/10' : 'border-gray-700'}`}
                  >
                    <p className="text-sm text-gray-500 px-6">Drag the top person here to start the org chart.<br /><span className="text-xs text-gray-600">Then drop others onto a card to place them underneath.</span></p>
                  </div>
                ) : (
                  <div style={{ zoom: orgZoom }}>
                    {orgLayout === 'vertical' ? (
                      <div id="org-chart-capture" className="org-vtree py-2 pl-2 pr-4">
                        {renderNodeV(root.staff_id)}
                      </div>
                    ) : (
                      <div id="org-chart-capture" className="org-tree py-4 pr-4">
                        <ul>{renderNode(root.staff_id)}</ul>
                      </div>
                    )}
                  </div>
                )}
              </div>
            </div>
          )
        })()}
      </div>
      </div>{/* end spin wrapper */}

      {renderDropModals()}

      {overviewOpen && (() => {
        const original = buildOverview('preview')
        const missing = original?.missing ?? []
        const live = buildOverview('preview', overviewIdInputs)
        return (
          <div className="fixed inset-0 bg-black/50 flex items-center justify-center z-50 p-6" onClick={() => setOverviewOpen(false)}>
            <div className="bg-gray-900 border border-gray-700 rounded-xl p-6 w-[560px] max-w-full shadow-xl" onClick={e => e.stopPropagation()}>
              <p className="text-sm font-semibold text-gray-100 mb-1">Preview — Milestones &amp; PP overview</p>
              <p className="text-xs text-gray-500 mb-3">This will post to <span className="text-gray-300">#code_pod</span>. STOs, plus supervisors &amp; leads in the cc, are @-mentioned in Slack (names shown here). Review before sending.</p>
              <pre className="text-xs text-gray-200 bg-gray-800/60 border border-gray-700 rounded-lg p-3 max-h-[40vh] overflow-y-auto whitespace-pre-wrap">{live?.text}</pre>
              {missing.length > 0 && (
                <div className="mt-4">
                  <p className="text-xs text-gray-400 mb-2">
                    {missing.length > 1 ? 'These people have' : 'This person has'} no Slack ID yet — add to tag them (Slack profile → ⋮ → Copy member ID). Saved to their staff profile. Leave blank to post their name instead.
                  </p>
                  <div className="space-y-2">
                    {missing.map(person => (
                      <div key={person.id} className="flex items-center gap-2">
                        <span className="text-xs text-gray-300 w-36 truncate">{person.name} <span className="text-gray-600">({person.role})</span></span>
                        <input
                          className={inputSmClass + ' flex-1'}
                          placeholder="U012ABC… (Slack member ID)"
                          value={overviewIdInputs[person.id] ?? ''}
                          onChange={e => setOverviewIdInputs(prev => ({ ...prev, [person.id]: e.target.value }))}
                        />
                      </div>
                    ))}
                  </div>
                </div>
              )}
              <div className="flex gap-2 justify-end mt-5">
                <button className={btnCancel} onClick={() => setOverviewOpen(false)}>Cancel</button>
                <button className={btnPrimary} style={{ backgroundColor: '#193a29' }} onClick={sendOverview}>Post to #code_pod</button>
              </div>
            </div>
          </div>
        )
      })()}

      {pendingSlackIds && (
        <div className="fixed inset-0 bg-black/50 flex items-center justify-center z-50" onClick={() => { setPendingSlackIds(null); setSlackIdInputs({}) }}>
          <div className="bg-gray-900 border border-gray-700 rounded-xl p-6 w-96 shadow-xl" onClick={e => e.stopPropagation()}>
            <p className="text-sm font-semibold text-gray-100 mb-1">Add Slack ID{pendingSlackIds.people.length > 1 ? 's' : ''}</p>
            <p className="text-xs text-gray-500 mb-4">
              {pendingSlackIds.people.length > 1 ? 'Some people' : 'This person'} on <span className="text-gray-300">{pendingSlackIds.project.name}</span> {pendingSlackIds.people.length > 1 ? "don't" : "doesn't"} have a Slack ID yet. Add it to tag them (Slack profile → ⋮ → Copy member ID). Saved to their staff profile.
            </p>
            <div className="space-y-2 mb-4">
              {pendingSlackIds.people.map((person, i) => (
                <div key={person.id} className="flex items-center gap-2">
                  <span className="text-xs text-gray-300 w-36 truncate">{person.name} <span className="text-gray-600">({person.role})</span></span>
                  <input
                    type="text"
                    autoFocus={i === 0}
                    className={inputSmClass + ' flex-1'}
                    placeholder="U0123ABCD"
                    value={slackIdInputs[person.id] ?? ''}
                    onChange={e => setSlackIdInputs(prev => ({ ...prev, [person.id]: e.target.value }))}
                    onKeyDown={e => e.key === 'Enter' && submitSlackIdsAndRequest()}
                  />
                </div>
              ))}
            </div>
            <div className="flex gap-2 justify-end">
              <button className={btnCancel} onClick={() => { setPendingSlackIds(null); setSlackIdInputs({}) }}>Cancel</button>
              <button className={btnEdit} onClick={() => { const p = pendingSlackIds.project; setPendingSlackIds(null); setSlackIdInputs({}); requestMilestoneUpdate(p) }}>Skip &amp; send</button>
              <button className={btnPrimary} style={{ backgroundColor: '#193a29' }} onClick={submitSlackIdsAndRequest}>Save &amp; request</button>
            </div>
          </div>
        </div>
      )}

      {leadIdPrompt && (
        <div className="fixed inset-0 bg-black/50 flex items-center justify-center z-50" onClick={() => { setLeadIdPrompt(null); setLeadIdInputs({}) }}>
          <div className="bg-gray-900 border border-gray-700 rounded-xl p-6 w-96 shadow-xl" onClick={e => e.stopPropagation()}>
            <p className="text-sm font-semibold text-gray-100 mb-1">Add lead Slack ID{leadIdPrompt.length > 1 ? 's' : ''}</p>
            <p className="text-xs text-gray-500 mb-4">
              {leadIdPrompt.length > 1 ? 'Some leads' : 'This lead'} {leadIdPrompt.length > 1 ? "don't" : "doesn't"} have a Slack ID yet. Add {leadIdPrompt.length > 1 ? 'them' : 'it'} to tag them in the confirmation (Slack profile → ⋮ → Copy member ID). Saved to their staff profile.
            </p>
            <div className="space-y-2 mb-4">
              {leadIdPrompt.map((l, i) => (
                <div key={l.id} className="flex items-center gap-2">
                  <span className="text-xs text-gray-300 w-36 truncate">{l.name}</span>
                  <input
                    type="text"
                    autoFocus={i === 0}
                    className={inputSmClass + ' flex-1'}
                    placeholder="U0123ABCD"
                    value={leadIdInputs[l.id] ?? ''}
                    onChange={e => setLeadIdInputs(prev => ({ ...prev, [l.id]: e.target.value }))}
                    onKeyDown={e => e.key === 'Enter' && submitLeadIds(false)}
                  />
                </div>
              ))}
            </div>
            <div className="flex gap-2 justify-end">
              <button className={btnCancel} onClick={() => { setLeadIdPrompt(null); setLeadIdInputs({}) }}>Cancel</button>
              <button className={btnEdit} onClick={() => submitLeadIds(true)}>Skip &amp; preview</button>
              <button className={btnPrimary} style={{ backgroundColor: '#193a29' }} onClick={() => submitLeadIds(false)}>Save &amp; preview</button>
            </div>
          </div>
        </div>
      )}

      {leadPreview !== null && (
        <div className="fixed inset-0 bg-black/50 flex items-center justify-center z-50 p-6" onClick={() => setLeadPreview(null)}>
          <div className="bg-gray-900 border border-gray-700 rounded-xl p-6 w-[560px] max-w-full shadow-xl" onClick={e => e.stopPropagation()}>
            <p className="text-sm font-semibold text-gray-100 mb-1">Preview — Confirm assignments with leads</p>
            <p className="text-xs text-gray-500 mb-3">This will post to <span className="text-gray-300">#code_leads_plus</span>. Each person is @-mentioned in Slack (names shown here). Review before sending.</p>
            <pre className="text-xs text-gray-200 bg-gray-800/60 border border-gray-700 rounded-lg p-3 max-h-[50vh] overflow-y-auto whitespace-pre-wrap">{buildLeadText(leadPreview.leads, leadPreview.override, 'preview')}</pre>
            <div className="flex gap-2 justify-end mt-4">
              <button className={btnCancel} onClick={() => setLeadPreview(null)}>Cancel</button>
              <button className={btnPrimary} style={{ backgroundColor: '#193a29' }} onClick={sendLeadConfirmation}>Send to #code_leads_plus</button>
            </div>
          </div>
        </div>
      )}

      {projectCsvPrompt && (
        <div className="fixed inset-0 bg-black/50 flex items-center justify-center z-50 p-6" onClick={() => setProjectCsvPrompt(false)}>
          <div className="bg-gray-900 border border-gray-700 rounded-xl p-6 w-96 max-w-full shadow-xl" onClick={e => e.stopPropagation()}>
            <p className="text-sm font-semibold text-gray-100 mb-2">Export projects to CSV</p>
            <p className="text-xs text-gray-500 mb-5">Include completed projects in the export?</p>
            <div className="flex flex-col gap-2">
              <button className={btnPrimary} style={{ backgroundColor: '#193a29' }} onClick={() => exportProjectsCSV(true)}>Include completed</button>
              <button className="text-sm font-medium px-4 py-2 rounded-lg border border-gray-700 text-gray-300 hover:text-gray-100 hover:border-gray-600 transition-colors" onClick={() => exportProjectsCSV(false)}>Exclude completed</button>
              <button className={btnCancel + ' mt-1'} onClick={() => setProjectCsvPrompt(false)}>Cancel</button>
            </div>
          </div>
        </div>
      )}

      {confirmEmptyUpdate && (
        <div className="fixed inset-0 bg-black/50 flex items-center justify-center z-50" onClick={() => setConfirmEmptyUpdate(null)}>
          <div className="bg-gray-900 border border-gray-700 rounded-xl p-6 w-96 shadow-xl" onClick={e => e.stopPropagation()}>
            <p className="text-sm font-semibold text-gray-100 mb-2">Request update anyway?</p>
            <p className="text-xs text-gray-400 mb-5">
              <span className="text-gray-200">{confirmEmptyUpdate.name}</span> has <span className="text-amber-400 font-medium">no production plan and no open milestones</span> set. The thread will just ask for a general status update. Continue?
            </p>
            <div className="flex gap-2 justify-end">
              <button className={btnCancel} onClick={() => setConfirmEmptyUpdate(null)}>No, cancel</button>
              <button
                className={btnPrimary}
                style={{ backgroundColor: '#193a29' }}
                onClick={() => { const p = confirmEmptyUpdate; setConfirmEmptyUpdate(null); proceedRequest(p) }}
              >
                Yes, request update
              </button>
            </div>
          </div>
        </div>
      )}

      {pendingOrgRemove && (() => {
        const s = staff.find(x => x.id === pendingOrgRemove)
        const node = orgChart.find(n => n.staff_id === pendingOrgRemove)
        const isRoot = node ? !node.parent_staff_id : false
        const managerName = node?.parent_staff_id ? staff.find(x => x.id === node.parent_staff_id)?.name : null
        const kidCount = orgChart.filter(n => n.parent_staff_id === pendingOrgRemove).length
        return (
          <div className="fixed inset-0 bg-black/50 flex items-center justify-center z-50" onClick={() => setPendingOrgRemove(null)}>
            <div className="bg-gray-900 border border-gray-700 rounded-xl p-6 w-96 shadow-xl" onClick={e => e.stopPropagation()}>
              <p className="text-sm font-semibold text-gray-100 mb-2">Remove {s?.name} from the chart</p>
              <p className="text-xs text-gray-400 mb-5">
                {s?.name} has {kidCount} direct report{kidCount === 1 ? '' : 's'}. What should happen to {kidCount === 1 ? 'them' : 'those below'}?
              </p>
              <div className="flex flex-col gap-2">
                {!isRoot && (
                  <button
                    className="text-sm px-4 py-2 rounded-lg text-white font-medium hover:brightness-125 transition-colors"
                    style={{ backgroundColor: '#193a29' }}
                    onClick={() => { const id = pendingOrgRemove; setPendingOrgRemove(null); removeOrgReattach(id) }}
                  >
                    Remove only {s?.name} — move reports up{managerName ? ` to ${managerName}` : ''}
                  </button>
                )}
                <button
                  className="text-sm px-4 py-2 rounded-lg border border-red-500/40 text-red-400 hover:bg-red-500/10 transition-colors"
                  onClick={() => { const id = pendingOrgRemove; setPendingOrgRemove(null); removeFromOrg(id) }}
                >
                  Remove {s?.name} and everyone below
                </button>
                <button className={btnCancel + ' mt-1'} onClick={() => setPendingOrgRemove(null)}>Cancel</button>
              </div>
              {isRoot && <p className="text-[11px] text-gray-600 mt-3">This is the top of the chart, so reports can&apos;t move up — removing takes the whole chart with it.</p>}
            </div>
          </div>
        )
      })()}

      {confirmClearAll && (
        <div className="fixed inset-0 bg-black/50 flex items-center justify-center z-50" onClick={() => setConfirmClearAll(false)}>
          <div className="bg-gray-900 border border-gray-700 rounded-xl p-6 w-96 shadow-xl" onClick={e => e.stopPropagation()}>
            <p className="text-sm font-semibold text-gray-100 mb-2">Clear ALL assignments?</p>
            <p className="text-xs text-gray-400 mb-5">
              This permanently removes <span className="text-red-400 font-medium">all {assignments.length} assignment{assignments.length === 1 ? '' : 's'}</span> from <span className="text-gray-200">every project</span>, freeing up all staff. This can&apos;t be undone.
            </p>
            <div className="flex gap-2 justify-end">
              <button className={btnCancel} onClick={() => setConfirmClearAll(false)}>Cancel</button>
              <button className={btnPrimary} style={{ backgroundColor: '#193a29' }} onClick={clearAllAssignments}>Clear all</button>
            </div>
          </div>
        </div>
      )}

      {utilEdit && (() => {
        const e = utilEdit
        const set = (patch: Partial<typeof e>) => setUtilEdit({ ...e, ...patch })
        const free = !(e.assignment_name ?? '').trim() || (e.assignment_name ?? '').trim().toLowerCase() === 'free'
        const field = (label: string, el: React.ReactNode, hint?: string) => (
          <label className="block">
            <span className="text-[11px] text-gray-500 uppercase tracking-wider">{label}</span>
            <div className="mt-1">{el}</div>
            {hint && <span className="text-[11px] text-gray-600">{hint}</span>}
          </label>
        )
        const person = (k: 'lead' | 'supervisor' | 'sto' | 'ops', label: string) => field(label,
          <input className={inputClass + ' w-full'} list="util-people" value={e[k] ?? ''} placeholder="—" onChange={ev => set({ [k]: ev.target.value })} />)
        const cxSelect = (k: 'queue_complexity' | 'assignment_complexity', disabled = false) => (
          <select className={selectClass + ' w-full'} value={disabled ? '' : (e[k] ?? '')} disabled={disabled} onChange={ev => set({ [k]: ev.target.value || null })}>
            <option value="">—</option>
            {['Low', 'Mid', 'High'].map(o => <option key={o}>{o}</option>)}
          </select>
        )
        return (
          <div className="fixed inset-0 bg-black/50 flex items-center justify-center z-50 p-6" onClick={() => setUtilEdit(null)}>
            <div className="bg-gray-900 border border-gray-700 rounded-xl p-6 w-[640px] max-w-full max-h-[90vh] overflow-y-auto shadow-xl" onClick={ev => ev.stopPropagation()}>
              <p className="text-sm font-semibold text-gray-100 mb-4">{e.id ? 'Edit team' : 'Add team'}</p>

              <p className="text-xs font-semibold text-gray-400 mb-2">Configuration</p>
              <div className="grid grid-cols-2 gap-3 mb-3">
                {person('lead', 'Lead')}
                {person('supervisor', 'Supervisor')}
                {person('sto', 'STO')}
                {person('ops', 'Ops')}
              </div>
              <div className="grid grid-cols-2 gap-3 mb-5">
                {field('Team complexity', cxSelect('queue_complexity'), 'The most complex queue this team can run')}
              </div>

              <p className="text-xs font-semibold text-gray-400 mb-2">Current utilization</p>
              <div className="grid grid-cols-2 gap-3 mb-3">
                {field('Current assignment', (
                  <div className="flex gap-2">
                    <input className={inputClass + ' w-full'} value={free ? '' : (e.assignment_name ?? '')} placeholder="Free (no assignment)" onChange={ev => set({ assignment_name: ev.target.value || 'Free' })} />
                    {!free && <button type="button" onClick={() => set({ assignment_name: 'Free' })} className="text-xs px-2 rounded-lg border border-gray-700 text-gray-400 hover:text-white whitespace-nowrap">Make free</button>}
                  </div>
                ))}
                {field('Queue complexity', cxSelect('assignment_complexity', free))}
              </div>
              <div className="mb-3">
                {field('Utilization score', (
                  <div className="grid grid-cols-4 gap-2">
                    {UTIL_TIERS.map(k => {
                      const on = free ? k === 'C' : e.tier === k
                      return (
                        <button
                          key={k}
                          type="button"
                          disabled={free}
                          onClick={() => set({ tier: e.tier === k ? null : k })}
                          title={UTIL_TIER_INFO[k].detail}
                          className={`rounded-lg border px-2 py-2 text-left transition-colors ${on ? `${UTIL_TIER_INFO[k].cls} border-transparent` : 'border-gray-700 text-gray-400 hover:border-gray-500'} ${free ? 'opacity-60 cursor-not-allowed' : ''}`}
                        >
                          <p className="text-sm font-bold">{k}</p>
                          <p className="text-[10px] leading-tight opacity-80">{UTIL_TIER_INFO[k].detail}</p>
                        </button>
                      )
                    })}
                  </div>
                ), free ? 'Free teams count as C (not assigned).' : undefined)}
              </div>
              <div className="grid grid-cols-2 gap-3 mb-3">
                {field('Priority to assign', (
                  <select className={selectClass + ' w-full'} value={free ? (e.assign_priority ?? '') : ''} disabled={!free} onChange={ev => set({ assign_priority: ev.target.value ? Number(ev.target.value) : null })}>
                    <option value="">—</option>
                    {[1, 2, 3].map(n => <option key={n} value={n}>P{n}</option>)}
                  </select>
                ), free ? '1 = staff this team first' : 'Only for free teams')}
                {field('Why (utilization status)', (
                  <input className={inputClass + ' w-full'} value={e.reason ?? ''} placeholder="Leave blank for the automatic explanation" onChange={ev => set({ reason: ev.target.value })} />
                ))}
              </div>
              <div className="mb-5">
                {field('Notes', <textarea className={inputClass + ' w-full'} rows={2} value={e.notes ?? ''} onChange={ev => set({ notes: ev.target.value })} />)}
              </div>

              <div className="flex items-center gap-2">
                {e.id && <button className={btnDanger} onClick={() => deleteUtilTeam(e)}>Remove team</button>}
                <div className="flex gap-2 ml-auto">
                  <button className={btnCancel} onClick={() => setUtilEdit(null)}>Cancel</button>
                  <button className={btnPrimary} style={{ backgroundColor: '#193a29' }} onClick={saveUtilEdit}>{e.id ? 'Save' : 'Add team'}</button>
                </div>
              </div>
            </div>
          </div>
        )
      })()}

      {manageTypesOpen && (
        <div className="fixed inset-0 bg-black/50 flex items-center justify-center z-50 p-6" onClick={() => setManageTypesOpen(false)}>
          <div className="bg-gray-900 border border-gray-700 rounded-xl p-6 w-[420px] max-w-full shadow-xl" onClick={e => e.stopPropagation()}>
            <p className="text-sm font-semibold text-gray-100 mb-1">Project types</p>
            <p className="text-xs text-gray-500 mb-4">The values available in the Project Type dropdown. Rename inline (Enter to save) — projects using a type update with it.</p>
            <div className="space-y-2 max-h-[45vh] overflow-y-auto mb-4">
              {projectTypeOptions().length === 0 && <p className="text-xs text-gray-600">No types yet — add one below.</p>}
              {projectTypeOptions().map(t => {
                const used = projects.filter(p => p.project_type === t).length
                return (
                  <div key={t} className="flex items-center gap-2">
                    <input
                      className={inputSmClass + ' flex-1'}
                      defaultValue={t}
                      onBlur={e => renameProjectType(t, e.target.value)}
                      onKeyDown={e => { if (e.key === 'Enter') (e.target as HTMLInputElement).blur() }}
                    />
                    <span className="text-xs text-gray-500 w-20 text-right whitespace-nowrap">{used} project{used === 1 ? '' : 's'}</span>
                    <button className={btnDanger} onClick={() => deleteProjectType(t)} title="Delete type">×</button>
                  </div>
                )
              })}
            </div>
            <div className="flex gap-2 pt-3 border-t border-gray-800">
              <input
                className={inputSmClass + ' flex-1'}
                placeholder="New type, e.g. RLHF, Evals, SFT…"
                value={newTypeName}
                onChange={e => setNewTypeName(e.target.value)}
                onKeyDown={async e => { if (e.key === 'Enter' && newTypeName.trim()) { await addProjectType(newTypeName); setNewTypeName('') } }}
              />
              <button
                className="text-xs px-3 py-1 rounded bg-gray-700 text-gray-200 hover:text-white transition-colors"
                onClick={async () => { if (newTypeName.trim()) { await addProjectType(newTypeName); setNewTypeName('') } }}
              >Add</button>
            </div>
            <div className="flex justify-end mt-4">
              <button className={btnCancel} onClick={() => setManageTypesOpen(false)}>Done</button>
            </div>
          </div>
        </div>
      )}

      {clearMsPrompt && (() => {
        const activeIds = new Set(projects.filter(p => p.status === 'active').map(p => p.id))
        const msCount = milestones.filter(m => activeIds.has(m.project_id)).length
        const tgtCount = projects.filter(p => activeIds.has(p.id) && p.weekly_target != null).length
        return (
          <div className="fixed inset-0 bg-black/50 flex items-center justify-center z-50 p-6" onClick={() => setClearMsPrompt(false)}>
            <div className="bg-gray-900 border border-gray-700 rounded-xl p-6 w-96 max-w-full shadow-xl" onClick={e => e.stopPropagation()}>
              <p className="text-sm font-semibold text-gray-100 mb-2">Clear milestones &amp; targets</p>
              <p className="text-xs text-gray-400 mb-5">
                Across active projects: <span className="text-gray-200 font-medium">{msCount}</span> milestone{msCount === 1 ? '' : 's'} and <span className="text-gray-200 font-medium">{tgtCount}</span> production-plan target{tgtCount === 1 ? '' : 's'}. This can&apos;t be undone.
              </p>
              <div className="flex flex-col gap-2">
                <button className="text-sm font-medium px-4 py-2 rounded-lg border border-gray-700 text-gray-300 hover:text-gray-100 hover:border-gray-600 transition-colors" disabled={msCount === 0} onClick={() => clearMilestonesAndTargets({ milestones: true, targets: false })}>Clear milestones only</button>
                <button className="text-sm font-medium px-4 py-2 rounded-lg border border-gray-700 text-gray-300 hover:text-gray-100 hover:border-gray-600 transition-colors" disabled={tgtCount === 0} onClick={() => clearMilestonesAndTargets({ milestones: false, targets: true })}>Clear targets only</button>
                <button className={btnPrimary} style={{ backgroundColor: '#7f1d1d' }} disabled={msCount === 0 && tgtCount === 0} onClick={() => clearMilestonesAndTargets({ milestones: true, targets: true })}>Clear both</button>
                <button className={btnCancel + ' mt-1'} onClick={() => setClearMsPrompt(false)}>Cancel</button>
              </div>
            </div>
          </div>
        )
      })()}

      {confirmClearInactive && (
        <div className="fixed inset-0 bg-black/50 flex items-center justify-center z-50" onClick={() => setConfirmClearInactive(false)}>
          <div className="bg-gray-900 border border-gray-700 rounded-xl p-6 w-96 shadow-xl" onClick={e => e.stopPropagation()}>
            <p className="text-sm font-semibold text-gray-100 mb-2">Clear staffing from inactive projects?</p>
            <p className="text-xs text-gray-400 mb-5">
              This removes <span className="text-red-400 font-medium">{inactiveAssignmentIds().length}</span> assignment{inactiveAssignmentIds().length === 1 ? '' : 's'} from all <span className="text-gray-200">paused and completed</span> projects, freeing up that staff. This can&apos;t be undone.
            </p>
            <div className="flex gap-2 justify-end">
              <button className={btnCancel} onClick={() => setConfirmClearInactive(false)}>Cancel</button>
              <button className={btnPrimary} style={{ backgroundColor: '#193a29' }} onClick={clearInactiveStaffing}>Clear staffing</button>
            </div>
          </div>
        </div>
      )}

      {pendingFlex && (() => {
        const member = staff.find(s => s.id === pendingFlex.staffId)
        return (
          <div className="fixed inset-0 bg-black/50 flex items-center justify-center z-50" onClick={() => { setPendingFlex(null); setFlexNote('') }}>
            <div className="bg-gray-900 border border-gray-700 rounded-xl p-6 w-80 shadow-xl" onClick={e => e.stopPropagation()}>
              <p className="text-sm font-semibold text-gray-100 mb-1">Flex {member?.name}</p>
              <p className="text-xs text-gray-500 mb-4">Add a note about this flex (optional):</p>
              <input
                type="text"
                className={inputClass + ' w-full mb-4'}
                placeholder="e.g. Flexed to MM until Q3"
                value={flexNote}
                autoFocus
                onChange={e => setFlexNote(e.target.value)}
                onKeyDown={async e => {
                  if (e.key === 'Enter') {
                    const { data } = await supabase.from('staff').update({ flexed: true, ooo: false, onboarding: false, flex_notes: flexNote.trim() || null }).eq('id', pendingFlex.staffId).select().single()
                    if (data) setStaff(prev => prev.map(s => s.id === pendingFlex.staffId ? data : s))
                    if (pendingFlex.assignmentId) await removeAssignment(pendingFlex.assignmentId)
                    setPendingFlex(null); setFlexNote('')
                  }
                }}
              />
              <div className="flex gap-2 justify-end">
                <button className={btnCancel} onClick={() => { setPendingFlex(null); setFlexNote('') }}>Cancel</button>
                <button
                  className={btnPrimary}
                  style={{ backgroundColor: '#193a29' }}
                  onClick={async () => {
                    const { data } = await supabase.from('staff').update({ flexed: true, ooo: false, onboarding: false, flex_notes: flexNote.trim() || null }).eq('id', pendingFlex.staffId).select().single()
                    if (data) setStaff(prev => prev.map(s => s.id === pendingFlex.staffId ? data : s))
                    if (pendingFlex.assignmentId) await removeAssignment(pendingFlex.assignmentId)
                    setPendingFlex(null); setFlexNote('')
                  }}
                >
                  Flex
                </button>
              </div>
            </div>
          </div>
        )
      })()}

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
              <option>Lead</option>
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
