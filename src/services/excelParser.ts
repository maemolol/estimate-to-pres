import * as XLSX from 'xlsx'

// ─── Types ────────────────────────────────────────────────────────────────────

export type Role = 'FED' | 'BED' | 'Arch' | 'QA/BA' | 'PM' | 'Gov'

export interface RoleHours {
  role: Role
  low: number
  high: number
  rate: number
  costLow: number
  costHigh: number
}

export interface TaskRow {
  task: string
  section: string          // e.g. "Onboarding", "Implementation > Content and Layout"
  roles: RoleHours[]
  totalHoursLow: number
  totalHoursHigh: number
  totalCostLow: number
  totalCostHigh: number
  comments: string
}

export interface SheetSummary {
  sheetName: string
  tasks: TaskRow[]
  sections: string[]
  totalHoursLow: number
  totalHoursHigh: number
  totalCostLow: number
  totalCostHigh: number
  midCost: number          // mid-point estimate
  roles: { role: Role; hoursLow: number; hoursHigh: number; costLow: number; costHigh: number }[]
}

export interface ParseResult {
  sheets: SheetSummary[]
  primary: SheetSummary   // "Reduced" or first substantive sheet
  warnings: string[]
}

// ─── Constants ───────────────────────────────────────────────────────────────

const ROLE_COLUMNS: { role: Role; lowIdx: number; highIdx: number }[] = [
  { role: 'FED',   lowIdx: 1, highIdx: 2  },
  { role: 'BED',   lowIdx: 3, highIdx: 4  },
  { role: 'Arch',  lowIdx: 5, highIdx: 6  },
  { role: 'QA/BA', lowIdx: 7, highIdx: 8  },
  { role: 'PM',    lowIdx: 11, highIdx: 12 },
  { role: 'Gov',   lowIdx: 13, highIdx: 14 },
]

const SKIP_TASKS = new Set([
  '', 'task', 'subtotal', 'agile', 'total hours', 'total w/buffer',
  'weeks', 'rate', 'cost', 'assumptions', 'nan', 'implementation',
  'content and layout', 'forms', 'search migration (solr to sitecore search)',
  'security-related', 'miscellaneous', 'compliance and accessibility',
  'documentation', 'training and handoff sessions', 'deployments',
  'onboarding',
])

const SECTION_HEADERS = new Set([
  'onboarding', 'implementation', 'content and layout', 'forms',
  'search migration (solr to sitecore search)', 'security-related',
  'miscellaneous', 'compliance and accessibility', 'documentation',
  'training and handoff sessions', 'deployments',
])

const PREFERRED_SHEETS = ['reduced', 'lt rates reduced', 'original']

// ─── Helpers ─────────────────────────────────────────────────────────────────

function num(v: unknown): number {
  if (v === null || v === undefined || v === '' || (typeof v === 'string' && v.toLowerCase() === 'nan')) return 0
  const n = Number(v)
  return isNaN(n) ? 0 : n
}

function isDataRow(row: unknown[]): boolean {
  const task = String(row[0] ?? '').trim().toLowerCase()
  if (!task || task === 'nan') return false
  if (SKIP_TASKS.has(task)) return false
  // Must have at least one numeric value across role columns
  const hasNumbers = ROLE_COLUMNS.some(({ lowIdx, highIdx }) =>
    num(row[lowIdx]) > 0 || num(row[highIdx]) > 0
  )
  return hasNumbers
}

function isSectionHeader(val: unknown): boolean {
  const s = String(val ?? '').trim().toLowerCase()
  return SECTION_HEADERS.has(s)
}

// ─── Per-sheet parser ─────────────────────────────────────────────────────────

function parseSheet(ws: XLSX.WorkSheet, sheetName: string, warnings: string[]): SheetSummary | null {
  const raw: unknown[][] = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null })
  if (!raw.length) return null

  // Find header row (contains "Task" and role names)
  let headerRowIdx = -1
  for (let i = 0; i < Math.min(raw.length, 5); i++) {
    const first = String(raw[i][0] ?? '').trim().toLowerCase()
    if (first === 'task') { headerRowIdx = i; break }
  }
  if (headerRowIdx < 0) return null

  // Find rate row to compute per-role rates
  const rates: Record<Role, number> = { FED: 100, BED: 100, Arch: 150, 'QA/BA': 95, PM: 125, Gov: 240 }
  for (let i = headerRowIdx + 1; i < raw.length; i++) {
    const label = String(raw[i][0] ?? '').trim().toLowerCase()
    if (label === 'rate') {
      const row = raw[i]
      ROLE_COLUMNS.forEach(({ role, lowIdx }) => {
        const r = num(row[lowIdx])
        if (r > 0) rates[role] = r
      })
      break
    }
  }

  const tasks: TaskRow[] = []
  let currentSection = 'General'
  let currentSubSection = ''

  for (let i = headerRowIdx + 1; i < raw.length; i++) {
    const row = raw[i]
    const taskLabel = String(row[0] ?? '').trim()
    if (!taskLabel || taskLabel.toLowerCase() === 'nan') continue

    const taskLower = taskLabel.toLowerCase()

    // Skip footer/summary rows
    if (['subtotal','agile','total hours','total w/buffer','weeks','rate','cost','assumptions'].includes(taskLower)) break

    // Section header detection
    if (isSectionHeader(taskLabel)) {
      currentSection = taskLabel
      currentSubSection = ''
      continue
    }

    // Sub-section headers (rows with a label but all-zero numeric columns, no numbers)
    const hasAnyNum = ROLE_COLUMNS.some(({ lowIdx, highIdx }) => num(row[lowIdx]) > 0 || num(row[highIdx]) > 0)
    if (!hasAnyNum) {
      // Could be a sub-section label
      if (taskLabel.length > 2 && taskLabel.length < 60) {
        currentSubSection = taskLabel
      }
      continue
    }

    // Build role hours
    const roleHours: RoleHours[] = []
    ROLE_COLUMNS.forEach(({ role, lowIdx, highIdx }) => {
      const low = num(row[lowIdx])
      const high = num(row[highIdx])
      if (low > 0 || high > 0) {
        roleHours.push({
          role, low, high,
          rate: rates[role],
          costLow: Math.round(low * rates[role]),
          costHigh: Math.round(high * rates[role]),
        })
      }
    })

    const totalHoursLow = roleHours.reduce((s, r) => s + r.low, 0)
    const totalHoursHigh = roleHours.reduce((s, r) => s + r.high, 0)
    const totalCostLow = roleHours.reduce((s, r) => s + r.costLow, 0)
    const totalCostHigh = roleHours.reduce((s, r) => s + r.costHigh, 0)

    const sectionPath = currentSubSection
      ? `${currentSection} › ${currentSubSection}`
      : currentSection

    const comments = String(row[15] ?? row[13] ?? '').trim()
    if (comments.toLowerCase() === 'nan') {
      // ignore
    }

    tasks.push({
      task: taskLabel,
      section: sectionPath,
      roles: roleHours,
      totalHoursLow: Math.round(totalHoursLow * 10) / 10,
      totalHoursHigh: Math.round(totalHoursHigh * 10) / 10,
      totalCostLow,
      totalCostHigh,
      comments: comments === 'nan' ? '' : comments,
    })
  }

  if (!tasks.length) return null

  // Aggregate totals
  const totalHoursLow = tasks.reduce((s, t) => s + t.totalHoursLow, 0)
  const totalHoursHigh = tasks.reduce((s, t) => s + t.totalHoursHigh, 0)
  const totalCostLow = tasks.reduce((s, t) => s + t.totalCostLow, 0)
  const totalCostHigh = tasks.reduce((s, t) => s + t.totalCostHigh, 0)
  const midCost = Math.round((totalCostLow + totalCostHigh) / 2)

  // Unique sections
  const sections = [...new Set(tasks.map((t) => t.section.split(' › ')[0]))]

  // Role summaries
  const roleTotals = new Map<Role, { hoursLow: number; hoursHigh: number; costLow: number; costHigh: number }>()
  tasks.forEach((t) => {
    t.roles.forEach((r) => {
      const existing = roleTotals.get(r.role) ?? { hoursLow: 0, hoursHigh: 0, costLow: 0, costHigh: 0 }
      roleTotals.set(r.role, {
        hoursLow: existing.hoursLow + r.low,
        hoursHigh: existing.hoursHigh + r.high,
        costLow: existing.costLow + r.costLow,
        costHigh: existing.costHigh + r.costHigh,
      })
    })
  })

  const roleOrder: Role[] = ['FED', 'BED', 'Arch', 'QA/BA', 'PM', 'Gov']
  const roles = roleOrder
    .filter((r) => roleTotals.has(r))
    .map((r) => ({ role: r, ...roleTotals.get(r)! }))

  return {
    sheetName,
    tasks,
    sections,
    totalHoursLow: Math.round(totalHoursLow),
    totalHoursHigh: Math.round(totalHoursHigh),
    totalCostLow,
    totalCostHigh,
    midCost,
    roles,
  }
}

// ─── Main entry ───────────────────────────────────────────────────────────────

export function parseEstimateFile(file: File): Promise<ParseResult> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader()
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target!.result as ArrayBuffer)
        const wb = XLSX.read(data, { type: 'array' })

        const warnings: string[] = []
        const sheets: SheetSummary[] = []

        // Parse each sheet
        for (const name of wb.SheetNames) {
          const ws = wb.Sheets[name]
          const result = parseSheet(ws, name, warnings)
          if (result) sheets.push(result)
        }

        if (!sheets.length) {
          reject(new Error(
            'Could not find any estimate data. ' +
            'Expected sheets with a "Task" header row and role columns (FED, BED, Arch, QA/BA, PM).'
          ))
          return
        }

        // Pick the primary sheet
        const nameLower = sheets.map((s) => s.sheetName.toLowerCase())
        let primaryIdx = 0
        for (const preferred of PREFERRED_SHEETS) {
          const idx = nameLower.indexOf(preferred)
          if (idx >= 0) { primaryIdx = idx; break }
        }

        resolve({ sheets, primary: sheets[primaryIdx], warnings })
      } catch (err) {
        reject(new Error(`Failed to parse file: ${(err as Error).message}`))
      }
    }
    reader.onerror = () => reject(new Error('Failed to read file.'))
    reader.readAsArrayBuffer(file)
  })
}

// ─── Demo data (the Lattice estimate, pre-parsed for instant preview) ─────────

export function getLatticeDemo(): ParseResult {
  const tasks: TaskRow[] = [
    { task: 'Project setup', section: 'Onboarding', roles: [{role:'Arch',low:1,high:1,rate:130,costLow:130,costHigh:130},{role:'QA/BA',low:4,high:4,rate:75,costLow:300,costHigh:300},{role:'PM',low:4,high:6,rate:100,costLow:400,costHigh:600}], totalHoursLow:9, totalHoursHigh:11, totalCostLow:830, totalCostHigh:1030, comments:'' },
    { task: 'Discovery workshops', section: 'Onboarding', roles: [{role:'Arch',low:1,high:2,rate:130,costLow:130,costHigh:260},{role:'QA/BA',low:8,high:12,rate:75,costLow:600,costHigh:900},{role:'PM',low:2,high:2,rate:100,costLow:200,costHigh:200}], totalHoursLow:11, totalHoursHigh:16, totalCostLow:930, totalCostHigh:1360, comments:'' },
    { task: 'Local Dev Setup', section: 'Onboarding', roles: [{role:'FED',low:2,high:6,rate:80,costLow:160,costHigh:480},{role:'BED',low:8,high:12,rate:80,costLow:640,costHigh:960},{role:'Arch',low:1,high:1,rate:130,costLow:130,costHigh:130},{role:'PM',low:1.65,high:2.25,rate:100,costLow:165,costHigh:225}], totalHoursLow:13, totalHoursHigh:21, totalCostLow:1095, totalCostHigh:1795, comments:'' },
    { task: 'Functional Spec', section: 'Onboarding', roles: [{role:'Arch',low:2,high:3,rate:130,costLow:260,costHigh:390},{role:'QA/BA',low:32,high:40,rate:75,costLow:2400,costHigh:3000},{role:'PM',low:5.1,high:6.45,rate:100,costLow:510,costHigh:645}], totalHoursLow:39, totalHoursHigh:49, totalCostLow:3170, totalCostHigh:4035, comments:'' },
    { task: 'Content Migration (11,700+ pages)', section: 'Implementation › Content and Layout', roles: [{role:'Arch',low:8,high:12,rate:130,costLow:1040,costHigh:1560}], totalHoursLow:8, totalHoursHigh:12, totalCostLow:1040, totalCostHigh:1560, comments:'' },
    { task: 'Media Migration', section: 'Implementation › Content and Layout', roles: [{role:'Arch',low:3,high:6,rate:130,costLow:390,costHigh:780}], totalHoursLow:3, totalHoursHigh:6, totalCostLow:390, totalCostHigh:780, comments:'' },
    { task: 'Component conversion (112 components)', section: 'Implementation › Content and Layout', roles: [{role:'FED',low:200,high:350,rate:80,costLow:16000,costHigh:28000},{role:'Arch',low:2,high:4,rate:130,costLow:260,costHigh:520},{role:'QA/BA',low:50,high:87.5,rate:75,costLow:3750,costHigh:6563}], totalHoursLow:252, totalHoursHigh:442, totalCostLow:20010, totalCostHigh:35083, comments:'112 total: 50 simple, 30 medium, 20 hard, 12 search' },
    { task: 'Forms analysis & recreation', section: 'Implementation › Forms', roles: [{role:'BED',low:14,high:24,rate:80,costLow:1120,costHigh:1920},{role:'Arch',low:1.75,high:3,rate:130,costLow:228,costHigh:390},{role:'QA/BA',low:3.5,high:6,rate:75,costLow:263,costHigh:450},{role:'PM',low:2.89,high:4.95,rate:100,costLow:289,costHigh:495}], totalHoursLow:22, totalHoursHigh:38, totalCostLow:1900, totalCostHigh:3255, comments:'~10 forms' },
    { task: 'Sitecore Search setup & migration', section: 'Implementation › Search', roles: [{role:'BED',low:16,high:26,rate:80,costLow:1280,costHigh:2080},{role:'Arch',low:3,high:5.5,rate:130,costLow:390,costHigh:715},{role:'QA/BA',low:4,high:7.5,rate:75,costLow:300,costHigh:563},{role:'PM',low:3.3,high:5.4,rate:100,costLow:330,costHigh:540}], totalHoursLow:26, totalHoursHigh:44, totalCostLow:2300, totalCostHigh:3898, comments:'Solr to Sitecore Search' },
    { task: 'Azure B2C & SSO Enablement', section: 'Implementation › Security', roles: [{role:'FED',low:6,high:10,rate:80,costLow:480,costHigh:800},{role:'BED',low:12,high:20,rate:80,costLow:960,costHigh:1600},{role:'Arch',low:2,high:3.5,rate:130,costLow:260,costHigh:455},{role:'QA/BA',low:4,high:7,rate:75,costLow:300,costHigh:525}], totalHoursLow:24, totalHoursHigh:41, totalCostLow:2000, totalCostHigh:3380, comments:'OIDC/Entra ID' },
    { task: 'Next.js integrations (GTM, GetResponse, Zoho, Agiloft)', section: 'Implementation › Miscellaneous', roles: [{role:'FED',low:40,high:56,rate:80,costLow:3200,costHigh:4480},{role:'Arch',low:5,high:7,rate:130,costLow:650,costHigh:910},{role:'QA/BA',low:10,high:14,rate:75,costLow:750,costHigh:1050},{role:'PM',low:8.25,high:9.15,rate:100,costLow:825,costHigh:915}], totalHoursLow:63, totalHoursHigh:86, totalCostLow:5425, totalCostHigh:7355, comments:'' },
    { task: 'ADA 508 Compliance carry-over', section: 'Compliance and Accessibility', roles: [{role:'FED',low:8,high:10,rate:80,costLow:640,costHigh:800},{role:'Arch',low:1,high:1.25,rate:130,costLow:130,costHigh:163}], totalHoursLow:9, totalHoursHigh:11, totalCostLow:770, totalCostHigh:963, comments:'' },
    { task: 'Regression & UAT', section: 'Testing', roles: [{role:'FED',low:40,high:60,rate:80,costLow:3200,costHigh:4800},{role:'QA/BA',low:18,high:31,rate:75,costLow:1350,costHigh:2325},{role:'PM',low:9.45,high:11.78,rate:100,costLow:945,costHigh:1178}], totalHoursLow:67, totalHoursHigh:103, totalCostLow:5495, totalCostHigh:8303, comments:'2 weeks UAT' },
    { task: 'Technical Design Document', section: 'Documentation', roles: [{role:'Arch',low:8,high:12,rate:130,costLow:1040,costHigh:1560},{role:'PM',low:1.2,high:1.95,rate:100,costLow:120,costHigh:195}], totalHoursLow:9, totalHoursHigh:14, totalCostLow:1160, totalCostHigh:1755, comments:'' },
    { task: 'Content author training & handoff', section: 'Documentation', roles: [{role:'Arch',low:3,high:3,rate:130,costLow:390,costHigh:390},{role:'PM',low:0.45,high:0.45,rate:100,costLow:45,costHigh:45}], totalHoursLow:3.5, totalHoursHigh:3.5, totalCostLow:435, totalCostHigh:435, comments:'prep + 1hr session' },
    { task: 'DEV / QA / PROD Deployments', section: 'Deployments', roles: [{role:'BED',low:8,high:12,rate:80,costLow:640,costHigh:960},{role:'Arch',low:1,high:1.5,rate:130,costLow:130,costHigh:195},{role:'QA/BA',low:2,high:3,rate:75,costLow:150,costHigh:225},{role:'PM',low:1.65,high:2.475,rate:100,costLow:165,costHigh:248}], totalHoursLow:13, totalHoursHigh:19, totalCostLow:1085, totalCostHigh:1628, comments:'4 DEV, 4 QA, 1 PROD' },
  ]

  const totalCostLow = tasks.reduce((s, t) => s + t.totalCostLow, 0)
  const totalCostHigh = tasks.reduce((s, t) => s + t.totalCostHigh, 0)
  const totalHoursLow = tasks.reduce((s, t) => s + t.totalHoursLow, 0)
  const totalHoursHigh = tasks.reduce((s, t) => s + t.totalHoursHigh, 0)

  const sheet: SheetSummary = {
    sheetName: 'LT Rates Reduced',
    tasks,
    sections: ['Onboarding', 'Implementation', 'Compliance and Accessibility', 'Testing', 'Documentation', 'Deployments'],
    totalHoursLow: Math.round(totalHoursLow),
    totalHoursHigh: Math.round(totalHoursHigh),
    totalCostLow,
    totalCostHigh,
    midCost: Math.round((totalCostLow + totalCostHigh) / 2),
    roles: [
      { role: 'FED',   hoursLow: 296, hoursHigh: 492, costLow: 23680, costHigh: 39360 },
      { role: 'BED',   hoursLow: 58,  hoursHigh: 94,  costLow: 4640,  costHigh: 7520  },
      { role: 'Arch',  hoursLow: 41,  hoursHigh: 61,  costLow: 5330,  costHigh: 7930  },
      { role: 'QA/BA', hoursLow: 136, hoursHigh: 215, costLow: 10200, costHigh: 16125 },
      { role: 'PM',    hoursLow: 40,  hoursHigh: 55,  costLow: 4000,  costHigh: 5500  },
    ],
  }

  return { sheets: [sheet], primary: sheet, warnings: [] }
}
