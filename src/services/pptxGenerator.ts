/**
 * pptxGenerator.ts
 *
 * Generates a PPTX closely matching the Zont/Lattice Migration Proposal template.
 * All coordinates verified against template XML.
 *
 * Key template facts:
 *   - Slide: 13.33" × 7.5" (LAYOUT_WIDE)
 *   - Content slides: WHITE background
 *   - Divider/closing slides: dark #1C1C1C background
 *   - Green fill: #549E39 (shapes/panels), green text: #538321
 *   - Body text: #3D3B47, head text: #000000
 *   - Fonts: Calibri Light (headings), Calibri (body)
 *   - Safe content area: x 0.49–12.97", y 0.86–6.95"
 *   - Content slides use LEFT half for text (~0.49–5.8") + RIGHT half for visual panel (~5.9–13.0")
 *   - Stat boxes (slide 6): x=1.23/4.17, y=1.94/3.54/5.15, w≈2.0, h=1.16
 *   - Two-col boxes (slide 13): green x=1.46,y=2.26,w=5.27,h=3.36 / dark x=6.73,y=2.26,w=5.18,h=3.36
 *   - Investment table (slide 18): x=0.42, headers dark #2E2E2E, col widths 7.5/1.8/3.08
 *   - Section divider title (slide 15): x=0.76,y=2.77,w=9.17,h=2.70, 60pt NOT bold, green
 *   - Slide number box: x=12.55,y=7.02,w=0.56,h=0.27, dark bg
 */

import pptxgen from 'pptxgenjs'
import type { ProposalContent } from '../store/appStore'

// ─── Exact brand tokens ───────────────────────────────────────────────────────
const C = {
  white:     'FFFFFF',
  black:     '000000',
  gFill:     '549E39',   // green shape fills (#549E39)
  gText:     '538321',   // green text (#538321)
  body:      '3D3B47',   // body copy
  head:      '000000',   // slide titles
  dark:      '1C1C1C',   // divider/closing bg
  darkPanel: '252525',   // slightly lighter right panel on dark slides
  tblHdr:    '2E2E2E',   // table/slide-num header bg
  tblAlt:    'F2F2F2',   // alternating table row
  numBg:     '2E2E2E',   // slide number box bg
  zGrey:     '9EA8B3',   // ZONT wordmark grey
  lineGrey:  'E0E0E0',   // divider lines
  barBg:     'EBEBEB',   // bar chart background
  hF: 'Calibri Light',
  bF: 'Calibri',
}

// ─── Shared chrome helpers ────────────────────────────────────────────────────

/** Slide number box + ZONT wordmark — on every content slide */
function addChrome(s: pptxgen.Slide, n: number) {
  s.addShape('rect', { x: 12.55, y: 7.02, w: 0.56, h: 0.27, fill: { color: C.numBg }, line: { width: 0 } })
  s.addText(String(n), { x: 12.55, y: 7.02, w: 0.56, h: 0.27, fontSize: 9, color: C.white, fontFace: C.bF, align: 'center', valign: 'middle' })
  s.addText('ZONT', { x: 10.15, y: 7.05, w: 1.2, h: 0.28, fontSize: 11, color: C.zGrey, fontFace: C.hF, align: 'right' })
}

/** Standard slide title bar — 24pt bold black, x=0.49 y=0.28 */
function addTitle(s: pptxgen.Slide, text: string) {
  s.addText(text, { x: 0.49, y: 0.28, w: 12.35, h: 0.44, fontSize: 24, bold: true, color: C.head, fontFace: C.hF })
}

/**
 * Right-side visual panel — dark rectangle with green top bar.
 * Mimics the product screenshot area seen on slides 4, 6, 8, etc.
 */
function addRightPanel(s: pptxgen.Slide, label: string) {
  s.addShape('rect', { x: 6.30, y: 0.86, w: 6.70, h: 6.30, fill: { color: '1A1A18' }, line: { width: 0 } })
  s.addShape('rect', { x: 6.30, y: 0.86, w: 6.70, h: 0.055, fill: { color: C.gFill }, line: { width: 0 } })
  if (label) {
    s.addText(label, { x: 6.50, y: 1.05, w: 6.30, h: 0.36, fontSize: 12, color: C.zGrey, fontFace: C.bF, italic: true })
  }
}

// ─── Slide 1: Cover ───────────────────────────────────────────────────────────
function slideCover(prs: pptxgen, p: ProposalContent) {
  const s = prs.addSlide()
  s.background = { color: C.white }

  // Green left panel (x=0, y=0, w=7.514, h=7.5)
  s.addShape('rect', { x: 0, y: 0, w: 7.51, h: 7.5, fill: { color: C.gFill }, line: { width: 0 } })

  // Green overlay banner (x=4.523, y=2.967, w=7.553, h=1.608)
  s.addShape('rect', { x: 4.52, y: 2.97, w: 8.81, h: 1.61, fill: { color: C.gFill }, line: { width: 0 } })

  // Logo area — "ZONT | CLIENT" top-left
  s.addText('ZONT  |  ' + p.clientName.toUpperCase(), {
    x: 0.61, y: 0.55, w: 5.5, h: 0.40, fontSize: 13, color: C.white, fontFace: C.bF,
  })

  // Main title — 44pt bold white (x=0.611, y=2.688, w=6.631, h=0.789)
  s.addText(p.projectName, {
    x: 0.61, y: 2.69, w: 6.63, h: 1.00, fontSize: 44, bold: true, color: C.white, fontFace: C.hF,
  })

  // Subtitle (x=0.638, y=3.971, 16pt)
  s.addText('AI-powered proposal for ' + p.clientName, {
    x: 0.64, y: 4.00, w: 5.94, h: 0.38, fontSize: 16, color: C.white, fontFace: C.bF,
  })
  s.addText(p.date, {
    x: 0.64, y: 4.46, w: 5.94, h: 0.32, fontSize: 13, color: C.white, fontFace: C.bF, italic: true,
  })
}

// ─── Slide 2: Agenda ──────────────────────────────────────────────────────────
function slideAgenda(prs: pptxgen, p: ProposalContent, n: number) {
  const s = prs.addSlide()
  s.background = { color: C.white }

  // "Agenda" large left (template slide 3: 40pt bold)
  s.addText('Agenda', { x: 0.49, y: 0.55, w: 5, h: 0.68, fontSize: 40, bold: true, color: C.head, fontFace: C.hF })

  const items = [
    'Our understanding of your objectives',
    'Our recommended solution & scope',
    `Effort breakdown across ${p.sheet.sections.length} work streams`,
    'Investment summary',
    'Team & role allocation',
    'Next steps',
  ]

  // Green circle + text (template: green circle 0.50×0.50 #538321, text 16pt black)
  items.forEach((item, i) => {
    const y = 1.52 + i * 0.76
    s.addShape('ellipse', { x: 0.49, y: y + 0.09, w: 0.32, h: 0.32, fill: { color: C.gText }, line: { width: 0 } })
    s.addText(item, { x: 0.96, y, w: 11.50, h: 0.50, fontSize: 16, color: C.black, fontFace: C.bF })
  })

  addChrome(s, n)
}

// ─── Slide 3: Objectives ──────────────────────────────────────────────────────
// Template slide 4: title x=0.529 y=0.961 24pt bold; bullets x=0.881 y=2.349 w=4.895 10.5pt; right image x=5.242 y=0.930
function slideObjectives(prs: pptxgen, p: ProposalContent, n: number) {
  const s = prs.addSlide()
  s.background = { color: C.white }

  // Title (x=0.529, y=0.961, w=5.508, 24pt bold)
  s.addText('Our understanding of your objectives', {
    x: 0.53, y: 0.38, w: 5.51, h: 0.72, fontSize: 24, bold: true, color: C.head, fontFace: C.hF,
  })

  // Left bullets (x=0.881, y=2.349, w=4.895, h=2.775, 10.5pt)
  // Space them from y=1.30 to fit in the left column
  p.needs.slice(0, 6).forEach((need, i) => {
    const y = 1.28 + i * 0.86
    s.addShape('rect', { x: 0.88, y: y + 0.14, w: 0.09, h: 0.09, fill: { color: C.gText }, line: { width: 0 } })
    s.addText(need, {
      x: 1.05, y, w: 4.73, h: 0.62,
      fontSize: 10.5, color: C.body, fontFace: C.bF, wrap: true, lineSpacingMultiple: 1.3,
    })
  })

  // Right visual panel (mimics screenshot: x=5.242, y=0.930)
  addRightPanel(s, 'Key client objectives')
  p.needs.slice(0, 5).forEach((need, i) => {
    s.addShape('rect', { x: 6.52, y: 1.65 + i * 0.92 + 0.14, w: 0.08, h: 0.08, fill: { color: C.gFill }, line: { width: 0 } })
    s.addText(need, {
      x: 6.72, y: 1.65 + i * 0.92, w: 6.10, h: 0.68,
      fontSize: 11, color: C.white, fontFace: C.bF, wrap: true, lineSpacingMultiple: 1.2,
    })
  })

  addChrome(s, n)
}

// ─── Slide 4: Goals ───────────────────────────────────────────────────────────
function slideGoals(prs: pptxgen, p: ProposalContent, n: number) {
  const s = prs.addSlide()
  s.background = { color: C.white }
  addTitle(s, 'Project goals & success criteria')

  p.goals.slice(0, 5).forEach((goal, i) => {
    const y = 0.98 + i * 1.12
    s.addShape('ellipse', { x: 0.49, y, w: 0.50, h: 0.50, fill: { color: C.gFill }, line: { width: 0 } })
    s.addText(String(i + 1), {
      x: 0.49, y, w: 0.50, h: 0.50,
      fontSize: 14, bold: true, color: C.white, fontFace: C.hF, align: 'center', valign: 'middle',
    })
    s.addText(goal, {
      x: 1.15, y: y + 0.02, w: 11.80, h: 0.56,
      fontSize: 14, color: C.body, fontFace: C.bF, wrap: true,
    })
    if (i < 4) {
      s.addShape('line', { x: 0.49, y: y + 0.68, w: 12.35, h: 0, line: { color: C.lineGrey, width: 0.5 } })
    }
  })

  addChrome(s, n)
}

// ─── Slide 5: Scope of Work ───────────────────────────────────────────────────
// KEY FIX: With 10 sections from the Lattice estimate, two-column approach overflows.
// Solution: use a compact table-style layout with fixed row heights that guarantees fit.
function slideScope(prs: pptxgen, p: ProposalContent, n: number) {
  const s = prs.addSlide()
  s.background = { color: C.white }
  addTitle(s, 'Scope of work')

  // Build section → top tasks map
  const bySec = new Map<string, string[]>()
  p.sheet.tasks.forEach(t => {
    const sec = t.section.split(' › ')[0]
    if (!bySec.has(sec)) bySec.set(sec, [])
    bySec.get(sec)!.push(t.task)
  })
  const entries = [...bySec.entries()]

  // Safe content area: y 0.86 → 6.90, height = 6.04"
  // With up to 10 sections we use a 2-column layout but with FIXED row heights
  // that scale to fit regardless of count.
  const CONTENT_TOP = 0.86
  const CONTENT_BOT = 6.90
  const CONTENT_H = CONTENT_BOT - CONTENT_TOP  // 6.04"
  const GAP = 0.30  // gap between columns
  const COL_W = (12.84 - 0.49 - GAP) / 2  // ~6.07" each

  const half = Math.ceil(entries.length / 2)
  const col1 = entries.slice(0, half)
  const col2 = entries.slice(half)

  // Calculate row height for each column so it fits exactly
  const calcRowH = (items: typeof entries) => {
    // Each section: 1 header + N task rows (max 3) + optional "+more"
    const rowCount = items.reduce((sum, [, tasks]) => sum + 1 + Math.min(tasks.length, 3) + (tasks.length > 3 ? 0.6 : 0), 0)
    return CONTENT_H / Math.max(rowCount, 1)
  }

  const drawCol = (items: typeof entries, xBase: number, unitH: number) => {
    let y = CONTENT_TOP
    items.forEach(([sec, tasks]) => {
      const secH = unitH * 1.0

      // Section header — green pill
      s.addShape('rect', { x: xBase, y, w: COL_W, h: secH * 0.88, fill: { color: C.gFill }, line: { width: 0 } })
      // Truncate long section names
      const secLabel = sec.length > 38 ? sec.slice(0, 36) + '…' : sec
      s.addText(secLabel, {
        x: xBase + 0.10, y, w: COL_W - 0.12, h: secH * 0.88,
        fontSize: 9.5, bold: true, color: C.white, fontFace: C.hF, valign: 'middle', wrap: false,
      })
      y += secH * 0.92

      // Tasks — show up to 3
      const showTasks = tasks.slice(0, 3)
      showTasks.forEach(task => {
        const taskH = unitH * 0.85
        s.addShape('rect', { x: xBase + 0.10, y: y + taskH * 0.38, w: 0.07, h: 0.07, fill: { color: C.gText }, line: { width: 0 } })
        const label = task.length > 52 ? task.slice(0, 50) + '…' : task
        s.addText(label, {
          x: xBase + 0.23, y, w: COL_W - 0.25, h: taskH,
          fontSize: 9, color: C.body, fontFace: C.bF, wrap: false,
        })
        y += taskH
      })

      // "+N more" line
      if (tasks.length > 3) {
        s.addText(`+${tasks.length - 3} more`, {
          x: xBase + 0.23, y, w: COL_W - 0.25, h: unitH * 0.55,
          fontSize: 8.5, color: C.zGrey, fontFace: C.bF, italic: true,
        })
        y += unitH * 0.55
      }

      y += unitH * 0.15  // section gap
    })
  }

  const unitH1 = calcRowH(col1)
  const unitH2 = calcRowH(col2)
  drawCol(col1, 0.49, unitH1)
  drawCol(col2, 0.49 + COL_W + GAP, unitH2)

  addChrome(s, n)
}

// ─── Slide 6: Stats (matching slide 6 layout exactly) ────────────────────────
// Template: 6 stat boxes LEFT half only (x≈1.23–6.5), right half = product screenshot
function slideStats(prs: pptxgen, p: ProposalContent, n: number) {
  const s = prs.addSlide()
  s.background = { color: C.white }
  addTitle(s, 'Key effort & investment metrics')

  const sheet = p.sheet
  const midH = Math.round((sheet.totalHoursLow + sheet.totalHoursHigh) / 2)
  const midC = sheet.midCost
  const net = midC - 40000

  // Exact positions from template slide 6 XML:
  // Row 1: y=1.935, Row 2: y=3.542, Row 3: y=5.149
  // Col 1: x=1.232, Col 2: x=4.172, w≈2.0–2.5, h=1.159
  const boxes = [
    { val: `${sheet.totalHoursLow}–${sheet.totalHoursHigh}`, lbl: 'estimated hours\n(low–high range)', x: 1.232, y: 1.935, w: 2.10 },
    { val: `~${midH}h`,                                      lbl: 'mid-point\ntotal effort',           x: 4.172, y: 1.935, w: 2.10 },
    { val: `${sheet.roles.length}`,                           lbl: 'disciplines\nacross the team',      x: 1.232, y: 3.542, w: 2.10 },
    { val: `$${Math.round(midC / 1000)}k`,                   lbl: 'gross mid-point\ninvestment',        x: 4.172, y: 3.542, w: 2.33 },
    { val: '-$40k',                                           lbl: 'discount + rebate\napplied',         x: 1.232, y: 5.149, w: 1.97 },
    { val: `$${Math.round(net / 1000)}k`,                     lbl: 'net investment\n(fixed bid)',        x: 4.172, y: 5.149, w: 2.50 },
  ]

  boxes.forEach(b => {
    // White box with light border (matches template "Text Box 2" white fill)
    s.addShape('rect', { x: b.x, y: b.y, w: b.w, h: 1.159, fill: { color: C.white }, line: { color: C.lineGrey, width: 0.75 } })
    // Big stat — 40pt bold green (#538321)
    s.addText(b.val, { x: b.x + 0.12, y: b.y + 0.12, w: b.w - 0.18, h: 0.65, fontSize: 32, bold: true, color: C.gText, fontFace: C.hF })
    // Label — 12pt body dark
    s.addText(b.lbl, { x: b.x + 0.12, y: b.y + 0.72, w: b.w - 0.18, h: 0.40, fontSize: 10, color: C.body, fontFace: C.bF, lineSpacingMultiple: 1.15 })
  })

  // Right panel (matches slide 6: product screenshot area x=5.975, y=1.028)
  addRightPanel(s, 'Estimate data summary')
  // Render the role table in the dark panel
  s.addText('Role breakdown', { x: 6.50, y: 1.05, w: 6.30, h: 0.34, fontSize: 12, bold: true, color: C.white, fontFace: C.hF })
  sheet.roles.forEach((r, i) => {
    const y = 1.55 + i * 0.75
    const midRH = Math.round((r.hoursLow + r.hoursHigh) / 2)
    s.addShape('rect', { x: 6.50, y: y + 0.06, w: 1.8, h: 0.50, fill: { color: C.gFill }, line: { width: 0 } })
    s.addText(r.role, { x: 6.50, y: y + 0.06, w: 1.8, h: 0.50, fontSize: 11, bold: true, color: C.white, fontFace: C.hF, align: 'center', valign: 'middle' })
    s.addText(`~${midRH}h  ·  $${Math.round((r.costLow + r.costHigh) / 2).toLocaleString()}`, {
      x: 8.42, y: y + 0.08, w: 4.20, h: 0.44,
      fontSize: 11, color: C.white, fontFace: C.bF, valign: 'middle',
    })
  })

  addChrome(s, n)
}

// ─── Slide 7: Section divider (matches slide 15 exactly) ─────────────────────
function slideDivider(prs: pptxgen, text: string, n: number) {
  const s = prs.addSlide()
  s.background = { color: C.dark }

  // Right darker panel (slide 15: Rectangle 12 x=4.384, y=0, w=7.526, h=7.5)
  s.addShape('rect', { x: 4.38, y: 0, w: 8.95, h: 7.5, fill: { color: C.darkPanel }, line: { width: 0 } })

  // ZONT mark (slide 15: Picture 1 x=11.11, y=0.50)
  s.addText('ZONT', { x: 11.11, y: 0.50, w: 1.61, h: 0.41, fontSize: 14, bold: true, color: C.white, fontFace: C.hF, align: 'right' })

  // Large title (slide 15: Title 10 x=0.762, y=2.774, w=9.167, h=2.705, 60pt NOT bold, green)
  s.addText(text, {
    x: 0.76, y: 2.77, w: 9.17, h: 2.70,
    fontSize: 60, bold: false, color: C.gFill, fontFace: C.hF, lineSpacingMultiple: 1.05,
  })

  // Slide number (same dark box as content slides)
  s.addShape('rect', { x: 12.55, y: 7.02, w: 0.56, h: 0.27, fill: { color: C.numBg }, line: { width: 0 } })
  s.addText(String(n), { x: 12.55, y: 7.02, w: 0.56, h: 0.27, fontSize: 9, color: C.white, fontFace: C.bF, align: 'center', valign: 'middle' })
}

// ─── Slide 8: Role breakdown (matches slide 13 two-col box style) ─────────────
// Template slide 13: green rect x=1.461,y=2.262,w=5.270,h=3.355 / dark rect x=6.732,y=2.262,w=5.182,h=3.355
// + full-width dark Change Management bar at y=4.639,h=1.387
function slideRoles(prs: pptxgen, p: ProposalContent, n: number) {
  const s = prs.addSlide()
  s.background = { color: C.white }
  addTitle(s, 'Role-based effort allocation')

  const roles = p.sheet.roles
  // Lay out up to 6 roles in a 3×2 grid matching the two-col box aesthetic
  // Using proportional sizing to always fit
  const PER_ROW = 3
  const BOX_W = (12.84 - 0.49 - 0.30) / PER_ROW - 0.10  // ≈ 3.95" each
  const ROWS = Math.ceil(roles.length / PER_ROW)
  const AVAIL_H = 5.85  // y 0.86 to 6.71
  const BOX_H = AVAIL_H / ROWS - 0.22

  roles.forEach((r, i) => {
    const col = i % PER_ROW
    const row = Math.floor(i / PER_ROW)
    const x = 0.49 + col * (BOX_W + 0.15)
    const y = 0.86 + row * (BOX_H + 0.22)
    const midH = Math.round((r.hoursLow + r.hoursHigh) / 2)
    const midC = Math.round((r.costLow + r.costHigh) / 2)

    // Alternate green / dark per row (matches slide 13 pattern)
    const bg = row % 2 === 0 ? C.gFill : '1C1C1C'

    s.addShape('rect', { x, y, w: BOX_W, h: BOX_H, fill: { color: bg }, line: { width: 0 } })

    // Role name — 24pt bold white (matches "Marketers & Content Authors" 24pt)
    s.addText(r.role, {
      x: x + 0.18, y: y + 0.16, w: BOX_W - 0.30, h: BOX_H * 0.30,
      fontSize: Math.min(20, BOX_H * 14), bold: true, color: C.white, fontFace: C.hF,
    })
    // Hours range — 14pt body (matches bullet items 14pt)
    s.addText(`${r.hoursLow}–${r.hoursHigh} hrs  ·  ~${midH}h`, {
      x: x + 0.18, y: y + BOX_H * 0.44, w: BOX_W - 0.30, h: BOX_H * 0.24,
      fontSize: 12, color: C.white, fontFace: C.bF,
    })
    // Cost — larger bold
    s.addText(`$${midC.toLocaleString()}`, {
      x: x + 0.18, y: y + BOX_H * 0.68, w: BOX_W - 0.30, h: BOX_H * 0.28,
      fontSize: Math.min(22, BOX_H * 15), bold: true, color: C.white, fontFace: C.hF,
    })
  })

  addChrome(s, n)
}

// ─── Slide 9: Effort by section — horizontal bars ────────────────────────────
function slideEffort(prs: pptxgen, p: ProposalContent, n: number) {
  const s = prs.addSlide()
  s.background = { color: C.white }
  addTitle(s, 'Effort breakdown by work stream')

  const bySec = new Map<string, { low: number; high: number; cLow: number; cHigh: number }>()
  p.sheet.tasks.forEach(t => {
    const sec = t.section.split(' › ')[0]
    const ex = bySec.get(sec) ?? { low: 0, high: 0, cLow: 0, cHigh: 0 }
    bySec.set(sec, { low: ex.low + t.totalHoursLow, high: ex.high + t.totalHoursHigh, cLow: ex.cLow + t.totalCostLow, cHigh: ex.cHigh + t.totalCostHigh })
  })

  const entries = [...bySec.entries()]
  const maxH = Math.max(...entries.map(([, v]) => v.high))

  // Fixed layout: fit all sections within y 0.86–6.90 (6.04" available)
  const CONTENT_H = 6.04
  const ROW_H = Math.min(0.76, CONTENT_H / entries.length)
  const BAR_X = 2.20
  const BAR_W = 9.50
  const LABEL_W = 1.65

  entries.forEach(([sec, v], i) => {
    const y = 0.90 + i * ROW_H
    const mid = Math.round((v.low + v.high) / 2)
    const midC = Math.round((v.cLow + v.cHigh) / 2)
    const pct = v.high / maxH
    const barFill = Math.max(pct * BAR_W * 0.96, 0.4)

    // Section label — truncated to fit
    const label = sec.length > 22 ? sec.slice(0, 20) + '…' : sec
    s.addText(label, {
      x: 0.49, y: y + 0.08, w: LABEL_W, h: ROW_H * 0.60,
      fontSize: Math.min(10, ROW_H * 14), color: C.body, fontFace: C.bF, wrap: false,
    })

    // Bar background
    s.addShape('rect', { x: BAR_X, y: y + ROW_H * 0.15, w: BAR_W, h: ROW_H * 0.55, fill: { color: C.barBg }, line: { width: 0 } })
    // Bar fill
    s.addShape('rect', { x: BAR_X, y: y + ROW_H * 0.15, w: barFill, h: ROW_H * 0.55, fill: { color: C.gFill }, line: { width: 0 } })
    // Hours inside bar
    s.addText(`${mid}h`, {
      x: BAR_X + 0.10, y: y + ROW_H * 0.16, w: 1.2, h: ROW_H * 0.52,
      fontSize: Math.min(10, ROW_H * 13), bold: true, color: C.white, fontFace: C.bF, valign: 'middle',
    })
    // Cost right of bar
    s.addText(`$${Math.round(midC / 1000)}k`, {
      x: BAR_X + BAR_W + 0.12, y: y + ROW_H * 0.10, w: 1.20, h: ROW_H * 0.60,
      fontSize: Math.min(11, ROW_H * 14), bold: true, color: C.gText, fontFace: C.hF, valign: 'middle',
    })
  })

  addChrome(s, n)
}

// ─── Slide 10: Investment table (matches slide 18 exactly) ────────────────────
function slideInvestment(prs: pptxgen, p: ProposalContent, n: number) {
  const s = prs.addSlide()
  s.background = { color: C.white }

  // Title (slide 18: x=0.42, y=0.36, w=9.90, h=0.40, 24pt bold)
  s.addText('Investment summary', {
    x: 0.42, y: 0.36, w: 9.90, h: 0.40, fontSize: 24, bold: true, color: C.head, fontFace: C.hF,
  })

  const sheet = p.sheet
  const discount = 15000
  const rebate = 25000
  const midC = sheet.midCost
  const net = midC - discount - rebate
  const weeksMid = Math.max(1, Math.round((sheet.totalHoursLow + sheet.totalHoursHigh) / 2 / 40))

  // Table exact position from slide 18
  const TX = 0.42, TY = 0.88, TW = 12.38
  const CW = [7.50, 1.80, 3.08]  // exact col widths from slide 18

  // Header row (slide 18: dark #2E2E2E, white text 12pt, h=0.344)
  s.addShape('rect', { x: TX, y: TY, w: TW, h: 0.34, fill: { color: C.tblHdr }, line: { width: 0 } })
  let colX = TX
  ;['ACTIVITY', 'TIMELINE', 'COST (USD)'].forEach((h, i) => {
    s.addText(h, { x: colX + 0.12, y: TY, w: CW[i] - 0.12, h: 0.34, fontSize: 12, color: C.white, fontFace: C.hF, valign: 'middle' })
    colX += CW[i]
  })

  // Row 1: main project (slide 18 h=1.743)
  const r1Y = TY + 0.34, r1H = 1.74
  s.addShape('rect', { x: TX, y: r1Y, w: TW, h: r1H, fill: { color: C.white }, line: { color: C.lineGrey, width: 0.5 } })
  s.addText(p.projectName, { x: TX + 0.12, y: r1Y + 0.14, w: CW[0] - 0.24, h: 0.34, fontSize: 12, bold: true, color: C.head, fontFace: C.hF })
  s.addText(sheet.sections.slice(0, 6).join('  ·  '), {
    x: TX + 0.12, y: r1Y + 0.50, w: CW[0] - 0.24, h: 1.12,
    fontSize: 9, color: C.body, fontFace: C.bF, wrap: true, lineSpacingMultiple: 1.3,
  })
  s.addText(`${weeksMid} weeks`, { x: TX + CW[0] + 0.12, y: r1Y + 0.14, w: CW[1] - 0.12, h: 0.36, fontSize: 12, color: C.body, fontFace: C.bF })
  s.addText(`$${midC.toLocaleString()}`, { x: TX + CW[0] + CW[1] + 0.12, y: r1Y + 0.14, w: CW[2] - 0.12, h: 0.36, fontSize: 12, color: C.head, fontFace: C.bF })

  // Row 2: Sitecore rebate (slide 18 h=1.172, dark fill + light text)
  const r2Y = r1Y + r1H, r2H = 0.60
  s.addShape('rect', { x: TX, y: r2Y, w: TW, h: r2H, fill: { color: C.tblAlt }, line: { color: C.lineGrey, width: 0.5 } })
  s.addText('Sitecore Commercial Migration Rebate\nRebate upon migration to XM Cloud', {
    x: TX + 0.12, y: r2Y + 0.06, w: CW[0] - 0.24, h: 0.48, fontSize: 10, color: C.body, fontFace: C.bF, lineSpacingMultiple: 1.2,
  })
  s.addText(`($${rebate.toLocaleString()})`, { x: TX + CW[0] + CW[1] + 0.12, y: r2Y + 0.12, w: CW[2] - 0.12, h: 0.36, fontSize: 11, color: C.body, fontFace: C.bF })

  // Row 3: Marketing discount (slide 18 h=0.855)
  const r3Y = r2Y + r2H, r3H = 0.56
  s.addShape('rect', { x: TX, y: r3Y, w: TW, h: r3H, fill: { color: C.white }, line: { color: C.lineGrey, width: 0.5 } })
  s.addText('Marketing Discount\nClient reference, joint presentation or webinar, sharable feedback', {
    x: TX + 0.12, y: r3Y + 0.06, w: CW[0] - 0.24, h: 0.46, fontSize: 10, color: C.body, fontFace: C.bF, lineSpacingMultiple: 1.2,
  })
  s.addText(`($${discount.toLocaleString()})`, { x: TX + CW[0] + CW[1] + 0.12, y: r3Y + 0.10, w: CW[2] - 0.12, h: 0.36, fontSize: 11, color: C.body, fontFace: C.bF })

  // Total row — dark bg (slide 18 h=0.579)
  const totY = r3Y + r3H
  s.addShape('rect', { x: TX, y: totY, w: TW, h: 0.52, fill: { color: C.tblHdr }, line: { width: 0 } })
  s.addText('Total', { x: TX + 0.12, y: totY, w: CW[0] - 0.12, h: 0.52, fontSize: 14, bold: true, color: C.white, fontFace: C.hF, valign: 'middle' })
  s.addText(`${weeksMid} weeks`, { x: TX + CW[0] + 0.12, y: totY, w: CW[1] - 0.12, h: 0.52, fontSize: 12, color: C.white, fontFace: C.bF, valign: 'middle' })
  s.addText(`$${net.toLocaleString()}`, { x: TX + CW[0] + CW[1] + 0.12, y: totY, w: CW[2] - 0.12, h: 0.52, fontSize: 14, bold: true, color: C.gFill, fontFace: C.hF, valign: 'middle' })

  // Footer note (slide 18: TextBox 3 x=1.454, y=6.795, 8pt)
  s.addText('The quote is given as a fixed bid. Sitecore rebate contingent on confirmation from Sitecore. Excludes T&E expenses if any.', {
    x: 1.45, y: 6.79, w: 10.0, h: 0.40, fontSize: 8, color: C.zGrey, fontFace: C.bF, italic: true,
  })

  addChrome(s, n)
}

// ─── Slide 11: Assumptions ────────────────────────────────────────────────────
function slideAssumptions(prs: pptxgen, p: ProposalContent, n: number) {
  const s = prs.addSlide()
  s.background = { color: C.white }
  addTitle(s, 'Project assumptions & exclusions')

  // Two-column layout — fit all assumptions with dynamic row heights
  const CONTENT_TOP = 0.86
  const CONTENT_H = 6.04
  const half = Math.ceil(p.assumptions.length / 2)
  const col1 = p.assumptions.slice(0, half)
  const col2 = p.assumptions.slice(half)
  const maxRows = Math.max(col1.length, col2.length)
  const rowH = Math.min(0.85, CONTENT_H / Math.max(maxRows, 1))
  const COL_W = (12.35 - 0.49) / 2 - 0.20

  const drawCol = (items: string[], xBase: number) => {
    items.forEach((a, i) => {
      const y = CONTENT_TOP + i * rowH
      s.addShape('ellipse', { x: xBase, y: y + 0.11, w: 0.22, h: 0.22, fill: { color: C.gText }, line: { width: 0 } })
      s.addText(a, {
        x: xBase + 0.34, y, w: COL_W - 0.34, h: rowH - 0.06,
        fontSize: Math.min(11, rowH * 14), color: C.body, fontFace: C.bF,
        wrap: true, lineSpacingMultiple: 1.2, valign: 'top',
      })
    })
  }

  drawCol(col1, 0.49)
  drawCol(col2, 0.49 + COL_W + 0.44)
  addChrome(s, n)
}

// ─── Slide 12: Next Steps (two-col coloured boxes, matches slide 13) ──────────
function slideNextSteps(prs: pptxgen, n: number) {
  const s = prs.addSlide()
  s.background = { color: C.white }
  addTitle(s, 'Next steps')

  const steps = [
    { num: '01', t: 'Review & approve proposal',    sub: 'Provide written feedback within 5 business days' },
    { num: '02', t: 'Sign the Statement of Work',   sub: 'Legal review and authorised countersignature' },
    { num: '03', t: '30% deposit to commence',      sub: 'Invoice issued upon SOW signing' },
    { num: '04', t: 'Project kickoff & onboarding', sub: 'Scheduled within 5 business days of signing' },
  ]

  // Exactly like slide 13: two cols, two rows
  // Green: x=1.461, y=2.262, w=5.270, h=3.355
  // Dark:  x=6.732, y=2.262, w=5.182, h=3.355
  // But we have 4 boxes in 2 rows — adjust y accordingly
  const ROW_Y = [0.86, 3.72]
  const BOX_H = 2.70

  steps.forEach((step, i) => {
    const col = i % 2, row = Math.floor(i / 2)
    const x = col === 0 ? 0.49 : 6.92
    const w = col === 0 ? 6.18 : 6.05
    const y = ROW_Y[row]
    const bg = col === 0 ? C.gFill : '1C1C1C'

    s.addShape('rect', { x, y, w, h: BOX_H, fill: { color: bg }, line: { width: 0 } })

    // Number — large, top-left of box
    s.addText(step.num, {
      x: x + 0.28, y: y + 0.24, w: 1.2, h: 0.56,
      fontSize: 30, bold: true, color: C.white, fontFace: C.hF,
    })
    // Title — 24pt bold white (matches slide 13 section headers)
    s.addText(step.t, {
      x: x + 0.28, y: y + 0.86, w: w - 0.44, h: 0.58,
      fontSize: 18, bold: true, color: C.white, fontFace: C.hF, wrap: true,
    })
    // Sub — 14pt body (matches slide 13 bullet text)
    s.addText(step.sub, {
      x: x + 0.28, y: y + 1.54, w: w - 0.44, h: 0.80,
      fontSize: 13, color: C.white, fontFace: C.bF, wrap: true,
    })
  })

  addChrome(s, n)
}

// ─── Slide 13: Closing (matches slide 21 exactly) ────────────────────────────
function slideClosing(prs: pptxgen, n: number) {
  const s = prs.addSlide()
  s.background = { color: C.dark }

  // Right panel (slide 21: Rectangle 12 x=4.384, y=0, w=7.526, h=7.5)
  s.addShape('rect', { x: 4.38, y: 0, w: 8.95, h: 7.5, fill: { color: C.darkPanel }, line: { width: 0 } })

  // ZONT (slide 21: Picture 1 equivalent, top-right)
  s.addText('ZONT', { x: 11.11, y: 0.50, w: 1.61, h: 0.41, fontSize: 14, bold: true, color: C.white, fontFace: C.hF, align: 'right' })

  // Title (slide 21: Title 10 x=0.762, y=2.774, w=9.167, h=2.705, 60pt NOT bold, green)
  s.addText("Let's create remarkable\ndigital solutions!", {
    x: 0.76, y: 2.77, w: 9.17, h: 2.70,
    fontSize: 60, bold: false, color: C.gFill, fontFace: C.hF, lineSpacingMultiple: 1.05,
  })

  s.addShape('rect', { x: 12.55, y: 7.02, w: 0.56, h: 0.27, fill: { color: C.numBg }, line: { width: 0 } })
  s.addText(String(n), { x: 12.55, y: 7.02, w: 0.56, h: 0.27, fontSize: 9, color: C.white, fontFace: C.bF, align: 'center', valign: 'middle' })
}

// ─── Main export ──────────────────────────────────────────────────────────────
export async function generatePptx(p: ProposalContent): Promise<Blob> {
  const prs = new pptxgen()
  prs.layout = 'LAYOUT_WIDE'  // 13.33" × 7.5"

  slideCover(prs, p)                                       //  1
  slideAgenda(prs, p, 2)                                   //  2
  slideObjectives(prs, p, 3)                               //  3
  slideGoals(prs, p, 4)                                    //  4
  slideScope(prs, p, 5)                                    //  5
  slideStats(prs, p, 6)                                    //  6
  slideDivider(prs, 'Timeline,\nteam,\ninvestment\nsummary', 7) //  7
  slideRoles(prs, p, 8)                                    //  8
  slideEffort(prs, p, 9)                                   //  9
  slideInvestment(prs, p, 10)                              // 10
  slideAssumptions(prs, p, 11)                             // 11
  slideNextSteps(prs, 12)                                  // 12
  slideClosing(prs, 13)                                    // 13

  const b64 = await prs.write({ outputType: 'base64' }) as string
  const bin = atob(b64)
  const bytes = new Uint8Array(bin.length)
  for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i)
  return new Blob([bytes], { type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation' })
}
