import pptxgen from 'pptxgenjs'
import type { ProposalContent } from '../store/appStore'

// ─── Exact brand tokens from template XML ────────────────────────────────────
const C = {
  white:    'FFFFFF',
  black:    '000000',
  gFill:    '549E39',   // green shape fill  (#549E39)
  gText:    '538321',   // green text / icon fill (#538321)
  body:     '3D3B47',   // body copy
  dark:     '1C1C1C',   // divider/closing bg
  darkP:    '252525',   // right panel on dark slides
  tblH:     '2E2E2E',   // table header / slide-num bg
  tblAlt:   'F2F2F2',   // alt table row
  zGrey:    '9EA8B3',   // ZONT wordmark
  line:     'E0E0E0',   // thin divider lines
  orange:   'ED7D31',   // accent2 — UAT bars
  gold:     'FFC000',   // accent4 — onboarding bars
  blue:     '4472C4',   // accent1
  hF: 'Calibri Light',
  bF: 'Calibri',
}

// ─── Chrome shared across all content slides ──────────────────────────────────
function chrome(s: pptxgen.Slide, n: number) {
  // Slide number box (x=12.55, y=7.02, w=0.56, h=0.27, fill=#2E2E2E) — exact from template
  s.addShape('rect', { x:12.55, y:7.02, w:0.56, h:0.27, fill:{color:C.tblH}, line:{width:0} })
  s.addText(String(n), { x:12.55, y:7.02, w:0.56, h:0.27, fontSize:9, color:C.white, fontFace:C.bF, align:'center', valign:'middle' })
  // ZONT wordmark (approximate position varies per slide, but approx bottom area)
  s.addText('ZONT', { x:10.15, y:7.05, w:1.2, h:0.28, fontSize:11, color:C.zGrey, fontFace:C.hF, align:'right' })
}

// Standard content-slide title (x=0.49, y=0.28–0.36, 24pt bold)
function H(s: pptxgen.Slide, text: string, y = 0.30) {
  s.addText(text, { x:0.49, y, w:12.35, h:0.44, fontSize:24, bold:true, color:C.black, fontFace:C.hF })
}

// ─── Slide 1: Cover ───────────────────────────────────────────────────────────
// Green left panel (w=7.51), right photo area, title 44pt, subtitle 16pt
function s1Cover(prs: pptxgen, p: ProposalContent) {
  const s = prs.addSlide()
  s.background = { color: C.white }
  // Green left panel  x=0 y=0.004 w=7.514 h=7.522
  s.addShape('rect', { x:0, y:0, w:7.51, h:7.5, fill:{color:C.gFill}, line:{width:0} })
  // Green overlay banner behind right logos  x=4.523 y=2.967 w=7.553 h=1.608
  s.addShape('rect', { x:4.52, y:2.97, w:8.81, h:1.61, fill:{color:C.gFill}, line:{width:0} })
  // Logo row  (ZONT | CLIENT)
  s.addText(`ZONT  |  ${p.clientName.toUpperCase()}`, {
    x:0.61, y:0.55, w:5.5, h:0.40, fontSize:13, color:C.white, fontFace:C.bF,
  })
  // Title  x=0.611 y=2.688 w=6.631 44pt bold
  s.addText(p.projectName, {
    x:0.61, y:2.69, w:6.63, h:1.00, fontSize:44, bold:true, color:C.white, fontFace:C.hF,
  })
  // Subtitle  x=0.638 y=3.971 16pt
  s.addText(`AI-powered proposal for ${p.clientName}`, {
    x:0.64, y:3.97, w:5.94, h:0.38, fontSize:16, color:C.white, fontFace:C.bF,
  })
  s.addText(p.date, {
    x:0.64, y:4.44, w:5.94, h:0.32, fontSize:13, color:C.white, fontFace:C.bF, italic:true,
  })
}

// ─── Slide 2: Agenda ─────────────────────────────────────────────────────────
// Template slide 3: "Agenda" centred ~x=2.24, three right-side items with green circles
function s2Agenda(prs: pptxgen, p: ProposalContent, n: number) {
  const s = prs.addSlide()
  s.background = { color: C.white }

  // Vertical green bar  x=7.138 y=-0.017 w=0.069 h=7.517
  s.addShape('rect', { x:7.138, y:0, w:0.069, h:7.5, fill:{color:C.gText}, line:{width:0} })

  // "Agenda"  x=2.239 y=3.279 40pt bold (template has it centre-left)
  s.addText('Agenda', { x:0.5, y:3.18, w:6.4, h:0.65, fontSize:40, bold:true, color:C.black, fontFace:C.hF })

  // Three agenda items with green circles (x=6.923, y=2.589/3.602/4.615, w=0.5, h=0.5)
  const items = [
    'Our understanding of your objectives',
    'Our recommended solution',
    'Timeline, team, investment summary',
  ]
  const dotY = [2.589, 3.602, 4.615]
  const txtX = 7.594
  const txtW = 4.157
  items.forEach((item, i) => {
    s.addShape('ellipse', { x:6.923, y:dotY[i]+0.08, w:0.32, h:0.32, fill:{color:C.gText}, line:{width:0} })
    s.addText(item, { x:txtX, y:dotY[i]+0.02, w:txtW, h:0.45, fontSize:16, color:C.black, fontFace:C.bF })
  })

  chrome(s, n)
}

// ─── Slide 3: Objectives ─────────────────────────────────────────────────────
// Template slide 4: title left, bullet list, dark right panel w/ screenshot
function s3Objectives(prs: pptxgen, p: ProposalContent, n: number) {
  const s = prs.addSlide()
  s.background = { color: C.white }

  // Title  x=0.529 y=0.961 w=5.508 24pt bold
  s.addText('Our understanding of your objectives', {
    x:0.53, y:0.38, w:5.51, h:0.65, fontSize:24, bold:true, color:C.black, fontFace:C.hF,
  })

  // Left bullet list  x=0.881 y=2.349 w=4.895 h=2.775 10.5pt
  const bullets = p.needs.slice(0, 6)
  bullets.forEach((b, i) => {
    const y = 1.25 + i * 0.88
    s.addShape('rect', { x:0.88, y:y+0.16, w:0.09, h:0.09, fill:{color:C.gText}, line:{width:0} })
    s.addText(b, { x:1.05, y, w:4.73, h:0.65, fontSize:10.5, color:C.body, fontFace:C.bF, wrap:true, lineSpacingMultiple:1.3 })
  })

  // Dark right panel (mimics product screenshot area x=5.242 y=0.930)
  s.addShape('rect', { x:6.30, y:0.88, w:6.70, h:6.28, fill:{color:'1A1A18'}, line:{width:0} })
  s.addShape('rect', { x:6.30, y:0.88, w:6.70, h:0.055, fill:{color:C.gFill}, line:{width:0} })
  s.addText('Client objectives', { x:6.50, y:1.06, w:6.30, h:0.36, fontSize:13, bold:true, color:C.white, fontFace:C.hF })
  bullets.slice(0,5).forEach((b, i) => {
    s.addShape('rect', { x:6.52, y:1.65+i*0.95+0.14, w:0.08, h:0.08, fill:{color:C.gFill}, line:{width:0} })
    s.addText(b, { x:6.72, y:1.65+i*0.95, w:6.10, h:0.70, fontSize:11, color:C.white, fontFace:C.bF, wrap:true, lineSpacingMultiple:1.2 })
  })

  chrome(s, n)
}

// ─── Slide 4: Core Tenants / Solution Pillars ────────────────────────────────
// Template slide 5: radial 6-item layout with centre circle
function s4Tenants(prs: pptxgen, p: ProposalContent, n: number) {
  const s = prs.addSlide()
  s.background = { color: C.white }
  H(s, `Core tenants of the ${p.clientName} digital modernization`)

  // Centre circle
  s.addShape('ellipse', { x:5.0, y:2.2, w:3.0, h:3.0, fill:{color:C.gFill}, line:{width:0} })
  s.addText('Modernization,\nagility,\ngrowth', {
    x:5.0, y:2.2, w:3.0, h:3.0, fontSize:14, bold:true, color:C.white, fontFace:C.hF, align:'center', valign:'middle',
  })

  // 6 pillar positions around the centre (roughly matching template)
  const pillars = [
    { num:'A1', title:'Authoring self-service',       desc:'WYSIWYG authoring, flexible and reusable components, new component creation without development.' },
    { num:'A2', title:'Rapid modernization',          desc:'Content and components migrated rapidly with AI-assisted scripting and code conversion.' },
    { num:'A3', title:'Security & quality control',   desc:'OWASP Top 10, PII standard practices; AI development instructions; code quality scans.' },
    { num:'A4', title:'Engaging & powerful search',   desc:'Search that supports Q&A, personalized results, and AI-based content recommendations.' },
    { num:'A5', title:'Modern & future-proof stack',  desc:'Scalable, secure, long-living and cost-effective environments. Modern light-weight implementation.' },
    { num:'A6', title:'Scalability and flexibility',  desc:'Global atomic content model — flexible, extensible, scalable, modular, and reusable.' },
  ]

  // Positions: top-left, top-right, mid-left, mid-right, bottom-left, bottom-right
  const positions = [
    { x:0.3,  y:1.1 }, { x:9.1,  y:1.1 },
    { x:0.3,  y:3.3 }, { x:9.1,  y:3.3 },
    { x:0.3,  y:5.3 }, { x:9.1,  y:5.3 },
  ]

  pillars.forEach((pl, i) => {
    const { x, y } = positions[i]
    const bw = 3.4
    s.addShape('rect', { x, y, w:bw, h:1.75, fill:{color:'F8F8F8'}, line:{color:C.line, width:0.5} })
    // Green number badge
    s.addShape('ellipse', { x:x+0.12, y:y+0.12, w:0.38, h:0.38, fill:{color:C.gFill}, line:{width:0} })
    s.addText(String(i+1).padStart(2,'0'), { x:x+0.12, y:y+0.12, w:0.38, h:0.38, fontSize:10, bold:true, color:C.white, fontFace:C.hF, align:'center', valign:'middle' })
    s.addText(pl.title, { x:x+0.58, y:y+0.12, w:bw-0.68, h:0.36, fontSize:11, bold:true, color:C.gText, fontFace:C.hF })
    s.addText(pl.desc, { x:x+0.12, y:y+0.54, w:bw-0.18, h:1.1, fontSize:9.5, color:C.body, fontFace:C.bF, wrap:true, lineSpacingMultiple:1.2 })
  })

  chrome(s, n)
}

// ─── Slide 5: Stats callouts ──────────────────────────────────────────────────
// Template slide 6: 6 small stat boxes left half + right screenshot
// Exact positions: Row1 y=1.935, Row2 y=3.542, Row3 y=5.149  Col1 x=1.232, Col2 x=4.172
function s5Stats(prs: pptxgen, p: ProposalContent, n: number) {
  const s = prs.addSlide()
  s.background = { color: C.white }
  H(s, `${p.projectName} is a leap towards self-service and lower TCO`)

  const sheet = p.sheet
  const midH  = Math.round((sheet.totalHoursLow  + sheet.totalHoursHigh)  / 2)
  const midC  = sheet.midCost
  const net   = midC - 40000

  const boxes = [
    { val:`${sheet.totalHoursLow}–${sheet.totalHoursHigh}`, lbl:'estimated\nhours range',         x:1.232, y:1.935, w:2.10 },
    { val:`~${midH}h`,                                       lbl:'mid-point\ntotal effort',         x:4.172, y:1.935, w:2.10 },
    { val:`${sheet.roles.length}`,                            lbl:'roles across\nthe team',          x:1.232, y:3.542, w:2.10 },
    { val:`$${Math.round(midC/1000)}k`,                      lbl:'gross mid-point\ninvestment',      x:4.172, y:3.542, w:2.33 },
    { val:'-$40k',                                            lbl:'discount +\nrebate applied',      x:1.232, y:5.149, w:1.97 },
    { val:`$${Math.round(net/1000)}k`,                       lbl:'net investment\n(fixed bid)',       x:4.172, y:5.149, w:2.50 },
  ]

  boxes.forEach(b => {
    s.addShape('rect', { x:b.x, y:b.y, w:b.w, h:1.159, fill:{color:C.white}, line:{color:C.line, width:0.75} })
    s.addText(b.val, { x:b.x+0.12, y:b.y+0.12, w:b.w-0.18, h:0.62, fontSize:32, bold:true, color:C.gText, fontFace:C.hF })
    s.addText(b.lbl, { x:b.x+0.12, y:b.y+0.72, w:b.w-0.18, h:0.40, fontSize:10, color:C.body, fontFace:C.bF, lineSpacingMultiple:1.1 })
  })

  // Right screenshot panel
  s.addShape('rect', { x:6.85, y:0.88, w:6.15, h:6.28, fill:{color:'1A1A18'}, line:{width:0} })
  s.addShape('rect', { x:6.85, y:0.88, w:6.15, h:0.055, fill:{color:C.gFill}, line:{width:0} })
  s.addText('Role & effort summary', { x:7.05, y:1.06, w:5.8, h:0.36, fontSize:12, bold:true, color:C.white, fontFace:C.hF })
  sheet.roles.forEach((r, i) => {
    const y = 1.58 + i * 0.77
    const midRH = Math.round((r.hoursLow + r.hoursHigh) / 2)
    s.addShape('rect', { x:7.05, y:y+0.06, w:1.8, h:0.50, fill:{color:C.gFill}, line:{width:0} })
    s.addText(r.role, { x:7.05, y:y+0.06, w:1.8, h:0.50, fontSize:11, bold:true, color:C.white, fontFace:C.hF, align:'center', valign:'middle' })
    s.addText(`~${midRH}h  ·  $${Math.round((r.costLow+r.costHigh)/2).toLocaleString()}`,
      { x:8.97, y:y+0.08, w:3.85, h:0.44, fontSize:11, color:C.white, fontFace:C.bF, valign:'middle' })
  })

  chrome(s, n)
}

// ─── Slide 6: Scope of work ───────────────────────────────────────────────────
// Full-width two-col compact layout. Dynamic row sizing to prevent overflow.
function s6Scope(prs: pptxgen, p: ProposalContent, n: number) {
  const s = prs.addSlide()
  s.background = { color: C.white }
  H(s, 'Scope of work')

  const bySec = new Map<string, string[]>()
  p.sheet.tasks.forEach(t => {
    const sec = t.section.split(' › ')[0]
    if (!bySec.has(sec)) bySec.set(sec, [])
    bySec.get(sec)!.push(t.task)
  })
  const entries = [...bySec.entries()]
  const half    = Math.ceil(entries.length / 2)
  const col1    = entries.slice(0, half)
  const col2    = entries.slice(half)

  // Available height: 0.86 → 6.90 = 6.04"
  const TOP = 0.86, BOT = 6.90, AVAIL = BOT - TOP
  const GAP  = 0.28
  const COL_W = (12.84 - 0.49 - GAP) / 2  // ≈ 6.035"

  const calcUnit = (items: typeof entries) => {
    const totalRows = items.reduce((s, [, tasks]) =>
      s + 1 + Math.min(tasks.length, 3) + (tasks.length > 3 ? 0.55 : 0), 0)
    return AVAIL / Math.max(totalRows, 1)
  }

  const drawCol = (items: typeof entries, xBase: number, unit: number) => {
    let y = TOP
    items.forEach(([sec, tasks]) => {
      const hdrH = unit * 0.88
      // Green header pill
      s.addShape('rect', { x:xBase, y, w:COL_W, h:hdrH, fill:{color:C.gFill}, line:{width:0} })
      const lbl = sec.length > 40 ? sec.slice(0,38)+'…' : sec
      s.addText(lbl, { x:xBase+0.10, y, w:COL_W-0.12, h:hdrH, fontSize:9.5, bold:true, color:C.white, fontFace:C.hF, valign:'middle', wrap:false })
      y += hdrH + unit * 0.04

      tasks.slice(0, 3).forEach(task => {
        const th = unit * 0.82
        s.addShape('rect', { x:xBase+0.10, y:y+th*0.38, w:0.07, h:0.07, fill:{color:C.gText}, line:{width:0} })
        const tl = task.length > 54 ? task.slice(0,52)+'…' : task
        s.addText(tl, { x:xBase+0.23, y, w:COL_W-0.25, h:th, fontSize:9, color:C.body, fontFace:C.bF, wrap:false })
        y += th
      })
      if (tasks.length > 3) {
        const mh = unit * 0.55
        s.addText(`+${tasks.length-3} more`, { x:xBase+0.23, y, w:COL_W-0.25, h:mh, fontSize:8.5, color:C.zGrey, fontFace:C.bF, italic:true })
        y += mh
      }
      y += unit * 0.12
    })
  }

  drawCol(col1, 0.49, calcUnit(col1))
  drawCol(col2, 0.49 + COL_W + GAP, calcUnit(col2))
  chrome(s, n)
}

// ─── Slide 7: QA & Testing table ─────────────────────────────────────────────
// Matches template slide 12 exactly (full-width table, Included/N/A column)
function s7QA(prs: pptxgen, p: ProposalContent, n: number) {
  const s = prs.addSlide()
  s.background = { color: C.white }
  H(s, `We will ensure that the migration meets your quality standards`)

  const rows = [
    ['Code Quality Testing',         'Yes', 'Linting for a local quality scan of Next.js code.'],
    ['Functional Testing',           'Yes', 'Functional testing of components to ensure parity.'],
    ['Regression Testing',           'Yes', 'Automated visual regression testing of the website.'],
    ['User Acceptance Testing',      'Yes', `Ensures that the website works per ${p.clientName} standards.`],
    ['Page Load Speed Testing',      'N/A', `Ensures that the website pages load rapidly per ${p.clientName} standards.`],
    ['Load Testing',                 'N/A', `Ensures that the website can support a particular load.`],
    ['Pen Testing',                  'N/A', 'Manual and automated security penetration testing of the website.'],
    ['SAST',                         'N/A', 'A static security scan of the website source code.'],
    ['DAST',                         'N/A', 'A dynamic security scan of the website with a series of known attacks.'],
    ['Boundary/Exploratory Testing', 'N/A', 'Exploratory testing focused on testing use cases not in requirements.'],
    ['Stress Testing',               'N/A', 'Finds a maximum load that the website is capable of supporting.'],
  ]

  const TX = 0.42, TY = 0.86, TW = 12.50
  const CW = [3.2, 0.9, 8.4]

  // Header
  s.addShape('rect', { x:TX, y:TY, w:TW, h:0.34, fill:{color:C.tblH}, line:{width:0} })
  let cx = TX
  ;['Test Type','Included','Goal'].forEach((h, i) => {
    s.addText(h, { x:cx+0.10, y:TY, w:CW[i]-0.10, h:0.34, fontSize:11, color:C.white, fontFace:C.hF, valign:'middle' })
    cx += CW[i]
  })

  // Data rows — dynamic height to fit all 11
  const avail = 6.02
  const rowH  = avail / rows.length

  rows.forEach((row, i) => {
    const ry = TY + 0.34 + i * rowH
    const bg = i % 2 === 0 ? C.white : C.tblAlt
    s.addShape('rect', { x:TX, y:ry, w:TW, h:rowH, fill:{color:bg}, line:{color:C.line, width:0.3} })
    cx = TX
    row.forEach((cell, j) => {
      const color = j === 1 ? (cell === 'Yes' ? C.gText : C.body) : C.body
      const bold  = j === 1 && cell === 'Yes'
      s.addText(cell, { x:cx+0.10, y:ry+0.03, w:CW[j]-0.12, h:rowH-0.04, fontSize:9, color, fontFace:C.bF, valign:'middle', bold, wrap:true })
      cx += CW[j]
    })
  })

  chrome(s, n)
}

// ─── Slide 8: Enablement ─────────────────────────────────────────────────────
// Matches template slide 13: green box left, dark box right, full-width bar below
function s8Enablement(prs: pptxgen, p: ProposalContent, n: number) {
  const s = prs.addSlide()
  s.background = { color: C.white }
  H(s, 'Our approach ensures that all stakeholders go through full enablement')

  // Green box (x=1.461, y=2.262, w=5.270, h=3.355, fill=#538321)
  s.addShape('rect', { x:1.461, y:2.262, w:5.270, h:3.355, fill:{color:C.gText}, line:{width:0} })
  s.addText('Marketers & Content Authors', { x:1.65, y:2.42, w:4.9, h:0.48, fontSize:20, bold:true, color:C.white, fontFace:C.hF })
  ;[
    'A solution-based content authoring training',
    'Additional documentation, videos, and other resources for in-depth learning',
    'A two-week post-release hyper-support for the teams for newly developed components',
  ].forEach((line, i) => {
    s.addText('• '+line, { x:1.65, y:3.00+i*0.70, w:4.9, h:0.55, fontSize:13, color:C.white, fontFace:C.bF, wrap:true })
  })

  // Dark box (x=6.732, y=2.262, w=5.182, h=3.355, fill=#000000)
  s.addShape('rect', { x:6.732, y:2.262, w:5.182, h:3.355, fill:{color:C.black}, line:{width:0} })
  s.addText('Technologists', { x:6.93, y:2.42, w:4.8, h:0.48, fontSize:20, bold:true, color:C.white, fontFace:C.hF })
  ;[
    'A technical handoff session.',
    'Technical design documentation.',
    'A two-week post-release hyper-support for the internal technology teams',
  ].forEach((line, i) => {
    s.addText('• '+line, { x:6.93, y:3.00+i*0.70, w:4.8, h:0.55, fontSize:13, color:C.white, fontFace:C.bF, wrap:true })
  })

  // Full-width bar (x=0.694, y=4.639, w=11.856, h=1.387, fill=#000000 — template extends it lower)
  // We match: y=5.80 so it sits below the two boxes
  s.addShape('rect', { x:0.69, y:5.76, w:11.86, h:1.38, fill:{color:C.black}, line:{width:0} })
  s.addText('Change Management', { x:0.95, y:5.82, w:5.5, h:0.36, fontSize:20, bold:true, color:C.white, fontFace:C.hF })
  s.addText(
    `We will assist internal stakeholders updating internal processes and existing workflows based on the improved way of working with the new platform to ensure a successful start upon launch.`,
    { x:0.95, y:6.18, w:11.3, h:0.72, fontSize:12, color:C.white, fontFace:C.bF, wrap:true }
  )

  chrome(s, n)
}

// ─── Slide 9: Section divider ─────────────────────────────────────────────────
// Matches slide 15: dark bg, right darker panel, large 60pt green text NOT bold
function s9Divider(prs: pptxgen, text: string, n: number) {
  const s = prs.addSlide()
  s.background = { color: C.dark }
  // Right panel (x=4.384, y=0, w=7.526, h=7.5)
  s.addShape('rect', { x:4.38, y:0, w:8.95, h:7.5, fill:{color:C.darkP}, line:{width:0} })
  // ZONT mark top-right
  s.addText('ZONT', { x:11.11, y:0.50, w:1.61, h:0.41, fontSize:14, bold:true, color:C.white, fontFace:C.hF, align:'right' })
  // Large title (x=0.762, y=2.774, w=9.167, h=2.705, 60pt NOT bold, green)
  s.addText(text, { x:0.76, y:2.77, w:9.17, h:2.70, fontSize:60, bold:false, color:C.gFill, fontFace:C.hF, lineSpacingMultiple:1.05 })
  // Slide num
  s.addShape('rect', { x:12.55, y:7.02, w:0.56, h:0.27, fill:{color:C.tblH}, line:{width:0} })
  s.addText(String(n), { x:12.55, y:7.02, w:0.56, h:0.27, fontSize:9, color:C.white, fontFace:C.bF, align:'center', valign:'middle' })
}

// ─── Slide 10: Gantt timeline ─────────────────────────────────────────────────
// Reproduces template slide 16 faithfully using the parsed phase data
function s10Timeline(prs: pptxgen, p: ProposalContent, n: number) {
  const s = prs.addSlide()
  s.background = { color: C.white }
  H(s, 'Our proposed implementation timeline', 0.36)

  // Week columns — x positions from template XML
  // WEEK 1=0.799, WEEK 2=1.703, WEEK 3=2.607, WEEK 4=3.536, WEEK 5=4.473, WEEK 6=5.459
  // WEEK 7=6.427, WEEK 8=7.44,  WEEK 9=8.451, WEEK 10=9.308, WEEK 11=10.269, WEEK 12=11.143, WEEK 13=11.994
  const weekX = [0.799, 1.703, 2.607, 3.536, 4.473, 5.459, 6.427, 7.44, 8.451, 9.308, 10.269, 11.143, 11.994]
  const WEEK_LABEL_Y = 1.553
  const LINE_Y_START = 1.80
  const LINE_H = 4.95

  // Week labels + vertical grid lines
  weekX.forEach((x, i) => {
    s.addText(`WEEK ${i+1}`, { x, y:WEEK_LABEL_Y, w:0.90, h:0.25, fontSize:9, bold:true, color:C.black, fontFace:C.bF })
    s.addShape('line', { x:x+0.01, y:LINE_Y_START, w:0, h:LINE_H, line:{color:C.line, width:0.5} })
  })

  // Phase rows — derive from parsed sections
  // Map sections to phases
  const sections = p.sheet.sections
  const totalWeeks = Math.max(6, Math.ceil((p.sheet.totalHoursLow + p.sheet.totalHoursHigh) / 2 / 40))

  // Build phases from sections with proportional widths
  interface Phase { label: string; start: number; end: number; row: number; color: string; labelGroup: string }
  const phases: Phase[] = []

  // Group sections into onboarding / implementation / launch phases
  const onboarding  = sections.filter(s => s.toLowerCase().includes('onboard'))
  const launch      = sections.filter(s => ['testing','uat','documentation','deployment'].some(k => s.toLowerCase().includes(k)))
  const impl        = sections.filter(s => !onboarding.includes(s) && !launch.includes(s))

  // Onboarding: weeks 1–2
  if (onboarding.length) {
    phases.push({ label:'Requirements', start:0, end:1, row:0, color:C.gold, labelGroup:'ONBOARDING' })
    phases.push({ label:'Backlog setup', start:1, end:3, row:1, color:C.gold, labelGroup:'' })
    phases.push({ label:'Functional Spec', start:0, end:2, row:2, color:C.gold, labelGroup:'' })
  }

  // Implementation: weeks 3–10
  const implStart = 2, implEnd = Math.min(11, implStart + impl.length + 3)
  if (impl.length) {
    phases.push({ label:'Migration & Development', start:implStart, end:implEnd-2, row:0, color:C.gText, labelGroup:'IMPLEMENTATION' })
    phases.push({ label:'Development', start:implStart, end:implEnd-1, row:1, color:C.gText, labelGroup:'' })
    phases.push({ label:'Testing', start:implStart+2, end:implEnd, row:2, color:C.gText, labelGroup:'' })
    if (sections.some(s=>s.toLowerCase().includes('search')))
      phases.push({ label:'Search Implementation', start:implStart+1, end:implStart+4, row:3, color:C.gText, labelGroup:'' })
    phases.push({ label:'Stabilization', start:implEnd-2, end:implEnd, row:4, color:C.gText, labelGroup:'' })
    phases.push({ label:'Documentation', start:implEnd-3, end:implEnd-1, row:5, color:C.gText, labelGroup:'' })
  }

  // UAT & Launch: last 3 weeks
  const uatStart = Math.max(8, totalWeeks - 3)
  phases.push({ label:'Handoff', start:uatStart, end:uatStart+1, row:0, color:C.orange, labelGroup:'UAT & LAUNCH' })
  phases.push({ label:'UAT Support', start:uatStart+1, end:uatStart+2, row:1, color:C.orange, labelGroup:'' })
  phases.push({ label:'Go-live', start:uatStart+2, end:uatStart+3, row:2, color:C.orange, labelGroup:'' })
  if (sections.some(s=>s.toLowerCase().includes('doc')||s.toLowerCase().includes('train')))
    phases.push({ label:'Training', start:uatStart+1, end:uatStart+3, row:3, color:C.orange, labelGroup:'' })

  // Phase group labels (ONBOARDING, IMPLEMENTATION, UAT & LAUNCH)
  const groupLabels: {label:string; brace:string; y:number; braceY:number; braceH:number}[] = [
    { label:'ONBOARDING',    brace:'Right Brace', y:3.14, braceY:2.16, braceH:0.97 },
    { label:'IMPLEMENTATION',brace:'Right Brace', y:4.08, braceY:3.32, braceH:1.70 },
    { label:'UAT & LAUNCH',  brace:'Right Brace', y:4.83, braceY:4.12, braceH:1.26 },
  ]
  groupLabels.forEach(g => {
    s.addText(g.label, { x:0.25, y:g.y, w:1.4, h:0.25, fontSize:9, bold:true, color:C.black, fontFace:C.bF })
  })

  // Phase bars
  const BAR_Y0 = 2.15   // y of first phase row
  const ROW_H  = 0.28
  const ROW_GAP = 0.02

  phases.forEach(ph => {
    const x1 = weekX[Math.min(ph.start, weekX.length-1)]
    const x2 = ph.end < weekX.length ? weekX[ph.end] : weekX[weekX.length-1] + 0.90
    const bw  = Math.max(x2 - x1, 0.3)
    const by  = BAR_Y0 + ph.row * (ROW_H + ROW_GAP)

    s.addShape('rect', { x:x1, y:by, w:bw, h:ROW_H, fill:{color:ph.color}, line:{width:0}, rounding:0.08 })
    if (bw > 0.5) {
      s.addText(ph.label, { x:x1+0.06, y:by+0.03, w:bw-0.08, h:ROW_H-0.04, fontSize:8, color:C.white, fontFace:C.bF, wrap:false })
    }
  })

  chrome(s, n)
}

// ─── Slide 11: Team structure ─────────────────────────────────────────────────
// Matches template slide 17: org chart with client left + delivery team right
function s11Team(prs: pptxgen, p: ProposalContent, n: number) {
  const s = prs.addSlide()
  s.background = { color: C.white }
  H(s, 'Our recommended team structure')

  // Left column — client side (dark boxes, matching template)
  const CLIENT_X = 0.832, BOX_W = 1.745, BOX_H = 0.467, BOX_GAP = 0.09

  s.addText(`${p.clientName} Marketing`, { x:CLIENT_X, y:1.195, w:BOX_W, h:0.40, fontSize:14, bold:true, color:C.black, fontFace:C.hF })
  ;['Project Sponsor', 'Marketing Lead', 'Content Author'].forEach((name, i) => {
    s.addShape('rect', { x:CLIENT_X, y:2.0+i*(BOX_H+BOX_GAP), w:BOX_W, h:BOX_H, fill:{color:C.tblH}, line:{width:0} })
    s.addText(name, { x:CLIENT_X, y:2.0+i*(BOX_H+BOX_GAP), w:BOX_W, h:BOX_H, fontSize:10, bold:true, color:C.white, fontFace:C.bF, align:'center', valign:'middle' })
  })
  s.addText(`${p.clientName} IT`, { x:CLIENT_X, y:3.40, w:BOX_W, h:0.38, fontSize:14, bold:true, color:C.black, fontFace:C.hF })
  ;['IT Lead', 'DevOps'].forEach((name, i) => {
    s.addShape('rect', { x:CLIENT_X, y:3.90+i*(BOX_H+BOX_GAP), w:BOX_W, h:BOX_H, fill:{color:C.tblH}, line:{width:0} })
    s.addText(name, { x:CLIENT_X, y:3.90+i*(BOX_H+BOX_GAP), w:BOX_W, h:BOX_H, fontSize:10, bold:true, color:C.white, fontFace:C.bF, align:'center', valign:'middle' })
  })

  // Arrow in middle
  s.addShape('rect', { x:2.85, y:3.38, w:1.67, h:0.73, fill:{color:C.white}, line:{color:C.line, width:0.5} })
  s.addText('← collaboration →', { x:2.85, y:3.38, w:1.67, h:0.73, fontSize:9, color:C.zGrey, fontFace:C.bF, align:'center', valign:'middle' })

  // Right side — Zont Digital delivery team
  s.addText('Zont Digital', { x:7.883, y:1.195, w:4.5, h:0.40, fontSize:16, bold:true, color:C.black, fontFace:C.hF })

  // Oversight (green)
  s.addShape('rect', { x:10.168, y:1.90, w:2.015, h:2.30, fill:{color:C.white}, line:{color:C.line, width:0.5} })
  s.addText('Oversight Team', { x:10.168, y:1.94, w:2.015, h:0.32, fontSize:10, color:C.zGrey, fontFace:C.bF, align:'center' })
  s.addShape('rect', { x:10.524, y:2.30, w:1.498, h:0.70, fill:{color:C.gText}, line:{width:0} })
  s.addText('Engagement\nLead, Governance', { x:10.524, y:2.30, w:1.498, h:0.70, fontSize:10, color:C.white, fontFace:C.bF, align:'center', valign:'middle' })
  s.addShape('rect', { x:10.524, y:3.10, w:1.498, h:0.70, fill:{color:C.orange}, line:{width:0} })
  s.addText('Delivery\nGovernance', { x:10.524, y:3.10, w:1.498, h:0.70, fontSize:10, color:C.white, fontFace:C.bF, align:'center', valign:'middle' })

  // PM
  s.addShape('rect', { x:7.958, y:2.14, w:1.3, h:0.70, fill:{color:C.orange}, line:{width:0} })
  s.addText('PM', { x:7.958, y:2.14, w:1.3, h:0.70, fontSize:14, bold:true, color:C.white, fontFace:C.hF, align:'center', valign:'middle' })

  // Delivery roles (green + orange alternating)
  const roles = [
    { title:'Architect',           x:8.00, y:3.49, color:C.gText },
    { title:'Business Analyst',    x:5.15, y:3.46, color:C.orange },
    { title:'Tester',              x:6.48, y:3.46, color:C.orange },
    { title:'Frontend Developer',  x:7.28, y:4.62, color:C.orange },
    { title:'Full-stack Developer',x:8.79, y:4.62, color:C.orange },
  ]
  roles.forEach(r => {
    s.addShape('rect', { x:r.x, y:r.y, w:1.216, h:0.73, fill:{color:r.color}, line:{width:0} })
    s.addText(r.title, { x:r.x, y:r.y, w:1.216, h:0.73, fontSize:10, color:C.white, fontFace:C.bF, align:'center', valign:'middle', wrap:true })
  })

  // Legend
  s.addShape('rect', { x:9.627, y:7.008, w:0.289, h:0.28, fill:{color:C.gText}, line:{width:0} })
  s.addText('USA', { x:9.93, y:7.01, w:0.9, h:0.28, fontSize:11, color:C.black, fontFace:C.bF })
  s.addShape('rect', { x:11.02, y:7.029, w:0.289, h:0.28, fill:{color:C.orange}, line:{width:0} })
  s.addText('Latvia', { x:11.37, y:7.02, w:0.9, h:0.28, fontSize:11, color:C.black, fontFace:C.bF })

  chrome(s, n)
}

// ─── Slide 12: Investment table ───────────────────────────────────────────────
// Matches template slide 18 exactly
function s12Investment(prs: pptxgen, p: ProposalContent, n: number) {
  const s = prs.addSlide()
  s.background = { color: C.white }
  // Title (slide 18: x=0.418, y=0.357, w=9.898, 24pt bold)
  s.addText('Investment summary', { x:0.42, y:0.36, w:9.90, h:0.40, fontSize:24, bold:true, color:C.black, fontFace:C.hF })

  const sheet    = p.sheet
  const discount = 15000
  const rebate   = 25000
  const midC     = sheet.midCost
  const net      = midC - discount - rebate
  const wks      = Math.max(1, Math.round((sheet.totalHoursLow + sheet.totalHoursHigh) / 2 / 40))

  const TX = 0.42, TY = 0.88, TW = 12.38
  const CW = [7.50, 1.80, 3.08]

  // Header row — dark #2E2E2E, white text 12pt, h=0.344 (from template)
  s.addShape('rect', { x:TX, y:TY, w:TW, h:0.34, fill:{color:C.tblH}, line:{width:0} })
  let cx = TX
  ;['ACTIVITY', 'TIMELINE', 'COST (USD)'].forEach((h, i) => {
    s.addText(h, { x:cx+0.12, y:TY, w:CW[i]-0.12, h:0.34, fontSize:12, color:C.white, fontFace:C.hF, valign:'middle' })
    cx += CW[i]
  })

  // Row 1: main project (h=1.743)
  const r1Y = TY+0.34, r1H = 1.74
  s.addShape('rect', { x:TX, y:r1Y, w:TW, h:r1H, fill:{color:C.white}, line:{color:C.line, width:0.5} })
  s.addText(p.projectName, { x:TX+0.12, y:r1Y+0.14, w:CW[0]-0.24, h:0.34, fontSize:12, bold:true, color:C.black, fontFace:C.hF })
  s.addText(sheet.sections.slice(0,7).join('  ·  '), {
    x:TX+0.12, y:r1Y+0.52, w:CW[0]-0.24, h:1.10, fontSize:9, color:C.body, fontFace:C.bF, wrap:true, lineSpacingMultiple:1.3,
  })
  s.addText(`${wks} weeks`, { x:TX+CW[0]+0.12, y:r1Y+0.14, w:CW[1]-0.12, h:0.36, fontSize:12, color:C.body, fontFace:C.bF })
  s.addText(`$${midC.toLocaleString()}`, { x:TX+CW[0]+CW[1]+0.12, y:r1Y+0.14, w:CW[2]-0.12, h:0.36, fontSize:12, color:C.black, fontFace:C.bF })

  // Row 2: rebate (h=1.172)
  const r2Y = r1Y+r1H, r2H = 0.60
  s.addShape('rect', { x:TX, y:r2Y, w:TW, h:r2H, fill:{color:C.tblAlt}, line:{color:C.line, width:0.5} })
  s.addText('Sitecore Commercial Migration Rebate\nRebate upon migration to XM Cloud', {
    x:TX+0.12, y:r2Y+0.06, w:CW[0]-0.24, h:0.48, fontSize:10, color:C.body, fontFace:C.bF, lineSpacingMultiple:1.2,
  })
  s.addText(`($${rebate.toLocaleString()})`, { x:TX+CW[0]+CW[1]+0.12, y:r2Y+0.12, w:CW[2]-0.12, h:0.36, fontSize:11, color:C.body, fontFace:C.bF })

  // Row 3: discount (h=0.855)
  const r3Y = r2Y+r2H, r3H = 0.56
  s.addShape('rect', { x:TX, y:r3Y, w:TW, h:r3H, fill:{color:C.white}, line:{color:C.line, width:0.5} })
  s.addText('Marketing Discount\nClient reference, joint presentation or webinar, sharable feedback', {
    x:TX+0.12, y:r3Y+0.06, w:CW[0]-0.24, h:0.46, fontSize:10, color:C.body, fontFace:C.bF, lineSpacingMultiple:1.2,
  })
  s.addText(`($${discount.toLocaleString()})`, { x:TX+CW[0]+CW[1]+0.12, y:r3Y+0.10, w:CW[2]-0.12, h:0.36, fontSize:11, color:C.body, fontFace:C.bF })

  // Total row — dark bg (h=0.579)
  const totY = r3Y+r3H
  s.addShape('rect', { x:TX, y:totY, w:TW, h:0.52, fill:{color:C.tblH}, line:{width:0} })
  s.addText('Total', { x:TX+0.12, y:totY, w:CW[0]-0.12, h:0.52, fontSize:14, bold:true, color:C.white, fontFace:C.hF, valign:'middle' })
  s.addText(`${wks} weeks`, { x:TX+CW[0]+0.12, y:totY, w:CW[1]-0.12, h:0.52, fontSize:12, color:C.white, fontFace:C.bF, valign:'middle' })
  s.addText(`$${net.toLocaleString()}`, { x:TX+CW[0]+CW[1]+0.12, y:totY, w:CW[2]-0.12, h:0.52, fontSize:14, bold:true, color:C.gFill, fontFace:C.hF, valign:'middle' })

  // Footer (slide 18: TextBox 3 x=1.454, y=6.795, 8pt)
  s.addText(
    'The quote is given as a fixed bid; Sitecore rebate contingent on confirmation from Sitecore.\nExcludes T&E expenses if any.',
    { x:1.45, y:6.79, w:10.0, h:0.46, fontSize:8, color:C.zGrey, fontFace:C.bF, italic:true, lineSpacingMultiple:1.3 }
  )

  chrome(s, n)
}

// ─── Slide 13: Closing ────────────────────────────────────────────────────────
// Matches slide 21 exactly
function s13Closing(prs: pptxgen, n: number) {
  const s = prs.addSlide()
  s.background = { color: C.dark }
  s.addShape('rect', { x:4.38, y:0, w:8.95, h:7.5, fill:{color:C.darkP}, line:{width:0} })
  s.addText('ZONT', { x:11.11, y:0.50, w:1.61, h:0.41, fontSize:14, bold:true, color:C.white, fontFace:C.hF, align:'right' })
  // Title: x=0.762, y=2.774, w=9.167, h=2.705, 60pt NOT bold, green fill colour
  s.addText("Let's create remarkable\ndigital solutions!", {
    x:0.76, y:2.77, w:9.17, h:2.70, fontSize:60, bold:false, color:C.gFill, fontFace:C.hF, lineSpacingMultiple:1.05,
  })
  s.addShape('rect', { x:12.55, y:7.02, w:0.56, h:0.27, fill:{color:C.tblH}, line:{width:0} })
  s.addText(String(n), { x:12.55, y:7.02, w:0.56, h:0.27, fontSize:9, color:C.white, fontFace:C.bF, align:'center', valign:'middle' })
}

// ─── Main export ──────────────────────────────────────────────────────────────
export async function generatePptx(p: ProposalContent): Promise<Blob> {
  const prs = new pptxgen()
  prs.layout = 'LAYOUT_WIDE'  // 13.33" × 7.5"

  s1Cover(prs, p)                                               //  1
  s2Agenda(prs, p, 2)                                           //  2
  s3Objectives(prs, p, 3)                                       //  3
  s4Tenants(prs, p, 4)                                          //  4
  s5Stats(prs, p, 5)                                            //  5
  s6Scope(prs, p, 6)                                            //  6
  s7QA(prs, p, 7)                                               //  7
  s8Enablement(prs, p, 8)                                       //  8
  s9Divider(prs, 'Timeline,\nteam,\ninvestment\nsummary', 9)    //  9
  s10Timeline(prs, p, 10)                                       // 10
  s11Team(prs, p, 11)                                           // 11
  s12Investment(prs, p, 12)                                     // 12
  s13Closing(prs, 13)                                           // 13

  const b64 = await prs.write({ outputType: 'base64' }) as string
  const bin = atob(b64)
  const bytes = new Uint8Array(bin.length)
  for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i)
  return new Blob([bytes], { type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation' })
}
