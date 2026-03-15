import { Document, Paragraph, TextRun, HeadingLevel, AlignmentType, BorderStyle, Table, TableRow, TableCell, WidthType, Packer, ShadingType } from 'docx'
import type { ProposalContent } from '../store/appStore'

export interface ContractSections {
  projectOverview: string
  scopeOfWork: string
  deliverables: string
  timeline: string
  paymentTerms: string
  responsibilities: string
  assumptions: string
  acceptanceCriteria: string
  changeManagement: string
}

export const SECTION_LABELS: { key: keyof ContractSections; title: string }[] = [
  { key: 'projectOverview',    title: '1. Project Overview' },
  { key: 'scopeOfWork',        title: '2. Scope of Work' },
  { key: 'deliverables',       title: '3. Deliverables' },
  { key: 'timeline',           title: '4. Timeline' },
  { key: 'paymentTerms',       title: '5. Payment Terms' },
  { key: 'responsibilities',   title: '6. Responsibilities' },
  { key: 'assumptions',        title: '7. Assumptions & Exclusions' },
  { key: 'acceptanceCriteria', title: '8. Acceptance Criteria' },
  { key: 'changeManagement',   title: '9. Change Management' },
]

// ─── GPT-4 call ───────────────────────────────────────────────────────────────

export async function generateContractSections(
  proposal: ProposalContent,
  apiKey: string
): Promise<ContractSections> {
  const sheet = proposal.sheet
  const mid = sheet.midCost
  const discount = 15000
  const rebate = 25000
  const net = mid - discount - rebate

  const sectionList = [...new Map(
    sheet.tasks.map((t) => [t.section.split(' › ')[0], true])
  ).keys()].join(', ')

  const taskList = sheet.tasks
    .slice(0, 20)
    .map((t) => `- ${t.task} (${t.totalHoursLow}–${t.totalHoursHigh} hrs)`)
    .join('\n')

  const roleList = sheet.roles
    .map((r) => `${r.role}: ${r.hoursLow}–${r.hoursHigh} hrs @ $${sheet.tasks.find(t=>t.roles.find(ro=>ro.role===r.role))?.roles.find(ro=>ro.role===r.role)?.rate ?? 'variable'}/hr`)
    .join('\n')

  const prompt = `You are a senior technology contracts lawyer specialising in digital transformation projects. Generate a detailed, professional Statement of Work (SOW) contract.

PROJECT DETAILS:
- Project: ${proposal.projectName}
- Client: ${proposal.clientName}
- Date: ${proposal.date}
- Gross mid-point estimate: $${mid.toLocaleString()}
- Client discount applied: $${discount.toLocaleString()}
- Sitecore migration rebate: $${rebate.toLocaleString()}
- Net total investment: $${net.toLocaleString()}
- Total hours (mid): ~${Math.round((sheet.totalHoursLow+sheet.totalHoursHigh)/2)} hours

WORK STREAMS: ${sectionList}

KEY TASKS (sample):
${taskList}

ROLE ALLOCATION:
${roleList}

PROJECT NEEDS:
${proposal.needs.map(n=>`- ${n}`).join('\n')}

GOALS:
${proposal.goals.map(g=>`- ${g}`).join('\n')}

PROJECT ASSUMPTIONS:
${proposal.assumptions.map(a=>`- ${a}`).join('\n')}

Respond ONLY with a valid JSON object — no markdown, no code fences, no preamble — with exactly these keys:
{
  "projectOverview": "...",
  "scopeOfWork": "...",
  "deliverables": "...",
  "timeline": "...",
  "paymentTerms": "...",
  "responsibilities": "...",
  "assumptions": "...",
  "acceptanceCriteria": "...",
  "changeManagement": "..."
}

Requirements:
- Write in formal legal English appropriate for a signed commercial contract
- Each section must be 3–5 substantive paragraphs
- Use the exact project name, client name, and dollar figures provided
- paymentTerms must specify a 30/30/30/10 payment schedule with the actual $ amounts (based on net investment of $${net.toLocaleString()})
- deliverables must enumerate specific outputs per work stream
- assumptions must incorporate all listed project assumptions verbatim or expanded`

  const res = await fetch('https://api.openai.com/v1/chat/completions', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json', Authorization: `Bearer ${apiKey}` },
    body: JSON.stringify({
      model: 'gpt-4-turbo-preview',
      messages: [{ role: 'user', content: prompt }],
      temperature: 0.25,
      max_tokens: 4096,
      response_format: { type: 'json_object' },
    }),
  })

  if (!res.ok) {
    const err = await res.json().catch(() => ({ error: { message: res.statusText } }))
    throw new Error(`OpenAI API error: ${err?.error?.message ?? res.statusText}`)
  }

  const data = await res.json()
  const raw = data.choices?.[0]?.message?.content ?? '{}'
  try { return JSON.parse(raw) as ContractSections }
  catch { throw new Error('Failed to parse GPT-4 response. Please try again.') }
}

// ─── DOCX builder ─────────────────────────────────────────────────────────────

function h1(text: string) {
  return new Paragraph({ text, heading: HeadingLevel.HEADING_1, spacing: { before: 480, after: 120 } })
}
function body(text: string) {
  return new Paragraph({
    children: [new TextRun({ text, size: 22, font: 'Calibri' })],
    spacing: { after: 180, line: 330 },
  })
}
function divider() {
  return new Paragraph({
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: '1A5C3A' } },
    spacing: { after: 240 },
    children: [],
  })
}

export async function buildDocx(proposal: ProposalContent, secs: ContractSections): Promise<Blob> {
  const sheet = proposal.sheet
  const mid = sheet.midCost
  const net = mid - 15000 - 25000

  const doc = new Document({
    styles: {
      default: { document: { run: { font: 'Calibri', size: 22 } } },
      paragraphStyles: [
        { id: 'Heading1', name: 'Heading 1', run: { font: 'Calibri', size: 28, bold: true, color: '1A1A18' } },
      ],
    },
    sections: [{
      properties: { page: { margin: { top: 1080, bottom: 1080, left: 1200, right: 1200 } } },
      children: [
        // Cover
        new Paragraph({ children: [new TextRun({ text: 'STATEMENT OF WORK', size: 40, bold: true })], alignment: AlignmentType.CENTER, spacing: { before: 600, after: 120 } }),
        new Paragraph({ children: [new TextRun({ text: proposal.projectName, size: 32, bold: true, color: '1A5C3A' })], alignment: AlignmentType.CENTER, spacing: { after: 80 } }),
        new Paragraph({ children: [new TextRun({ text: `Prepared for: ${proposal.clientName}`, size: 22, color: '555550' })], alignment: AlignmentType.CENTER, spacing: { after: 40 } }),
        new Paragraph({ children: [new TextRun({ text: `Date: ${proposal.date}`, size: 22, color: '555550' })], alignment: AlignmentType.CENTER, spacing: { after: 40 } }),
        new Paragraph({ children: [new TextRun({ text: `Net Investment: $${net.toLocaleString()}`, size: 24, bold: true, color: '1A5C3A' })], alignment: AlignmentType.CENTER, spacing: { after: 600 } }),
        divider(),
        // Sections
        ...SECTION_LABELS.flatMap(({ title, key }) => [
          h1(title),
          body(secs[key]),
          divider(),
        ]),
        // Signature block
        new Paragraph({ children: [new TextRun({ text: 'Signatures', size: 28, bold: true })], spacing: { before: 400, after: 200 } }),
        new Paragraph({ children: [new TextRun({ text: 'By signing below, both parties agree to the terms of this Statement of Work.', size: 20, italics: true, color: '888880' })], spacing: { after: 320 } }),
        new Table({
          width: { size: 100, type: WidthType.PERCENTAGE },
          rows: [new TableRow({
            children: [
              new TableCell({
                width: { size: 50, type: WidthType.PERCENTAGE },
                shading: { type: ShadingType.CLEAR, fill: 'F8F8F5' },
                children: [
                  new Paragraph({ children: [new TextRun({ text: 'Client', bold: true, size: 20 })], spacing: { after: 60 } }),
                  new Paragraph({ children: [new TextRun({ text: proposal.clientName, size: 18, color: '555550' })], spacing: { after: 60 } }),
                  new Paragraph({ border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: '1A5C3A' } }, children: [new TextRun({ text: ' ', size: 18 })], spacing: { before: 480, after: 60 } }),
                  new Paragraph({ children: [new TextRun({ text: 'Authorised Signatory', size: 16, color: '888880' })], spacing: { after: 40 } }),
                  new Paragraph({ children: [new TextRun({ text: 'Date: ___________________', size: 16, color: '888880' })], spacing: { after: 40 } }),
                ],
              }),
              new TableCell({
                width: { size: 50, type: WidthType.PERCENTAGE },
                shading: { type: ShadingType.CLEAR, fill: 'F8F8F5' },
                children: [
                  new Paragraph({ children: [new TextRun({ text: 'Service Provider', bold: true, size: 20 })], spacing: { after: 60 } }),
                  new Paragraph({ children: [new TextRun({ text: 'Zont Digital', size: 18, color: '555550' })], spacing: { after: 60 } }),
                  new Paragraph({ border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: '1A5C3A' } }, children: [new TextRun({ text: ' ', size: 18 })], spacing: { before: 480, after: 60 } }),
                  new Paragraph({ children: [new TextRun({ text: 'Authorised Signatory', size: 16, color: '888880' })], spacing: { after: 40 } }),
                  new Paragraph({ children: [new TextRun({ text: 'Date: ___________________', size: 16, color: '888880' })], spacing: { after: 40 } }),
                ],
              }),
            ],
          })],
        }),
      ],
    }],
  })

  const buf = await Packer.toBlob(doc)
  return buf
}

// ─── Markdown export ──────────────────────────────────────────────────────────

export function buildMarkdown(proposal: ProposalContent, secs: ContractSections): string {
  const net = proposal.sheet.midCost - 15000 - 25000
  return `# Statement of Work\n\n**Project:** ${proposal.projectName}\n**Client:** ${proposal.clientName}\n**Date:** ${proposal.date}\n**Net Investment:** $${net.toLocaleString()}\n\n---\n\n` +
    SECTION_LABELS.map(({ title, key }) => `## ${title}\n\n${secs[key]}\n`).join('\n---\n\n') +
    `\n---\n\n## Signatures\n\n**Client:** ___________________________________ **Date:** ___________\n\nAuthorised Signatory: ___________________________\n\n**Service Provider:** ___________________________________ **Date:** ___________\n\nAuthorised Signatory: ___________________________\n`
}
