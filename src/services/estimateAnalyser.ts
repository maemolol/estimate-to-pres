import type { EstimateRow, ProposalContent } from '../store/appStore'

export function groupByPhase(rows: EstimateRow[]): Record<string, EstimateRow[]> {
  return rows.reduce<Record<string, EstimateRow[]>>((acc, row) => {
    if (!acc[row.phase]) acc[row.phase] = []
    acc[row.phase].push(row)
    return acc
  }, {})
}

export function totalHours(rows: EstimateRow[]): number {
  return rows.reduce((s, r) => s + r.hours, 0)
}

export function totalCost(rows: EstimateRow[]): number {
  return rows.reduce((s, r) => s + r.cost, 0)
}

export function deriveNeeds(rows: EstimateRow[]): string[] {
  const needs: string[] = []
  const components = rows.map((r) => r.component.toLowerCase())

  if (components.some((c) => c.includes('ui') || c.includes('frontend') || c.includes('web')))
    needs.push('Modernise the digital interface with a performant, user-centric frontend experience')
  if (components.some((c) => c.includes('api') || c.includes('backend') || c.includes('server')))
    needs.push('Establish a scalable API layer to support current and future integration requirements')
  if (components.some((c) => c.includes('database') || c.includes('db') || c.includes('schema')))
    needs.push('Centralise data management with a well-structured, maintainable data architecture')
  if (components.some((c) => c.includes('admin') || c.includes('dashboard') || c.includes('portal')))
    needs.push('Provide internal stakeholders with self-service visibility and content management tools')
  if (components.some((c) => c.includes('analytics') || c.includes('report') || c.includes('kpi')))
    needs.push('Enable data-driven decision-making through real-time analytics and reporting')
  if (components.some((c) => c.includes('qa') || c.includes('test') || c.includes('quality')))
    needs.push('Ensure release confidence through automated quality assurance and regression testing')
  if (components.some((c) => c.includes('devops') || c.includes('ci') || c.includes('deploy')))
    needs.push('Reduce deployment friction and operational overhead via automated CI/CD pipelines')
  if (components.some((c) => c.includes('migration') || c.includes('migrat')))
    needs.push('Safely migrate existing content and configuration with minimal business disruption')
  if (components.some((c) => c.includes('search')))
    needs.push('Improve content discoverability and user engagement through intelligent search')

  // Fallback if no matches
  if (!needs.length) {
    needs.push('Deliver a high-quality, production-ready solution aligned with business objectives')
    needs.push('Ensure technical excellence, maintainability, and long-term scalability')
  }

  return needs
}

export function deriveGoals(rows: EstimateRow[], phases: Record<string, EstimateRow[]>): string[] {
  const phaseNames = Object.keys(phases)
  const total = totalCost(rows)
  const hours = totalHours(rows)
  const goals: string[] = [
    `Deliver a production-ready solution across ${phaseNames.length} phase${phaseNames.length > 1 ? 's' : ''} within the agreed timeline`,
    `Maintain full budget transparency against the $${total.toLocaleString()} fixed-bid estimate`,
    `Achieve measurable quality benchmarks through structured QA at each phase milestone`,
  ]
  if (hours > 200)
    goals.push('Coordinate a cross-functional team with clear ownership and weekly progress reporting')
  goals.push('Ensure knowledge transfer and enablement so the client team can operate independently post-launch')
  return goals
}

export function buildProposalFromEstimate(
  rows: EstimateRow[],
  projectName = 'New Project',
  clientName = 'Client'
): ProposalContent {
  const phases = groupByPhase(rows)
  return {
    projectName,
    clientName,
    date: new Date().toLocaleDateString('en-GB', { year: 'numeric', month: 'long', day: 'numeric' }),
    needs: deriveNeeds(rows),
    goals: deriveGoals(rows, phases),
    phases,
  }
}
