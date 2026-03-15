import type { SheetSummary } from './excelParser'
import type { ProposalContent } from '../store/appStore'

export function buildProposal(sheet: SheetSummary, projectName = '', clientName = ''): ProposalContent {
  const sections = sheet.sections.map((s) => s.toLowerCase())
  const taskNames = sheet.tasks.map((t) => t.task.toLowerCase()).join(' ')

  const needs: string[] = []

  if (sections.includes('onboarding') || taskNames.includes('discovery'))
    needs.push('Establish a structured onboarding process to align stakeholders and define scope before development begins')
  if (taskNames.includes('component') || taskNames.includes('migration'))
    needs.push('Migrate and convert all existing components and content to the new platform with parity and best-practice compliance')
  if (taskNames.includes('search'))
    needs.push('Deliver a modern, performant search experience that replaces the legacy Solr implementation with AI-assisted recommendations')
  if (taskNames.includes('form'))
    needs.push('Recreate and integrate all existing forms with proper validation, workflow actions, and third-party system integrations')
  if (taskNames.includes('security') || taskNames.includes('b2c') || taskNames.includes('sso'))
    needs.push('Ensure enterprise-grade security through Azure B2C identity management, SSO enablement, and OWASP-compliant configuration')
  if (taskNames.includes('gtm') || taskNames.includes('analytics') || taskNames.includes('integration'))
    needs.push('Integrate marketing and analytics tooling (GTM, GetResponse, Zoho, Agiloft) for seamless data flow and campaign tracking')
  if (taskNames.includes('ada') || taskNames.includes('508') || taskNames.includes('compliance'))
    needs.push('Maintain ADA 508 and WCAG 2.1 accessibility compliance throughout the migration to protect legal standing and user inclusivity')
  if (taskNames.includes('uat') || taskNames.includes('regression') || taskNames.includes('test'))
    needs.push('Validate all deliverables through structured QA, regression testing, and a client-led UAT cycle before go-live')
  if (taskNames.includes('training') || taskNames.includes('handoff') || taskNames.includes('documentation'))
    needs.push('Enable internal marketing and technology teams through structured training, documentation, and a post-launch hyper-support window')

  if (needs.length < 3)
    needs.push('Deliver a production-ready, scalable solution with full knowledge transfer and post-launch support')

  const mid = sheet.midCost
  const goals: string[] = [
    `Complete the full migration within the agreed timeline at a fixed investment of approximately $${Math.round(mid / 1000)}k`,
    'Preserve all existing site functionality while adopting modern headless architecture on Sitecore XM Cloud',
    'Enable marketing self-service through WYSIWYG authoring, reducing dependency on developer intervention for content changes',
    'Achieve measurable quality benchmarks — zero critical defects at go-live, full regression pass, and ADA 508 compliance verified',
    'Ensure all team members are fully enabled on the new platform before the hyper-support period ends',
  ]

  const assumptions: string[] = [
    'Lift & Shift approach to reach XM Cloud as quickly as possible — no redesign in scope',
    'Media Library and external blob storage configuration will be retained; no Content Hub migration',
    'Existing xDB data will not be migrated as it does not provide value in the new architecture',
    'Personalization is out of scope for this engagement',
    'Figma designs will be provided by the client prior to component development',
    'Functional regression testing will be conducted by the client team',
    'OneTrust integration will use existing code drop with no net-new configuration',
  ]

  return {
    projectName: projectName || 'XP to XM Cloud Migration',
    clientName: clientName || 'Lattice Semiconductor',
    date: new Date().toLocaleDateString('en-GB', { year: 'numeric', month: 'long', day: 'numeric' }),
    needs,
    goals,
    assumptions,
    sheet,
  }
}
