import { create } from 'zustand'
import type { SheetSummary } from '../services/excelParser'

export interface ProposalContent {
  projectName: string
  clientName: string
  date: string
  needs: string[]
  goals: string[]
  assumptions: string[]
  sheet: SheetSummary
}

export type Step = 'upload' | 'preview' | 'proposal' | 'generate' | 'contract'

interface AppState {
  step: Step
  done: Set<Step>
  sheet: SheetSummary | null
  proposal: ProposalContent | null
  generatedPptx: Blob | null
  tweakedPptx: Blob | null
  contractDocx: Blob | null
  contractMd: string | null
  isGeneratingPptx: boolean
  isGeneratingContract: boolean
  error: string | null

  setStep: (s: Step) => void
  markDone: (s: Step) => void
  setSheet: (s: SheetSummary) => void
  setProposal: (p: ProposalContent) => void
  patchProposal: (p: Partial<ProposalContent>) => void
  setGeneratedPptx: (b: Blob) => void
  setTweakedPptx: (b: Blob) => void
  setContract: (docx: Blob, md: string) => void
  setGeneratingPptx: (v: boolean) => void
  setGeneratingContract: (v: boolean) => void
  setError: (e: string | null) => void
}

export const useStore = create<AppState>((set) => ({
  step: 'upload',
  done: new Set(),
  sheet: null,
  proposal: null,
  generatedPptx: null,
  tweakedPptx: null,
  contractDocx: null,
  contractMd: null,
  isGeneratingPptx: false,
  isGeneratingContract: false,
  error: null,

  setStep: (s) => set({ step: s }),
  markDone: (s) => set((st) => ({ done: new Set([...st.done, s]) })),
  setSheet: (s) => set({ sheet: s }),
  setProposal: (p) => set({ proposal: p }),
  patchProposal: (p) => set((st) => ({ proposal: st.proposal ? { ...st.proposal, ...p } : null })),
  setGeneratedPptx: (b) => set({ generatedPptx: b }),
  setTweakedPptx: (b) => set({ tweakedPptx: b }),
  setContract: (docx, md) => set({ contractDocx: docx, contractMd: md }),
  setGeneratingPptx: (v) => set({ isGeneratingPptx: v }),
  setGeneratingContract: (v) => set({ isGeneratingContract: v }),
  setError: (e) => set({ error: e }),
}))
