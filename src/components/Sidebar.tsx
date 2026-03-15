import React from 'react'
import type { Step } from '../store/appStore'

const STEPS: { id: Step; label: string; num: string }[] = [
  { id: 'upload',   label: 'Upload Estimate',    num: '01' },
  { id: 'preview',  label: 'Data Preview',        num: '02' },
  { id: 'proposal', label: 'Proposal Editor',     num: '03' },
  { id: 'generate', label: 'Generate & Download', num: '04' },
  { id: 'contract', label: 'SLW Contract',        num: '05' },
]

interface Props { current: Step; done: Set<Step>; onNav: (s: Step) => void }

export function Sidebar({ current, done, onNav }: Props) {
  return (
    <aside className="w-52 bg-surface-card border-r border-border flex flex-col sticky top-0 h-screen shrink-0">
      <div className="px-5 py-5 border-b border-border">
        <div className="font-display text-[17px] font-semibold text-brand-green tracking-tight">EstimateForge</div>
        <div className="font-mono text-[10px] text-ink-3 uppercase tracking-widest mt-0.5">Proposal Pipeline</div>
      </div>
      <nav className="flex-1 py-3">
        {STEPS.map((step) => {
          const active = current === step.id
          const isDone = done.has(step.id) && !active
          return (
            <button key={step.id} onClick={() => onNav(step.id)}
              className={`w-full flex items-center gap-2.5 px-5 py-[10px] text-left transition-all text-[13px] border-l-[2.5px] ${active ? 'border-brand-green bg-brand-green-lt text-brand-green font-medium' : 'border-transparent text-ink-2 hover:bg-surface hover:text-ink font-normal'}`}>
              <span className={`w-4 h-4 rounded-full flex items-center justify-center shrink-0 text-[9px] font-mono font-medium ${active ? 'bg-brand-green text-white' : isDone ? 'bg-brand-green text-white' : 'bg-border text-ink-3'}`}>
                {isDone ? '✓' : step.num.slice(1)}
              </span>
              <span className="flex-1">{step.label}</span>
            </button>
          )
        })}
      </nav>
      <div className="px-5 py-4 border-t border-border">
        <div className="font-mono text-[10px] text-border-strong uppercase tracking-wider">v1.0.0</div>
      </div>
    </aside>
  )
}
