import React, { useState } from 'react'
import { useStore } from '../store/appStore'
import { fmtUSD } from '../utils/download'

const PHASE_COLORS = ['#1a5c3a','#1a3a5c','#b85c00','#6d2e46','#2d6b6b','#5c4a1a','#1a4a5c']

export function ProposalPage() {
  const { proposal, patchProposal, setStep, markDone } = useStore()
  const [editNeed, setEditNeed] = useState<number | null>(null)
  const [editGoal, setEditGoal] = useState<number | null>(null)
  const [editAssumption, setEditAssumption] = useState<number | null>(null)
  const [newNeed, setNewNeed] = useState('')
  const [newGoal, setNewGoal] = useState('')

  if (!proposal) return (
    <div className="flex flex-col items-center justify-center h-96 text-center">
      <p className="text-ink-2 font-medium">No proposal data — upload an estimate first</p>
      <button className="btn btn-primary mt-4" onClick={() => setStep('upload')}>← Upload</button>
    </div>
  )

  const sheet = proposal.sheet
  const mid = sheet.midCost

  // Group hours by top-level section for timeline bar
  const bySection = new Map<string, { low: number; high: number }>()
  sheet.tasks.forEach((t) => {
    const sec = t.section.split(' › ')[0]
    const ex = bySection.get(sec) ?? { low: 0, high: 0 }
    bySection.set(sec, { low: ex.low + t.totalHoursLow, high: ex.high + t.totalHoursHigh })
  })
  const maxH = Math.max(...[...bySection.values()].map((v) => v.high))

  const update = (key: 'needs' | 'goals' | 'assumptions', i: number, val: string) => {
    const arr = [...proposal[key]]
    arr[i] = val
    patchProposal({ [key]: arr })
  }
  const remove = (key: 'needs' | 'goals' | 'assumptions', i: number) =>
    patchProposal({ [key]: proposal[key].filter((_, idx) => idx !== i) })
  const add = (key: 'needs' | 'goals', val: string, clear: () => void) => {
    if (val.trim()) { patchProposal({ [key]: [...proposal[key], val.trim()] }); clear() }
  }

  const proceed = () => { markDone('proposal'); setStep('generate') }

  return (
    <div className="max-w-4xl">
      <div className="mb-8">
        <h1 className="page-title">Proposal Editor</h1>
        <p className="page-sub">Review and refine content before generating the presentation deck</p>
      </div>

      {/* Project details */}
      <div className="card mb-4">
        <div className="section-title mb-4">Project details</div>
        <div className="grid grid-cols-2 gap-4">
          <div>
            <label className="label">Project name</label>
            <input className="input" value={proposal.projectName}
              onChange={(e) => patchProposal({ projectName: e.target.value })} />
          </div>
          <div>
            <label className="label">Client name</label>
            <input className="input" value={proposal.clientName}
              onChange={(e) => patchProposal({ clientName: e.target.value })} />
          </div>
        </div>
      </div>

      {/* Estimate snapshot */}
      <div className="card mb-4">
        <div className="section-title mb-3">Estimate snapshot</div>
        <div className="grid grid-cols-3 gap-3">
          {[
            { label: 'Hours range', value: `${sheet.totalHoursLow}–${sheet.totalHoursHigh}` },
            { label: 'Mid-point cost', value: fmtUSD(mid) },
            { label: 'Net after discounts', value: fmtUSD(mid - 15000 - 25000) },
          ].map((s) => (
            <div key={s.label} className="bg-surface border border-border rounded-lg p-3">
              <div className="label">{s.label}</div>
              <div className="font-display text-lg font-semibold text-ink">{s.value}</div>
            </div>
          ))}
        </div>
      </div>

      {/* Needs + Goals */}
      <div className="grid grid-cols-2 gap-4 mb-4">
        {/* Needs */}
        <div className="card">
          <div className="flex items-center justify-between mb-3">
            <div className="section-title">Project needs</div>
            <span className="badge badge-gray">{proposal.needs.length}</span>
          </div>
          <ul className="space-y-2 mb-3">
            {proposal.needs.map((n, i) => (
              <li key={i} className="group flex items-start gap-2">
                <div className="w-1.5 h-1.5 rounded-full bg-brand-green mt-[7px] shrink-0" />
                {editNeed === i
                  ? <input autoFocus className="input text-xs flex-1"
                      value={n} onChange={(e) => update('needs', i, e.target.value)}
                      onBlur={() => setEditNeed(null)}
                      onKeyDown={(e) => e.key === 'Enter' && setEditNeed(null)} />
                  : <span onClick={() => setEditNeed(i)}
                      className="flex-1 text-xs text-ink leading-relaxed cursor-pointer hover:text-brand-green">{n}</span>
                }
                <button onClick={() => remove('needs', i)}
                  className="opacity-0 group-hover:opacity-100 text-ink-3 hover:text-red-400 text-xs transition-all shrink-0 mt-0.5">✕</button>
              </li>
            ))}
          </ul>
          <div className="flex gap-2">
            <input className="input text-xs flex-1" placeholder="Add a need…"
              value={newNeed} onChange={(e) => setNewNeed(e.target.value)}
              onKeyDown={(e) => e.key === 'Enter' && add('needs', newNeed, () => setNewNeed(''))} />
            <button className="btn btn-secondary btn-sm px-3"
              onClick={() => add('needs', newNeed, () => setNewNeed(''))}>+</button>
          </div>
        </div>

        {/* Goals */}
        <div className="card">
          <div className="flex items-center justify-between mb-3">
            <div className="section-title">Project goals</div>
            <span className="badge badge-gray">{proposal.goals.length}</span>
          </div>
          <ul className="space-y-2 mb-3">
            {proposal.goals.map((g, i) => (
              <li key={i} className="group flex items-start gap-2">
                <div className="w-4 h-4 rounded-full border border-brand-navy bg-brand-navy-lt flex items-center justify-center shrink-0 mt-0.5">
                  <span className="font-mono text-[9px] text-brand-navy font-medium">{i + 1}</span>
                </div>
                {editGoal === i
                  ? <input autoFocus className="input text-xs flex-1"
                      value={g} onChange={(e) => update('goals', i, e.target.value)}
                      onBlur={() => setEditGoal(null)}
                      onKeyDown={(e) => e.key === 'Enter' && setEditGoal(null)} />
                  : <span onClick={() => setEditGoal(i)}
                      className="flex-1 text-xs text-ink leading-relaxed cursor-pointer hover:text-brand-navy">{g}</span>
                }
                <button onClick={() => remove('goals', i)}
                  className="opacity-0 group-hover:opacity-100 text-ink-3 hover:text-red-400 text-xs transition-all shrink-0 mt-0.5">✕</button>
              </li>
            ))}
          </ul>
          <div className="flex gap-2">
            <input className="input text-xs flex-1" placeholder="Add a goal…"
              value={newGoal} onChange={(e) => setNewGoal(e.target.value)}
              onKeyDown={(e) => e.key === 'Enter' && add('goals', newGoal, () => setNewGoal(''))} />
            <button className="btn btn-secondary btn-sm px-3"
              onClick={() => add('goals', newGoal, () => setNewGoal(''))}>+</button>
          </div>
        </div>
      </div>

      {/* Effort by section — visual timeline */}
      <div className="card mb-4">
        <div className="section-title mb-4">Effort by work stream</div>
        <div className="space-y-3">
          {[...bySection.entries()].map(([sec, v], i) => {
            const midH = Math.round((v.low + v.high) / 2)
            const pct = Math.round((v.high / maxH) * 100)
            return (
              <div key={sec} className="flex items-center gap-4">
                <div className="font-mono text-xs text-ink-2 w-44 shrink-0 truncate" title={sec}>{sec}</div>
                <div className="flex-1 flex items-center gap-3">
                  <div className="flex-1 h-5 bg-surface-alt rounded overflow-hidden">
                    <div className="h-full rounded flex items-center px-2 transition-all duration-500"
                      style={{ width: `${Math.max(pct, 8)}%`, background: PHASE_COLORS[i % PHASE_COLORS.length], minWidth: 48 }}>
                      <span className="font-mono text-[10px] text-white font-medium">{midH}h</span>
                    </div>
                  </div>
                  <div className="font-mono text-xs text-ink-3 shrink-0 w-28 text-right">
                    {v.low}–{v.high}h
                  </div>
                </div>
              </div>
            )
          })}
        </div>
      </div>

      {/* Assumptions */}
      <div className="card mb-6">
        <div className="flex items-center justify-between mb-3">
          <div className="section-title">Assumptions & exclusions</div>
          <span className="badge badge-gray">{proposal.assumptions.length}</span>
        </div>
        <ul className="space-y-1.5">
          {proposal.assumptions.map((a, i) => (
            <li key={i} className="group flex items-start gap-2">
              <div className="w-1 h-1 rounded-full bg-ink-3 mt-[7px] shrink-0" />
              {editAssumption === i
                ? <input autoFocus className="input text-xs flex-1"
                    value={a} onChange={(e) => update('assumptions', i, e.target.value)}
                    onBlur={() => setEditAssumption(null)}
                    onKeyDown={(e) => e.key === 'Enter' && setEditAssumption(null)} />
                : <span onClick={() => setEditAssumption(i)}
                    className="flex-1 text-xs text-ink-2 leading-relaxed cursor-pointer hover:text-ink">{a}</span>
              }
              <button onClick={() => remove('assumptions', i)}
                className="opacity-0 group-hover:opacity-100 text-ink-3 hover:text-red-400 text-xs transition-all shrink-0 mt-0.5">✕</button>
            </li>
          ))}
        </ul>
      </div>

      <div className="flex gap-3">
        <button className="btn btn-secondary" onClick={() => setStep('preview')}>← Back</button>
        <button className="btn btn-primary" onClick={proceed}>Generate proposal →</button>
      </div>
    </div>
  )
}
