import React, { useState } from 'react'
import { useStore } from '../store/appStore'
import { fmtUSD } from '../utils/download'

export function PreviewPage() {
  const { sheet, setStep } = useStore()
  const [openSections, setOpenSections] = useState<Set<string>>(new Set())

  if (!sheet) return (
    <div className="flex flex-col items-center justify-center h-96 text-center">
      <p className="text-ink-2 font-medium">No data loaded</p>
      <button className="btn btn-primary mt-4" onClick={() => setStep('upload')}>← Upload</button>
    </div>
  )

  // Group tasks by top-level section
  const bySection = new Map<string, typeof sheet.tasks>()
  sheet.tasks.forEach((t) => {
    const sec = t.section.split(' › ')[0]
    if (!bySection.has(sec)) bySection.set(sec, [])
    bySection.get(sec)!.push(t)
  })

  const toggle = (sec: string) => setOpenSections((prev) => {
    const next = new Set(prev)
    next.has(sec) ? next.delete(sec) : next.add(sec)
    return next
  })

  const midH = Math.round((sheet.totalHoursLow + sheet.totalHoursHigh) / 2)
  const midC = sheet.midCost

  return (
    <div className="max-w-5xl">
      <div className="mb-6">
        <h1 className="page-title">Parsed Data Preview</h1>
        <p className="page-sub">Sheet: <span className="font-mono text-brand-green">{sheet.sheetName}</span> · {sheet.tasks.length} tasks across {sheet.sections.length} work streams</p>
      </div>

      {/* Stats */}
      <div className="grid grid-cols-4 gap-3 mb-6">
        {[
          { label: 'Tasks', value: sheet.tasks.length, sub: 'line items parsed' },
          { label: 'Est. hours', value: `${sheet.totalHoursLow}–${sheet.totalHoursHigh}`, sub: `~${midH}h mid-point` },
          { label: 'Mid estimate', value: fmtUSD(midC), sub: `${fmtUSD(sheet.totalCostLow)}–${fmtUSD(sheet.totalCostHigh)}` },
          { label: 'Work streams', value: sheet.sections.length, sub: bySection.size + ' sections' },
        ].map((s) => (
          <div key={s.label} className="stat-card">
            <div className="stat-label">{s.label}</div>
            <div className="stat-value text-xl">{s.value}</div>
            <div className="stat-sub">{s.sub}</div>
          </div>
        ))}
      </div>

      {/* Role summary */}
      <div className="card mb-4">
        <div className="section-title mb-4">Role allocation summary</div>
        <div className="overflow-x-auto">
          <table className="w-full text-sm">
            <thead><tr>
              {['Role','Hours (low)','Hours (high)','Mid hours','Cost (low)','Cost (high)','Mid cost'].map(h=><th key={h} className="th">{h}</th>)}
            </tr></thead>
            <tbody>
              {sheet.roles.map((r) => (
                <tr key={r.role} className="hover:bg-surface">
                  <td className="td font-medium">{r.role}</td>
                  <td className="td font-mono text-ink-2">{r.hoursLow}</td>
                  <td className="td font-mono text-ink-2">{r.hoursHigh}</td>
                  <td className="td font-mono font-medium">{Math.round((r.hoursLow+r.hoursHigh)/2)}</td>
                  <td className="td font-mono text-ink-2">{fmtUSD(r.costLow)}</td>
                  <td className="td font-mono text-ink-2">{fmtUSD(r.costHigh)}</td>
                  <td className="td font-mono font-medium text-brand-green">{fmtUSD(Math.round((r.costLow+r.costHigh)/2))}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>

      {/* Tasks by section — collapsible */}
      <div className="mb-4">
        <div className="section-title mb-3">Tasks by work stream</div>
        <div className="space-y-2">
          {[...bySection.entries()].map(([sec, tasks]) => {
            const open = openSections.has(sec)
            const secH = tasks.reduce((s,t)=>s+Math.round((t.totalHoursLow+t.totalHoursHigh)/2),0)
            const secC = tasks.reduce((s,t)=>s+Math.round((t.totalCostLow+t.totalCostHigh)/2),0)
            return (
              <div key={sec} className="border border-border rounded-xl overflow-hidden">
                <button
                  onClick={() => toggle(sec)}
                  className="w-full flex items-center justify-between px-4 py-3 bg-surface-card hover:bg-surface text-sm font-medium text-ink transition-colors"
                >
                  <span>{sec}</span>
                  <div className="flex items-center gap-3">
                    <span className="badge badge-gray">{tasks.length} tasks</span>
                    <span className="badge badge-green">~{secH}h</span>
                    <span className="badge badge-amber">{fmtUSD(secC)}</span>
                    <svg width="14" height="14" viewBox="0 0 14 14" fill="none" className={`transition-transform ${open?'rotate-180':''}`}>
                      <path d="M3.5 5.5l3.5 3.5 3.5-3.5" stroke="currentColor" strokeWidth="1.3" strokeLinecap="round"/>
                    </svg>
                  </div>
                </button>
                {open && (
                  <div className="border-t border-border bg-surface-alt overflow-x-auto">
                    <table className="w-full text-xs">
                      <thead><tr>
                        {['Task','Hours (low–high)','Cost (low–high)','Comments'].map(h=><th key={h} className="th">{h}</th>)}
                      </tr></thead>
                      <tbody>
                        {tasks.map((t, i) => (
                          <tr key={i} className="hover:bg-surface">
                            <td className="td font-medium max-w-[320px]">{t.task}</td>
                            <td className="td font-mono text-ink-2">{t.totalHoursLow}–{t.totalHoursHigh}</td>
                            <td className="td font-mono text-ink-2">{fmtUSD(t.totalCostLow)}–{fmtUSD(t.totalCostHigh)}</td>
                            <td className="td text-ink-3 italic max-w-[200px]">{t.comments || '—'}</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                )}
              </div>
            )
          })}
        </div>
      </div>

      <div className="flex gap-3">
        <button className="btn btn-secondary" onClick={() => setStep('upload')}>← Re-upload</button>
        <button className="btn btn-primary" onClick={() => setStep('proposal')}>Review proposal content →</button>
      </div>
    </div>
  )
}
