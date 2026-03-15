import React, { useState } from 'react'
import { useStore } from '../store/appStore'
import {
  generateContractSections,
  buildDocx,
  buildMarkdown,
  SECTION_LABELS,
  type ContractSections,
} from '../services/contractGenerator'
import { dlBlob, slug, fmtUSD } from '../utils/download'

type CState = 'idle' | 'generating' | 'ready'

export function ContractPage() {
  const { proposal, tweakedPptx, setStep, markDone, setContract } = useStore()
  const [apiKey, setApiKey] = useState((import.meta as any).env?.VITE_OPENAI_API_KEY ?? '')
  const [cState, setCState] = useState<CState>('idle')
  const [sections, setSections] = useState<ContractSections | null>(null)
  const [docxBlob, setDocxBlob] = useState<Blob | null>(null)
  const [md, setMd] = useState<string | null>(null)
  const [open, setOpen] = useState<string | null>('projectOverview')
  const [statusMsg, setStatusMsg] = useState('')
  const [err, setErr] = useState<string | null>(null)

  if (!tweakedPptx) return (
    <div className="flex flex-col items-center justify-center h-96 text-center">
      <div className="text-4xl opacity-10 mb-4">◻</div>
      <p className="text-ink-2 font-medium">No tweaked PPTX uploaded yet</p>
      <p className="text-xs text-ink-3 mt-1 mb-6">Go back to step 4, download the deck, tweak it, then re-upload</p>
      <button className="btn btn-primary" onClick={() => setStep('generate')}>← Back to Generate</button>
    </div>
  )

  if (!proposal) return (
    <div className="flex flex-col items-center justify-center h-96 text-center">
      <p className="text-ink-2 font-medium">No proposal data</p>
      <button className="btn btn-primary mt-4" onClick={() => setStep('upload')}>← Start over</button>
    </div>
  )

  const net = proposal.sheet.midCost - 15000 - 25000

  const handleGenerate = async () => {
    if (!apiKey.trim()) { setErr('Enter your OpenAI API key to continue'); return }
    setErr(null)
    setCState('generating')

    try {
      setStatusMsg('Sending proposal data to GPT-4 Turbo…')
      const secs = await generateContractSections(proposal, apiKey.trim())
      setSections(secs)

      setStatusMsg('Building DOCX document…')
      const docx = await buildDocx(proposal, secs)
      setDocxBlob(docx)

      setStatusMsg('Building Markdown export…')
      const markdown = buildMarkdown(proposal, secs)
      setMd(markdown)

      setContract(docx, markdown)
      markDone('contract')
      setCState('ready')
      setStatusMsg('')
    } catch (e) {
      setErr((e as Error).message)
      setCState('idle')
    }
  }

  return (
    <div className="max-w-3xl">
      <div className="mb-8">
        <h1 className="page-title">SLW Contract Draft</h1>
        <p className="page-sub">GPT-4 generates a Statement of Work from your final proposal deck</p>
      </div>

      {/* Source badge */}
      <div className="mb-5 flex items-center gap-3 bg-brand-green-lt border border-green-200 rounded-xl px-4 py-3 text-sm text-brand-green">
        <span className="shrink-0">✓</span>
        <span>
          Using tweaked PPTX · <strong>{proposal.projectName}</strong> · {proposal.clientName} · net {fmtUSD(net)}
        </span>
      </div>

      {/* API key input */}
      {!(import.meta as any).env?.VITE_OPENAI_API_KEY && cState === 'idle' && (
        <div className="card mb-4">
          <div className="section-title mb-3">OpenAI API key</div>
          <input type="password" className="input" placeholder="sk-…"
            value={apiKey} onChange={(e) => setApiKey(e.target.value)} />
          <p className="text-xs text-ink-3 mt-2">
            Used directly from the browser — never stored. Set{' '}
            <code className="font-mono bg-surface px-1 rounded">VITE_OPENAI_API_KEY</code> in{' '}
            <code className="font-mono bg-surface px-1 rounded">.env</code> to skip this step.
          </p>
        </div>
      )}

      {/* Prompt preview */}
      {cState === 'idle' && (
        <div className="card mb-4">
          <div className="section-title mb-3">GPT-4 prompt preview</div>
          <div className="bg-surface rounded-lg p-4 font-mono text-xs text-ink-2 leading-relaxed border border-border whitespace-pre-wrap">{
`SYSTEM  You are a senior technology contracts lawyer.

USER    Generate a Statement of Work for "${proposal.projectName}"
        Client: ${proposal.clientName}
        Net investment: ${fmtUSD(net)}
        Hours (mid): ~${Math.round((proposal.sheet.totalHoursLow + proposal.sheet.totalHoursHigh) / 2)}h
        Work streams: ${proposal.sheet.sections.slice(0, 5).join(', ')}${proposal.sheet.sections.length > 5 ? '…' : ''}
        Return valid JSON with 9 contract sections.`}
          </div>
        </div>
      )}

      {err && (
        <div className="mb-4 flex gap-3 bg-red-50 border border-red-200 rounded-xl p-4 text-sm text-red-700">
          <span className="shrink-0 mt-0.5">⚠</span>{err}
        </div>
      )}

      {/* Generate button */}
      {cState !== 'ready' && (
        <div className="card mb-6">
          {cState === 'generating' && (
            <div className="flex items-center gap-3 mb-4 p-3 bg-brand-green-lt rounded-lg">
              <div className="spinner shrink-0" />
              <span className="text-sm text-brand-green">{statusMsg}</span>
            </div>
          )}
          <button className="btn btn-primary"
            disabled={cState === 'generating' || !apiKey.trim()}
            onClick={handleGenerate}>
            {cState === 'generating'
              ? <><div className="spinner" />Generating contract…</>
              : <>Generate SLW contract with GPT-4</>}
          </button>
        </div>
      )}

      {/* Contract sections */}
      {cState === 'ready' && sections && (
        <>
          <div className="mb-4 flex items-center gap-3 bg-brand-green-lt border border-green-200 rounded-xl px-4 py-3 text-sm text-brand-green">
            <span>✓</span> Contract generated — {SECTION_LABELS.length} sections · review below before downloading
          </div>

          <div className="space-y-2 mb-6">
            {SECTION_LABELS.map(({ key, title }) => (
              <div key={key} className="border border-border rounded-xl overflow-hidden">
                <button
                  onClick={() => setOpen(open === key ? null : key)}
                  className="w-full flex items-center justify-between px-4 py-3 bg-surface-card hover:bg-surface text-sm font-medium text-ink transition-colors">
                  <span>{title}</span>
                  <svg width="14" height="14" viewBox="0 0 14 14" fill="none"
                    className={`transition-transform shrink-0 ${open === key ? 'rotate-180' : ''}`}>
                    <path d="M3.5 5.5l3.5 3.5 3.5-3.5" stroke="currentColor" strokeWidth="1.3" strokeLinecap="round"/>
                  </svg>
                </button>
                {open === key && (
                  <div className="px-4 pb-4 pt-2 bg-surface border-t border-border text-sm text-ink-2 leading-relaxed whitespace-pre-wrap">
                    {sections[key]}
                  </div>
                )}
              </div>
            ))}
          </div>

          {/* Signature block preview */}
          <div className="card mb-6">
            <div className="section-title mb-4">Signature block</div>
            <div className="grid grid-cols-2 gap-6">
              {[['Client', proposal.clientName], ['Service Provider', 'Zont Digital']].map(([role, name]) => (
                <div key={role} className="bg-surface rounded-xl p-4 border border-border">
                  <div className="label">{role}</div>
                  <div className="text-sm text-ink-2 mb-4">{name}</div>
                  {['Authorised Representative', 'Signature', 'Date'].map((f) => (
                    <div key={f} className="mb-3">
                      <div className="text-xs text-ink-3 mb-1">{f}</div>
                      <div className="h-7 border-b border-border-strong" />
                    </div>
                  ))}
                </div>
              ))}
            </div>
          </div>

          {/* Export */}
          <div className="card">
            <div className="section-title mb-4">Export contract</div>
            <div className="grid grid-cols-2 gap-3">
              <div className="bg-surface rounded-xl p-5 text-center border border-border">
                <div className="text-3xl mb-2 opacity-50">📝</div>
                <div className="font-medium text-sm text-ink mb-1">Word document</div>
                <div className="font-mono text-xs text-ink-3 mb-3">.docx · fully editable</div>
                <button className="btn btn-primary w-full justify-center"
                  onClick={() => docxBlob && dlBlob(docxBlob, `${slug(proposal.projectName)}_SLW_contract.docx`)}>
                  ↓ Download .docx
                </button>
              </div>
              <div className="bg-surface rounded-xl p-5 text-center border border-border">
                <div className="text-3xl mb-2 opacity-50">📋</div>
                <div className="font-medium text-sm text-ink mb-1">Markdown</div>
                <div className="font-mono text-xs text-ink-3 mb-3">.md · plain text</div>
                <button className="btn btn-secondary w-full justify-center"
                  onClick={() => md && dlBlob(new Blob([md], { type: 'text/markdown' }), `${slug(proposal.projectName)}_SLW_contract.md`)}>
                  ↓ Download .md
                </button>
              </div>
            </div>
            <p className="text-xs text-ink-3 mt-4 text-center">
              For PDF: open the .docx in Word → File → Export → PDF
            </p>
          </div>
        </>
      )}
    </div>
  )
}
