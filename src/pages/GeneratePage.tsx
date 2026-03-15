import React, { useCallback, useState } from 'react'
import { useDropzone } from 'react-dropzone'
import { useStore } from '../store/appStore'
import { generatePptx } from '../services/pptxGenerator'
import { dlBlob, slug } from '../utils/download'

type Stage = 'idle' | 'building' | 'ready' | 'awaiting' | 'tweaked'

export function GeneratePage() {
  const {
    proposal, setStep, markDone,
    setGeneratedPptx, setTweakedPptx,
    generatedPptx, tweakedPptx,
    isGeneratingPptx, setGeneratingPptx,
  } = useStore()

  const [stage, setStage] = useState<Stage>(
    tweakedPptx ? 'tweaked' : generatedPptx ? 'awaiting' : 'idle'
  )
  const [progress, setProgress] = useState(0)
  const [tweakName, setTweakName] = useState<string | null>(null)
  const [err, setErr] = useState<string | null>(null)

  if (!proposal) return (
    <div className="flex flex-col items-center justify-center h-96 text-center">
      <p className="text-ink-2 font-medium">No proposal data — complete previous steps first</p>
      <button className="btn btn-primary mt-4" onClick={() => setStep('upload')}>← Start over</button>
    </div>
  )

  const handleGenerate = async () => {
    setErr(null)
    setStage('building')
    setGeneratingPptx(true)
    setProgress(0)
    try {
      const ticker = setInterval(() => setProgress((p) => Math.min(p + 9, 88)), 180)
      const blob = await generatePptx(proposal)
      clearInterval(ticker)
      setProgress(100)
      setGeneratedPptx(blob)
      markDone('generate')
      setTimeout(() => { setStage('awaiting'); setGeneratingPptx(false) }, 400)
    } catch (e) {
      setErr((e as Error).message)
      setStage('idle')
      setGeneratingPptx(false)
    }
  }

  const handleDownload = () => {
    if (!generatedPptx) return
    dlBlob(generatedPptx, `${slug(proposal.projectName)}_proposal.pptx`)
  }

  const onDrop = useCallback((accepted: File[]) => {
    const f = accepted[0]
    if (!f) return
    setTweakedPptx(new Blob([f], { type: f.type }))
    setTweakName(f.name)
    setStage('tweaked')
    markDone('generate')
  }, [setTweakedPptx, markDone])

  const { getRootProps, getInputProps, isDragActive } = useDropzone({
    onDrop,
    accept: { 'application/vnd.openxmlformats-officedocument.presentationml.presentation': ['.pptx'] },
    maxFiles: 1,
    disabled: stage !== 'awaiting',
  })

  const stepDone = (n: number) => {
    if (n === 1) return stage !== 'idle' && stage !== 'building'
    if (n === 2) return stage === 'tweaked'
    if (n === 3) return stage === 'tweaked'
    return false
  }

  return (
    <div className="max-w-2xl">
      <div className="mb-8">
        <h1 className="page-title">Generate & Download</h1>
        <p className="page-sub">Build the deck, tweak it in PowerPoint, then re-upload the final version to unlock the contract</p>
      </div>

      {/* Progress steps */}
      <div className="flex items-center mb-8 gap-0">
        {['Generate PPTX', 'Download & edit', 'Re-upload final'].map((label, i) => (
          <React.Fragment key={i}>
            <div className="flex flex-col items-center gap-1.5">
              <div className={`w-7 h-7 rounded-full flex items-center justify-center text-xs font-mono font-medium border-2 transition-all duration-300 ${stepDone(i+1) ? 'bg-brand-green border-brand-green text-white' : 'bg-surface-card border-border text-ink-3'}`}>
                {stepDone(i+1) ? '✓' : i + 1}
              </div>
              <span className={`text-[10px] font-mono whitespace-nowrap ${stepDone(i+1) ? 'text-brand-green' : 'text-ink-3'}`}>{label}</span>
            </div>
            {i < 2 && (
              <div className={`flex-1 h-0.5 mb-5 mx-2 transition-all duration-300 ${stepDone(i+1) ? 'bg-brand-green' : 'bg-border'}`} />
            )}
          </React.Fragment>
        ))}
      </div>

      {err && (
        <div className="mb-4 flex gap-3 bg-red-50 border border-red-200 rounded-xl p-4 text-sm text-red-700">
          <span className="shrink-0">⚠</span>{err}
        </div>
      )}

      {/* Step 1 — Generate */}
      <div className={`card mb-4 transition-all ${stage === 'idle' || stage === 'building' ? 'ring-2 ring-brand-green ring-offset-1' : ''}`}>
        <div className="flex items-start justify-between">
          <div>
            <div className="font-medium text-sm text-ink">Step 1 — Generate PPTX</div>
            <div className="text-xs text-ink-2 mt-0.5">
              Builds a 12-slide branded deck for <span className="font-medium">{proposal.projectName}</span>
            </div>
          </div>
          {stepDone(1) && <span className="badge badge-green">Done</span>}
        </div>

        {stage === 'building' && (
          <div className="mt-3">
            <div className="h-1.5 bg-surface-alt rounded-full overflow-hidden">
              <div className="h-full bg-brand-green rounded-full transition-all duration-200" style={{ width: `${progress}%` }} />
            </div>
            <div className="text-[11px] text-ink-3 mt-1 font-mono">Building slides… {progress}%</div>
          </div>
        )}

        <div className="mt-4">
          <button className="btn btn-primary"
            disabled={stage !== 'idle'}
            onClick={handleGenerate}>
            {stage === 'building'
              ? <><div className="spinner" />Generating…</>
              : <>⚡ Generate proposal deck</>}
          </button>
        </div>
      </div>

      {/* Step 2 — Download */}
      <div className={`card mb-4 transition-all ${stage === 'awaiting' ? 'ring-2 ring-brand-green ring-offset-1' : ''}`}>
        <div className="flex items-start justify-between">
          <div>
            <div className="font-medium text-sm text-ink">Step 2 — Download & tweak in PowerPoint</div>
            <div className="text-xs text-ink-2 mt-0.5">Edit slides, adjust styling, update any numbers, then save as .pptx</div>
          </div>
          {stepDone(2) && <span className="badge badge-green">Done</span>}
        </div>
        <div className="mt-4 flex items-center gap-3">
          <button className="btn btn-primary" disabled={!generatedPptx} onClick={handleDownload}>
            ↓ Download .pptx
          </button>
          {generatedPptx && (
            <span className="text-xs text-ink-3">
              {proposal.projectName.slice(0, 30)} · 12 slides · ready
            </span>
          )}
        </div>
      </div>

      {/* Step 3 — Re-upload tweaked */}
      <div className={`card mb-6 transition-all ${stage === 'awaiting' ? 'ring-2 ring-brand-amber ring-offset-1' : ''}`}>
        <div className="flex items-start justify-between">
          <div>
            <div className="font-medium text-sm text-ink">Step 3 — Upload your tweaked .pptx</div>
            <div className="text-xs text-ink-2 mt-0.5">
              Once you're happy with the deck, upload the final version here — this is what the contract will be based on
            </div>
          </div>
          {stage === 'tweaked' && <span className="badge badge-green">Uploaded</span>}
        </div>

        <div {...getRootProps()} className={`mt-4 border-[1.5px] border-dashed rounded-xl p-8 text-center transition-all duration-200
          ${stage !== 'awaiting'
            ? 'border-border bg-surface-alt opacity-50 cursor-not-allowed'
            : isDragActive
              ? 'drop-active cursor-pointer'
              : 'border-amber-200 hover:border-brand-amber hover:bg-brand-amber-lt cursor-pointer'
          }`}>
          <input {...getInputProps()} />
          {stage === 'tweaked' ? (
            <div className="flex flex-col items-center gap-2">
              <div className="w-8 h-8 rounded-full bg-brand-green-lt flex items-center justify-center text-brand-green">✓</div>
              <p className="text-sm text-brand-green font-medium">{tweakName ?? 'Tweaked file loaded'}</p>
              <p className="text-xs text-ink-3">Click to replace</p>
            </div>
          ) : (
            <div className="flex flex-col items-center gap-2">
              <div className="text-2xl opacity-20">⬆</div>
              <p className="text-sm text-ink-2">
                {stage === 'awaiting' ? 'Drop your tweaked .pptx here, or click to browse' : 'Complete steps 1 & 2 first'}
              </p>
              <p className="font-mono text-xs text-ink-3">Only .pptx accepted</p>
            </div>
          )}
        </div>
      </div>

      {/* CTA */}
      {stage === 'tweaked' && (
        <div className="flex items-center gap-3">
          <button className="btn btn-primary" onClick={() => setStep('contract')}>
            Generate SLW contract →
          </button>
          <span className="text-xs text-ink-3">Contract will be generated from your tweaked deck</span>
        </div>
      )}
    </div>
  )
}
