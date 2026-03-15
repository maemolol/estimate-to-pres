import React, { useCallback, useState } from 'react'
import { useDropzone } from 'react-dropzone'
import { useStore } from '../store/appStore'
import { parseEstimateFile, getLatticeDemo } from '../services/excelParser'
import { buildProposal } from '../services/proposalBuilder'

export function UploadPage() {
  const { setStep, markDone, setSheet, setProposal, setError, error } = useStore()
  const [parsing, setParsing] = useState(false)
  const [warnings, setWarnings] = useState<string[]>([])
  const [loaded, setLoaded] = useState('')

  const handleFile = useCallback(async (file: File) => {
    setParsing(true); setWarnings([]); setError(null)
    try {
      const { sheets, primary, warnings: w } = await parseEstimateFile(file)
      setSheet(primary)
      setProposal(buildProposal(primary))
      setWarnings(w)
      setLoaded(file.name)
      markDone('upload')
      setTimeout(() => setStep('preview'), 500)
    } catch (e) { setError((e as Error).message) }
    finally { setParsing(false) }
  }, [setSheet, setProposal, setStep, markDone, setError])

  const loadDemo = () => {
    const { primary } = getLatticeDemo()
    setSheet(primary)
    setProposal(buildProposal(primary))
    setLoaded('Lattice_XP_to_XMC_Estimate.xlsx (demo)')
    markDone('upload')
    setTimeout(() => setStep('preview'), 300)
  }

  const { getRootProps, getInputProps, isDragActive } = useDropzone({
    onDrop: (f) => f[0] && handleFile(f[0]),
    accept: {
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'],
      'application/vnd.ms-excel': ['.xls'],
      'text/csv': ['.csv'],
    },
    maxFiles: 1,
    disabled: parsing,
  })

  return (
    <div className="max-w-2xl">
      <div className="mb-8">
        <h1 className="page-title">Upload Estimate</h1>
        <p className="page-sub">Import a Zont-style Excel estimate spreadsheet to begin the proposal pipeline</p>
      </div>

      {error && (
        <div className="mb-4 flex gap-3 bg-red-50 border border-red-200 rounded-xl p-4 text-sm text-red-700">
          <span className="shrink-0 mt-0.5">⚠</span>{error}
        </div>
      )}

      {loaded && !error && (
        <div className="mb-4 flex items-center gap-3 bg-brand-green-lt border border-green-200 rounded-xl px-4 py-3 text-sm text-brand-green">
          <span>✓</span><strong>{loaded}</strong><span className="text-brand-greenMd">— parsed, redirecting…</span>
        </div>
      )}

      <div className="card mb-4">
        <div className="section-title mb-4">Spreadsheet file</div>
        <div {...getRootProps()} className={`border-[1.5px] border-dashed rounded-xl p-12 text-center cursor-pointer transition-all ${isDragActive ? 'drop-active' : 'border-border hover:border-brand-green hover:bg-brand-green-lt'}`}>
          <input {...getInputProps()} />
          {parsing ? (
            <div className="flex flex-col items-center gap-3">
              <div className="spinner" /><span className="text-sm text-ink-2">Parsing…</span>
            </div>
          ) : (
            <>
              <div className="text-3xl mb-3 opacity-20">⬆</div>
              <p className="text-sm text-ink-2">{isDragActive ? 'Drop it here' : 'Drop your Excel estimate, or click to browse'}</p>
              <p className="font-mono text-xs text-ink-3 mt-1">Accepts .xlsx · .xls · .csv</p>
            </>
          )}
        </div>
      </div>

      {warnings.length > 0 && (
        <div className="mb-4 bg-brand-amber-lt border border-amber-200 rounded-xl p-4">
          <div className="section-title text-brand-amber mb-2">Parse warnings</div>
          {warnings.map((w, i) => <p key={i} className="font-mono text-xs text-brand-amber">{w}</p>)}
        </div>
      )}

      <div className="card mb-4">
        <div className="section-title mb-3">Expected format — Zont estimate structure</div>
        <p className="text-sm text-ink-2 mb-3 leading-relaxed">
          The parser understands multi-sheet estimates with role-based hour columns (FED Low/High, BED Low/High, Arch, QA/BA, PM, Gov), task groupings, and bottom summary rows with rates and costs.
        </p>
        <div className="overflow-x-auto">
          <table className="w-full text-xs">
            <thead><tr>
              {['Task','FED Low','FED High','BED Low','BED High','Arch Low','Arch High','QA/BA Low','QA/BA High','PM Low','PM High','Comments'].map(h=>(
                <th key={h} className="th whitespace-nowrap">{h}</th>
              ))}
            </tr></thead>
            <tbody>
              <tr className="border-b border-surface bg-surface">
                <td className="td italic text-ink-2 font-medium" colSpan={12}>Onboarding</td>
              </tr>
              {[
                ['Project setup','0','0','0','0','1','1','4','4','4','6',''],
                ['Discovery workshops','0','0','0','0','1','2','8','12','2','2',''],
                ['Functional Spec','0','0','0','0','2','3','32','40','5.1','6.5',''],
              ].map((r,i)=>(
                <tr key={i} className="hover:bg-surface">
                  {r.map((c,j)=><td key={j} className={`td ${j===0?'font-medium':'font-mono text-ink-2'}`}>{c}</td>)}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>

      <div className="flex items-center gap-3">
        <button className="btn btn-secondary" onClick={loadDemo}>
          Load Lattice demo data
        </button>
        <span className="text-xs text-ink-3">Uses the real Lattice XP→XMC estimate as a demo</span>
      </div>
    </div>
  )
}
