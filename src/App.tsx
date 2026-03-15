import React from 'react'
import { useStore } from './store/appStore'
import { Sidebar } from './components/Sidebar'
import { UploadPage } from './pages/UploadPage'
import { PreviewPage } from './pages/PreviewPage'
import { ProposalPage } from './pages/ProposalPage'
import { GeneratePage } from './pages/GeneratePage'
import { ContractPage } from './pages/ContractPage'

const PAGES = {
  upload:   <UploadPage />,
  preview:  <PreviewPage />,
  proposal: <ProposalPage />,
  generate: <GeneratePage />,
  contract: <ContractPage />,
}

export default function App() {
  const { step, done, setStep, setError } = useStore()
  return (
    <div className="flex min-h-screen bg-surface">
      <Sidebar current={step} done={done} onNav={(s) => { setError(null); setStep(s) }} />
      <main className="flex-1 px-10 py-10 overflow-y-auto min-h-screen">
        {PAGES[step]}
      </main>
    </div>
  )
}
