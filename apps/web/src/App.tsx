import { useConnectionState, useZero } from '@rocicorp/zero/react'
import type { BiligRuntimeConfig } from '@bilig/zero-sync'
import type { ZeroClient } from './runtime-session.js'
import { WorkerWorkbookApp } from './WorkerWorkbookApp'
import { WorkbookAppErrorBoundary } from './WorkbookAppErrorBoundary.js'
import type { ZeroConnectionState } from './worker-workbook-app-model.js'

export function App(props: { config: BiligRuntimeConfig; connectionState?: ZeroConnectionState; zero?: ZeroClient }) {
  const resetKey = `${props.config.defaultDocumentId}:${props.config.persistState ? 'persist' : 'memory'}`
  if (props.connectionState) {
    return (
      <WorkbookAppErrorBoundary resetKey={resetKey}>
        <WorkerWorkbookApp config={props.config} connectionState={props.connectionState} {...(props.zero ? { zero: props.zero } : {})} />
      </WorkbookAppErrorBoundary>
    )
  }
  return (
    <WorkbookAppErrorBoundary resetKey={resetKey}>
      <ConnectedApp config={props.config} />
    </WorkbookAppErrorBoundary>
  )
}

function ConnectedApp({ config }: { config: BiligRuntimeConfig }) {
  const zero = useZero()
  const connectionState = useConnectionState()
  return <WorkerWorkbookApp config={config} connectionState={connectionState} zero={zero} />
}
