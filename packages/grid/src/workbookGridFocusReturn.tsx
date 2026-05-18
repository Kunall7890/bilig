import { createContext, useContext, type ReactNode } from 'react'

export type WorkbookGridFocusRequester = () => void

const WorkbookGridFocusReturnContext = createContext<WorkbookGridFocusRequester | null>(null)

export function WorkbookGridFocusReturnProvider(props: {
  readonly children: ReactNode
  readonly requestGridFocus: WorkbookGridFocusRequester
}) {
  return <WorkbookGridFocusReturnContext.Provider value={props.requestGridFocus}>{props.children}</WorkbookGridFocusReturnContext.Provider>
}

export function useWorkbookGridFocusReturn(): WorkbookGridFocusRequester | null {
  return useContext(WorkbookGridFocusReturnContext)
}
