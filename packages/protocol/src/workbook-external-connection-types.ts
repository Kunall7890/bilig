export type WorkbookExternalConnectionSourceKind = 'database' | 'web-query' | 'text' | 'model' | 'external-workbook' | 'unknown'

export interface WorkbookExternalConnectionSnapshot {
  id: number
  name?: string
  sourceKind: WorkbookExternalConnectionSourceKind
  type?: string
  description?: string
  connection?: string
  command?: string
  commandType?: string
  refreshOnLoad?: boolean
  saveData?: boolean
  clause: '18.13'
}

export interface WorkbookExternalWorkbookLinkSnapshot {
  kind: 'external-workbook'
  bookIndex: number
  packagePath?: string
  target?: string
  targetMode?: string
  workbookName?: string
  sheetNames?: string[]
  definedNames?: string[]
  clause: '18.14'
}

export interface WorkbookDdeLinkSnapshot {
  kind: 'dde'
  service?: string
  topic?: string
  itemNames?: string[]
  refreshExecution: 'disabled'
  packagePath?: string
  clause: '18.14'
}

export interface WorkbookOleLinkSnapshot {
  kind: 'ole'
  progId?: string
  relationshipId?: string
  target?: string
  targetMode?: string
  itemNames?: string[]
  refreshExecution: 'disabled'
  packagePath?: string
  clause: '18.14'
}

export type WorkbookExternalLinkSnapshot = WorkbookExternalWorkbookLinkSnapshot | WorkbookDdeLinkSnapshot | WorkbookOleLinkSnapshot

export interface WorkbookExternalConnectionsSnapshot {
  refreshExecution: 'disabled'
  connections?: WorkbookExternalConnectionSnapshot[]
  externalLinks?: WorkbookExternalLinkSnapshot[]
}
