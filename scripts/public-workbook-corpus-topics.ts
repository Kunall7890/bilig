import { asRecordOrNull, readArray, readString } from './public-workbook-corpus-json.ts'

export type PublicWorkbookRequiredTopic = 'financial-workpapers'

export const defaultFinancialWorkbookQueries = [
  'accounting',
  'accounts payable',
  'accounts receivable',
  'annual report',
  'appropriation',
  'audit',
  'auditor',
  'budget',
  'cash',
  'cash flow',
  'expenditure',
  'expense',
  'financial',
  'financial report',
  'financial statement',
  'financial tables',
  'finance',
  'fiscal',
  'invoice',
  'ledger',
  'operating statement',
  'payment',
  'payroll',
  'procurement',
  'profit and loss',
  'public accounts',
  'revenue',
  'taxation',
  'trial balance',
  'workpaper',
  'work paper',
  'working paper',
] as const

interface FinancialTopicSourceCandidate {
  readonly dataset: Record<string, unknown>
  readonly resource: Record<string, unknown>
  readonly sourceUrl: string
  readonly downloadUrl: string
  readonly fileName: string
}

const financialTopicPatterns: readonly { readonly label: string; readonly pattern: RegExp }[] = [
  { label: 'accounting', pattern: /\baccounting\b/iu },
  { label: 'account-balance', pattern: /\ballocation\s+account\s+balance\b|\baccounts?\s+balances?\b/iu },
  { label: 'accounts-payable', pattern: /\baccounts?\s+payable\b/iu },
  { label: 'accounts-receivable', pattern: /\baccounts?\s+receivable\b/iu },
  { label: 'annual-report', pattern: /\bannual\s+reports?\b/iu },
  { label: 'appropriation', pattern: /\bappropriations?\b/iu },
  { label: 'audit', pattern: /\baudit(?:ed|ing|or|ors)?\b/iu },
  { label: 'balance-sheet', pattern: /\bbalance\s+sheets?\b/iu },
  { label: 'budget', pattern: /\bbudgets?\b/iu },
  { label: 'cash-flow', pattern: /\bcash\s*flows?\b/iu },
  { label: 'expenditure', pattern: /\bexpenditures?\b/iu },
  { label: 'expense', pattern: /\bexpenses?\b/iu },
  { label: 'finance', pattern: /\bfinanc(?:e|ial|ing)\b/iu },
  { label: 'financial-statement', pattern: /\bfinancial\s+statements?\b/iu },
  { label: 'general-ledger', pattern: /\bgeneral\s+ledger\b/iu },
  { label: 'invoice', pattern: /\binvoices?\b/iu },
  { label: 'ledger', pattern: /\bledgers?\b/iu },
  { label: 'payroll', pattern: /\bpayroll\b/iu },
  { label: 'public-accounts', pattern: /\bpublic\s+accounts?\b/iu },
  { label: 'revenue', pattern: /\brevenues?\b/iu },
  { label: 'tax', pattern: /\btax(?:es|ation)?\b/iu },
  { label: 'trial-balance', pattern: /\btrial\s+balance\b/iu },
  { label: 'workpaper', pattern: /\bwork\s*papers?\b|\bworkpapers?\b|\bworking\s+papers?\b/iu },
]

export function financialWorkbookTopicEvidence(candidate: FinancialTopicSourceCandidate): string[] {
  const fields = collectFinancialTopicFields(candidate)
  const evidence: string[] = []
  const seen = new Set<string>()
  for (const field of fields) {
    for (const topic of financialTopicPatterns) {
      if (!topic.pattern.test(field.value)) {
        continue
      }
      const key = `${topic.label}:${field.name}`
      if (seen.has(key)) {
        continue
      }
      seen.add(key)
      evidence.push(key)
    }
  }
  return evidence
}

function collectFinancialTopicFields(candidate: FinancialTopicSourceCandidate): { readonly name: string; readonly value: string }[] {
  return [
    ...collectRecordFields('dataset', candidate.dataset, ['title', 'name', 'notes', 'url']),
    ...collectNamedCollectionFields('dataset.tag', readArray(candidate.dataset, 'tags')),
    ...collectNamedCollectionFields('dataset.group', readArray(candidate.dataset, 'groups')),
    ...collectRecordFields('resource', candidate.resource, ['name', 'description', 'url', 'format', 'mimetype', 'filename']),
    { name: 'sourceUrl', value: candidate.sourceUrl },
    { name: 'downloadUrl', value: candidate.downloadUrl },
    { name: 'fileName', value: candidate.fileName },
  ]
}

function collectNamedCollectionFields(prefix: string, values: readonly unknown[]): { readonly name: string; readonly value: string }[] {
  return values.flatMap((value, index) => {
    const record = asRecordOrNull(value)
    return record ? collectRecordFields(`${prefix}[${String(index)}]`, record, ['display_name', 'title', 'name', 'description']) : []
  })
}

function collectRecordFields(
  prefix: string,
  record: Record<string, unknown>,
  keys: readonly string[],
): { readonly name: string; readonly value: string }[] {
  return keys.flatMap((key) => {
    const value = readString(record, key)
    return value ? [{ name: `${prefix}.${key}`, value }] : []
  })
}
