import type { CSSProperties } from 'react'

export const agentPanelThemeStyle: CSSProperties & Record<`--${string}`, string> = {
  '--wb-app-bg': 'var(--color-mauve-50)',
  '--wb-surface': 'white',
  '--wb-surface-subtle': 'var(--color-mauve-50)',
  '--wb-surface-muted': 'var(--color-mauve-100)',
  '--wb-border': 'var(--color-mauve-200)',
  '--wb-border-strong': 'var(--color-mauve-300)',
  '--wb-grid-border': 'var(--color-mauve-100)',
  '--wb-text': 'var(--color-mauve-900)',
  '--wb-text-muted': 'var(--color-mauve-700)',
  '--wb-text-subtle': 'var(--color-mauve-600)',
  '--wb-accent': 'var(--color-mauve-900)',
  '--wb-accent-soft': 'var(--color-mauve-100)',
  '--wb-accent-ring': 'var(--color-mauve-400)',
  '--wb-hover': 'var(--color-mauve-100)',
  '--wb-shadow-sm': '0 1px 2px rgba(15, 23, 42, 0.04)',
}
