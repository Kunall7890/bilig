import { Component, type ErrorInfo, type ReactNode } from 'react'

interface WorkbookAppErrorBoundaryProps {
  readonly children: ReactNode
  readonly onError?: ((error: Error, info: ErrorInfo) => void) | undefined
  readonly onReload?: (() => void) | undefined
  readonly resetKey?: string | number | undefined
}

interface WorkbookAppErrorBoundaryState {
  readonly error: Error | null
}

export class WorkbookAppErrorBoundary extends Component<WorkbookAppErrorBoundaryProps, WorkbookAppErrorBoundaryState> {
  override state: WorkbookAppErrorBoundaryState = {
    error: null,
  }

  static getDerivedStateFromError(error: Error): WorkbookAppErrorBoundaryState {
    return { error }
  }

  override componentDidCatch(error: Error, info: ErrorInfo): void {
    this.props.onError?.(error, info)
  }

  override componentDidUpdate(previousProps: WorkbookAppErrorBoundaryProps): void {
    if (previousProps.resetKey !== this.props.resetKey && this.state.error) {
      this.setState({ error: null })
    }
  }

  private readonly handleReload = (): void => {
    if (this.props.onReload) {
      this.props.onReload()
      return
    }
    window.location.reload()
  }

  override render(): ReactNode {
    if (!this.state.error) {
      return this.props.children
    }

    return (
      <div className="flex h-dvh min-h-0 items-center justify-center bg-[var(--wb-app-bg)] px-6" data-testid="workbook-app-error-state">
        <div className="w-full max-w-lg rounded-[var(--wb-radius-panel)] border border-[var(--wb-border)] bg-[var(--wb-surface)] p-5 shadow-[var(--wb-shadow-sm)]">
          <div className="text-[13px] font-semibold text-[var(--wb-text)]">Workbook render failed</div>
          <div className="mt-2 text-[12px] leading-5 text-[var(--wb-text-muted)]">
            The workbook shell hit a render error before the grid could be presented.
          </div>
          <div
            className="mt-3 max-h-24 overflow-auto rounded-[var(--wb-radius-control)] border border-[var(--wb-border)] bg-[var(--wb-surface-subtle)] px-3 py-2 font-mono text-[11px] leading-4 text-[var(--wb-text-muted)]"
            data-testid="workbook-app-error-message"
          >
            {this.state.error.message}
          </div>
          <button
            className="mt-4 inline-flex h-8 items-center rounded-[var(--wb-radius-control)] border border-[var(--wb-border)] bg-[var(--wb-surface)] px-3 text-[12px] font-medium text-[var(--wb-text)] shadow-[var(--wb-shadow-sm)] transition-colors hover:border-[var(--wb-border-strong)] hover:bg-[var(--wb-muted)] focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[var(--wb-accent-ring)] focus-visible:ring-offset-1"
            data-testid="workbook-app-error-reload"
            onClick={this.handleReload}
            type="button"
          >
            Reload workbook
          </button>
        </div>
      </div>
    )
  }
}
