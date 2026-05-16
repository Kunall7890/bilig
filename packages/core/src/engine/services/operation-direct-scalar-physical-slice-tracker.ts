export type DirectScalarPhysicalSliceKind = 'input' | 'formula'

export class DirectScalarPhysicalSliceTracker {
  private trustedSheetId: number | undefined
  private trusted = true
  private previousInputRow = -1
  private previousInputCol = -1
  private previousFormulaRow = -1
  private previousFormulaCol = -1

  markUntrusted(): void {
    this.trusted = false
  }

  noteCell(sheetId: number, row: number, col: number, kind: DirectScalarPhysicalSliceKind): void {
    if (!this.trusted) {
      return
    }
    if (this.trustedSheetId === undefined) {
      this.trustedSheetId = sheetId
    } else if (this.trustedSheetId !== sheetId) {
      this.markUntrusted()
      return
    }

    if (kind === 'input') {
      if (row < this.previousInputRow || (row === this.previousInputRow && col < this.previousInputCol)) {
        this.markUntrusted()
        return
      }
      this.previousInputRow = row
      this.previousInputCol = col
      return
    }

    if (row < this.previousFormulaRow || (row === this.previousFormulaRow && col < this.previousFormulaCol)) {
      this.markUntrusted()
      return
    }
    this.previousFormulaRow = row
    this.previousFormulaCol = col
  }

  getTrustedSheetIdForTrackedChanges(explicitChangedCount: number, changedLength: number): number | undefined {
    if (!this.trusted || this.trustedSheetId === undefined || explicitChangedCount <= 0 || explicitChangedCount >= changedLength) {
      return undefined
    }
    return this.trustedSheetId
  }
}
