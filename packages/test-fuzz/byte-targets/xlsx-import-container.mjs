import { importXlsx } from '@bilig/excel-import'

export function fuzz(data) {
  try {
    importXlsx(data, 'fuzz.xlsx')
  } catch (error) {
    if (!(error instanceof Error)) {
      throw error
    }
  }
}
