export const workbookTextControlProps = {
  'data-workbook-text-control': 'true',
} as const

export function isWorkbookTextControlElement(element: Element | null): element is HTMLInputElement | HTMLTextAreaElement {
  return (
    (element instanceof HTMLInputElement || element instanceof HTMLTextAreaElement) && element.dataset['workbookTextControl'] === 'true'
  )
}
