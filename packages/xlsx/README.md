# @bilig/xlsx

`@bilig/xlsx` provides ZIP, address, and source-preserving patch primitives used
by Bilig runtime packages. This package does not depend on SheetJS or the
`xlsx` CDN tarball.

## Source-Preserving Literal Patches

Use `exportXlsxSourceLiteralPatchesToFileAsync` when a service needs to apply
scalar cell edits to an existing XLSX package while preserving untouched package
parts and avoiding full-source reads:

```ts
import { exportXlsxSourceLiteralPatchesToFileAsync } from '@bilig/xlsx'

await exportXlsxSourceLiteralPatchesToFileAsync({
  source: fileBackedXlsxReader,
  outputPath: './patched.xlsx',
  sheetNames: ['Inputs'],
  patches: [{ sheetName: 'Inputs', address: 'B3', value: 96000 }],
})
```

The file-backed source only needs `byteLength` and `readRange(start, end)` for
the streaming patch path. `readBytes()` is a fallback, not the normal path.
