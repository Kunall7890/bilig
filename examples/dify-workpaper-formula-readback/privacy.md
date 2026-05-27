# Bilig WorkPaper Dify Plugin Privacy

This example plugin sends the selected input cell address, sheet name, and value
to the configured Bilig OpenAPI base URL. The example defaults to Bilig's hosted
request-local smoke endpoint:

```text
https://bilig.proompteng.ai/openapi/workpaper/set-cell-and-readback
```

The plugin does not collect API keys, user identity, files, workbook uploads, or
conversation history. Dify may store tool inputs and outputs according to the
host workspace configuration. Use a self-hosted Bilig base URL if the input data
is private.
