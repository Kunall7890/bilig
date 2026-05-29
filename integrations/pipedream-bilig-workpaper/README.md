# Bilig WorkPaper for Pipedream

This directory stages the Pipedream registry-style action for Bilig WorkPaper.

The first action is intentionally narrow: it writes one forecast input cell,
asks Bilig to recalculate dependent formulas, and returns proof that the
computed value changed and the exported WorkPaper JSON restores to the same
result.

It is meant for workflow builders who need spreadsheet formulas in an
automation without launching Excel, LibreOffice, Google Sheets, or a browser UI.

## Action

```text
components/bilig/actions/verify-formula-readback/verify-formula-readback.mjs
```

Default endpoint:

```text
POST https://bilig.proompteng.ai/api/workpaper/n8n/forecast
```

## Local Checks

```sh
node --check components/bilig/bilig.app.mjs
node --check components/bilig/actions/verify-formula-readback/verify-formula-readback.mjs
curl -sS -X POST https://bilig.proompteng.ai/api/workpaper/n8n/forecast \
  -H 'content-type: application/json' \
  --data '{"sheetName":"Inputs","address":"B3","value":0.4}'
```

## Pipedream Path

Pipedream registry components live under `components/[app-slug]` in the
`PipedreamHQ/pipedream` repository. The public review is:

```text
https://github.com/PipedreamHQ/pipedream/pull/20972
```

The reviewed upstream Pipedream action uses the app slug `bilig_workpaper`,
stores the base URL in the connected account as `$auth.base_url`, and does not
expose Base URL as an action prop.

For private testing, install the Pipedream CLI, log in, then publish the action:

```sh
pd publish components/bilig/actions/verify-formula-readback/verify-formula-readback.mjs
```

Use the hosted endpoint for a proof run, or point the connected account Base URL
at a self-hosted Bilig app.
