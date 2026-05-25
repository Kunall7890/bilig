---
title: Open WebUI WorkPaper MCP setup
published: true
description: Connect Bilig WorkPaper to Open WebUI through native Streamable HTTP MCP or mcpo so chats can edit workbook inputs and verify formula readback.
tags: open-webui, mcp, mcpo, spreadsheet-agent, workbook-api
canonical_url: https://proompteng.github.io/bilig/open-webui-workpaper-mcp.html
image: /assets/github-social-preview.png
---

# Open WebUI WorkPaper MCP setup

Use this when Open WebUI should call spreadsheet tools, but the spreadsheet
logic should stay in a formula-backed WorkPaper instead of an Excel browser
session.

Open WebUI has two useful integration paths:

- native **MCP (Streamable HTTP)** for an HTTP MCP endpoint;
- **mcpo** when the tool server is a local stdio MCP process that needs to be
  exposed as an OpenAPI tool server.

Official Open WebUI references:

- <https://docs.openwebui.com/features/extensibility/mcp/>
- <https://docs.openwebui.com/features/extensibility/plugin/tools/>
- <https://docs.openwebui.com/features/extensibility/plugin/tools/openapi-servers/mcp/>
- <https://docs.openwebui.com/features/extensibility/plugin/tools/openapi-servers/open-webui/>

## Fastest smoke test: hosted Streamable HTTP

Open WebUI's native MCP path can connect to a Streamable HTTP server from
**Admin Settings -> External Tools** with type **MCP (Streamable HTTP)**.

For a quick Bilig proof, add this server URL:

```text
https://bilig.proompteng.ai/mcp
```

Use **Auth: None** unless your deployment sits behind its own gateway token.

Then open a chat, enable the Bilig tool from the integrations/tools menu, and
ask:

```text
List the available Bilig WorkPaper tools. Then read the sample sheets, set the
conversion-rate input to 0.4, read the recalculated ARR output, and return the
proof object.
```

Expected tools:

- `list_sheets`
- `read_range`
- `read_cell`
- `set_cell_contents`
- `get_cell_display_value`
- `export_workpaper_document`
- `validate_formula`

The hosted endpoint is stateless and request-local. It is good for verifying
Open WebUI can discover and call the tools. It does not persist a private
project workbook.

## Persistent project file: mcpo bridge

Use `mcpo` when Open WebUI needs an OpenAPI tool server for a local stdio MCP
process. This is the right shape when the WorkPaper JSON file lives on the host
or in a container volume.

From the machine that can reach the WorkPaper file:

```sh
uvx mcpo --host 0.0.0.0 --port 8000 -- \
  npx -y --package @bilig/workpaper@0.105.0 \
    bilig-workpaper-mcp \
    --workpaper ./pricing.workpaper.json \
    --init-demo-workpaper \
    --writable
```

Open the generated tool docs:

```text
http://localhost:8000/docs
```

Then add the tool server URL in Open WebUI. For a personal user tool, the URL
can be:

```text
http://localhost:8000
```

For a global tool server configured from the Open WebUI backend, remember that
`localhost` means the Open WebUI backend container or host, not your laptop.
When Open WebUI runs in Docker and the mcpo server is on the host, use:

```text
http://host.docker.internal:8000
```

## Native MCP versus mcpo

| Path | Use when | URL to add |
| --- | --- | --- |
| Native MCP | Open WebUI can call a Streamable HTTP MCP endpoint directly. | `https://bilig.proompteng.ai/mcp` |
| mcpo | Open WebUI should call a local stdio MCP server through OpenAPI. | `http://localhost:8000` or `http://host.docker.internal:8000` |

Start with native MCP for tool-discovery smoke tests. Use mcpo for a real
writable WorkPaper file that must persist across turns or jobs.

## Proof object to ask for

Ask the model to return a concrete proof instead of "the cell was updated":

```json
{
  "editedCell": "Inputs!B3",
  "before": {
    "Summary!B2": "60000"
  },
  "after": {
    "Summary!B2": "96000"
  },
  "persistedDocumentBytes": 1000,
  "verified": true,
  "limitations": [
    "Hosted smoke endpoint is request-local.",
    "Use mcpo or local stdio for a private writable WorkPaper file."
  ]
}
```

`verified` should only be true after a readback of the dependent formula output.

## Troubleshooting

- If Open WebUI says the MCP server failed to connect, check that the tool type
  is **MCP (Streamable HTTP)**, not OpenAPI.
- If using the hosted endpoint, leave auth set to **None**.
- If using mcpo from Docker, replace `localhost` with `host.docker.internal` or
  the reachable host IP.
- If global tools do not show in chat, enable the tool from the chat
  integrations/tools picker. Global tool servers can be hidden until enabled.
- Use a model with native function calling for multi-step read/write/readback
  tool use.

## Related Bilig docs

- [MCP client setup](mcp-client-setup.md)
- [MCP WorkPaper tool server](mcp-workpaper-tool-server.md)
- [Agent framework workbook tools](agent-framework-workbook-tools.md)
- [Why agents need workbook APIs](why-agents-need-workbook-apis.md)
- [Headless WorkPaper agent handbook](headless-workpaper-agent-handbook.md)

