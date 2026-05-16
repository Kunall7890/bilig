import {
  createWorkPaperFromDocument,
  exportWorkPaperDocument,
  parseWorkPaperDocument,
  serializeWorkPaperDocument,
  type PersistedWorkPaperDocument,
} from './persistence.js'
import type { WorkPaper } from './work-paper.js'
import type { WorkPaperMcpCapabilities, WorkPaperMcpToolServer } from './work-paper-mcp-server.js'
import type { RawCellContent, WorkPaperCellAddress } from './work-paper-types.js'
import { formatCellDisplayValue } from '@bilig/protocol'
import { readFileSync, renameSync, writeFileSync } from 'node:fs'
import { basename, dirname, resolve } from 'node:path'

type JsonObject = Record<string, unknown>
type JsonRpcId = string | number | null

interface JsonRpcRequest {
  jsonrpc: '2.0'
  id: JsonRpcId | undefined
  method: string
  params?: JsonObject
}

interface JsonRpcSuccess<Result> {
  jsonrpc: '2.0'
  id: JsonRpcId | undefined
  result: Result
}

interface FileBackedWorkPaperMcpOptions {
  workbook: WorkPaper
  writable?: boolean
  sourcePath?: string
  persist?: (workbook: WorkPaper) => FileBackedWorkPaperPersistResult
}

interface FileBackedWorkPaperPersistResult {
  persisted: boolean
  path?: string
  serializedBytes: number
}

interface WorkPaperMcpToolAnnotations {
  title: string
  readOnlyHint: boolean
  destructiveHint: boolean
  idempotentHint: boolean
  openWorldHint: false
}

interface WorkPaperMcpToolDefinition {
  name:
    | 'list_sheets'
    | 'read_range'
    | 'read_cell'
    | 'set_cell_contents'
    | 'get_cell_display_value'
    | 'export_workpaper_document'
    | 'validate_formula'
  title: string
  description: string
  inputSchema: JsonObject
  annotations: WorkPaperMcpToolAnnotations
}

interface FileBackedToolCallResult {
  content: {
    type: 'text'
    text: string
  }[]
  structuredContent: JsonObject
  isError: false
}

const capabilities: WorkPaperMcpCapabilities = {
  tools: {
    listChanged: false,
  },
}

function createFileBackedWorkPaperMcpToolServer(options: FileBackedWorkPaperMcpOptions): WorkPaperMcpToolServer {
  const { workbook, writable = false, sourcePath } = options
  const persist = options.persist ?? createMemoryPersist(workbook)
  const toolDefinitions = createFileBackedToolDefinitions(writable)

  return {
    capabilities,

    handleJsonRpc(request: unknown): JsonRpcSuccess<unknown> {
      const parsedRequest = parseJsonRpcRequest(request)

      if (parsedRequest.method === 'tools/list') {
        return {
          jsonrpc: '2.0',
          id: parsedRequest.id,
          result: {
            tools: toolDefinitions,
          },
        }
      }

      if (parsedRequest.method === 'tools/call') {
        const structuredContent = callFileBackedTool({
          workbook,
          writable,
          sourcePath,
          persist,
          params: parsedRequest.params,
        })

        return {
          jsonrpc: '2.0',
          id: parsedRequest.id,
          result: toolResult(structuredContent),
        }
      }

      throw new Error(`Unsupported MCP method: ${parsedRequest.method}`)
    },
  }
}

function createFileBackedWorkPaperMcpToolServerFromFile(input: { workpaperPath: string; writable?: boolean }): WorkPaperMcpToolServer {
  const workpaperPath = resolve(input.workpaperPath)
  const workbook = createWorkPaperFromDocument(parseWorkPaperDocument(readFileSync(workpaperPath, 'utf8')))

  return createFileBackedWorkPaperMcpToolServer({
    workbook,
    writable: input.writable ?? false,
    sourcePath: workpaperPath,
    persist(updatedWorkbook) {
      const serialized = serializeWorkbook(updatedWorkbook)
      if (input.writable) {
        writeFileAtomically(workpaperPath, serialized)
      }
      return {
        persisted: input.writable ?? false,
        path: workpaperPath,
        serializedBytes: Buffer.byteLength(serialized, 'utf8'),
      }
    },
  })
}

function createFileBackedToolDefinitions(writable: boolean): WorkPaperMcpToolDefinition[] {
  return [
    {
      name: 'list_sheets',
      title: 'List WorkPaper Sheets',
      description: 'List sheets and their current used dimensions.',
      inputSchema: emptySchema(),
      annotations: readOnlyAnnotation('List WorkPaper Sheets'),
    },
    {
      name: 'read_range',
      title: 'Read WorkPaper Range',
      description: 'Read evaluated values and serialized cell contents for a WorkPaper range.',
      inputSchema: {
        type: 'object',
        required: ['range'],
        properties: {
          range: {
            type: 'string',
            description: 'A1 range. Include a sheet name or pass sheetName separately.',
          },
          sheetName: {
            type: 'string',
            description: 'Default sheet name when range omits a sheet name.',
          },
        },
        additionalProperties: false,
      },
      annotations: readOnlyAnnotation('Read WorkPaper Range'),
    },
    {
      name: 'read_cell',
      title: 'Read WorkPaper Cell',
      description: 'Read one cell with evaluated value, display value, formula, and serialized content.',
      inputSchema: cellAddressSchema(['sheetName', 'address']),
      annotations: readOnlyAnnotation('Read WorkPaper Cell'),
    },
    {
      name: 'set_cell_contents',
      title: 'Set WorkPaper Cell Contents',
      description: writable
        ? 'Set one cell, recalculate dependents, and persist the updated WorkPaper JSON file.'
        : 'Set one cell in memory. Start with --writable to persist the updated WorkPaper JSON file.',
      inputSchema: {
        type: 'object',
        required: ['sheetName', 'address', 'value'],
        properties: {
          sheetName: {
            type: 'string',
          },
          address: {
            type: 'string',
          },
          value: {
            type: ['string', 'number', 'boolean', 'null'],
            description: 'Raw cell content. Formula strings must start with =.',
          },
        },
        additionalProperties: false,
      },
      annotations: {
        title: 'Set WorkPaper Cell Contents',
        readOnlyHint: false,
        destructiveHint: writable,
        idempotentHint: true,
        openWorldHint: false,
      },
    },
    {
      name: 'get_cell_display_value',
      title: 'Get WorkPaper Cell Display Value',
      description: 'Return the formatted display value for one cell.',
      inputSchema: cellAddressSchema(['sheetName', 'address']),
      annotations: readOnlyAnnotation('Get WorkPaper Cell Display Value'),
    },
    {
      name: 'export_workpaper_document',
      title: 'Export WorkPaper Document',
      description: 'Export the current WorkPaper JSON document.',
      inputSchema: {
        type: 'object',
        properties: {
          includeConfig: {
            type: 'boolean',
            default: true,
          },
        },
        additionalProperties: false,
      },
      annotations: readOnlyAnnotation('Export WorkPaper Document'),
    },
    {
      name: 'validate_formula',
      title: 'Validate WorkPaper Formula',
      description: 'Validate formula syntax using the WorkPaper formula parser.',
      inputSchema: {
        type: 'object',
        required: ['formula'],
        properties: {
          formula: {
            type: 'string',
            description: 'Formula string, including the leading =.',
          },
        },
        additionalProperties: false,
      },
      annotations: readOnlyAnnotation('Validate WorkPaper Formula'),
    },
  ]
}

function callFileBackedTool(input: {
  workbook: WorkPaper
  writable: boolean
  sourcePath: string | undefined
  persist: (workbook: WorkPaper) => FileBackedWorkPaperPersistResult
  params: JsonObject | undefined
}): JsonObject {
  const parsedParams = requireRecord(input.params ?? {}, 'MCP tool call params')
  const toolName = parsedParams['name']
  const args = requireRecord(parsedParams['arguments'] ?? {}, `${String(toolName)} arguments`)

  if (toolName === 'list_sheets') {
    return {
      sourcePath: input.sourcePath,
      writable: input.writable,
      sheets: input.workbook.getSheetNames().map((name) => {
        const sheetId = requireSheet(input.workbook, name)
        return {
          id: sheetId,
          name,
          dimensions: input.workbook.getSheetDimensions(sheetId),
        }
      }),
    }
  }

  if (toolName === 'read_range') {
    const range = requireString(args['range'], 'range')
    const defaultSheet = optionalSheetId(input.workbook, args['sheetName'])
    const parsedRange = input.workbook.simpleCellRangeFromString(range, defaultSheet)
    if (parsedRange === undefined) {
      throw new Error(`Invalid range: ${range}`)
    }
    return {
      range: input.workbook.simpleCellRangeToString(parsedRange, { includeSheetName: true }),
      values: input.workbook.getRangeValues(parsedRange),
      serialized: input.workbook.getRangeSerialized(parsedRange),
    }
  }

  if (toolName === 'read_cell') {
    return readCell(input.workbook, parseCellArgs(input.workbook, args))
  }

  if (toolName === 'set_cell_contents') {
    const address = parseCellArgs(input.workbook, args)
    const value = parseRawCellContent(args['value'])
    const before = readCell(input.workbook, address)

    input.workbook.setCellContents(address, value)

    const after = readCell(input.workbook, address)
    const persistence = input.persist(input.workbook)
    const restored = createWorkPaperFromDocument(parseWorkPaperDocument(serializeWorkbook(input.workbook)))
    const restoredAddress = requireCellAddress(
      restored,
      requireString(args['sheetName'], 'sheetName'),
      requireString(args['address'], 'address'),
    )
    const restoredCell = readCell(restored, restoredAddress)

    return {
      editedCell: input.workbook.simpleCellAddressToString(address, { includeSheetName: true }),
      before,
      after,
      restored: restoredCell,
      persistence,
      checks: {
        persisted: persistence.persisted,
        restoredMatchesAfter: JSON.stringify(after) === JSON.stringify(restoredCell),
        previousSerialized: before['serialized'],
        newSerialized: after['serialized'],
      },
    }
  }

  if (toolName === 'get_cell_display_value') {
    const cell = readCell(input.workbook, parseCellArgs(input.workbook, args))
    return {
      address: cell['address'],
      displayValue: cell['displayValue'],
    }
  }

  if (toolName === 'export_workpaper_document') {
    const includeConfig = args['includeConfig'] === undefined ? true : requireBoolean(args['includeConfig'], 'includeConfig')
    const document = exportWorkPaperDocument(input.workbook, { includeConfig })
    const serialized = serializeWorkPaperDocument(document)
    return {
      sourcePath: input.sourcePath,
      document,
      serializedBytes: Buffer.byteLength(serialized, 'utf8'),
    }
  }

  if (toolName === 'validate_formula') {
    const formula = requireString(args['formula'], 'formula')
    return {
      formula,
      valid: input.workbook.validateFormula(formula),
    }
  }

  throw new Error(`Unknown WorkPaper tool: ${String(toolName)}`)
}

function readCell(workbook: WorkPaper, address: WorkPaperCellAddress): JsonObject {
  const value = workbook.getCellValue(address)
  const format = workbook.getCellValueFormat(address)
  return {
    address: workbook.simpleCellAddressToString(address, { includeSheetName: true }),
    value,
    serialized: workbook.getCellSerialized(address),
    formula: workbook.getCellFormula(address),
    displayValue: formatCellDisplayValue(value, format),
  }
}

function toolResult(structuredContent: JsonObject): FileBackedToolCallResult {
  return {
    content: [
      {
        type: 'text',
        text: JSON.stringify(structuredContent),
      },
    ],
    structuredContent,
    isError: false,
  }
}

function parseJsonRpcRequest(value: unknown): JsonRpcRequest {
  const request = requireRecord(value, 'JSON-RPC request')
  if (request['jsonrpc'] !== '2.0' || typeof request['method'] !== 'string') {
    throw new Error('Expected JSON-RPC 2.0 request')
  }

  const id = request['id']
  if (id !== undefined && id !== null && typeof id !== 'string' && typeof id !== 'number') {
    throw new Error(`Unsupported JSON-RPC id: ${JSON.stringify(id)}`)
  }

  return {
    jsonrpc: '2.0',
    id,
    method: request['method'],
    ...(request['params'] !== undefined ? { params: requireRecord(request['params'], 'JSON-RPC params') } : {}),
  }
}

function parseCellArgs(workbook: WorkPaper, args: JsonObject): WorkPaperCellAddress {
  return requireCellAddress(workbook, requireString(args['sheetName'], 'sheetName'), requireString(args['address'], 'address'))
}

function requireCellAddress(workbook: WorkPaper, sheetName: string, a1Address: string): WorkPaperCellAddress {
  const sheetId = requireSheet(workbook, sheetName)
  const parsed = workbook.simpleCellAddressFromString(a1Address, sheetId)
  if (parsed === undefined || parsed.sheet !== sheetId) {
    throw new Error(`Invalid cell address: ${sheetName}!${a1Address}`)
  }
  return parsed
}

function optionalSheetId(workbook: WorkPaper, value: unknown): number | undefined {
  if (value === undefined) {
    return undefined
  }
  return requireSheet(workbook, requireString(value, 'sheetName'))
}

function requireSheet(workbook: WorkPaper, sheetName: string): number {
  const sheetId = workbook.getSheetId(sheetName)
  if (sheetId === undefined) {
    throw new Error(`Expected sheet "${sheetName}" to exist`)
  }
  return sheetId
}

function parseRawCellContent(value: unknown): RawCellContent {
  if (value !== null && typeof value !== 'string' && typeof value !== 'number' && typeof value !== 'boolean') {
    throw new Error(`Unsupported cell value: ${JSON.stringify(value)}`)
  }
  return value
}

function requireRecord(value: unknown, label: string): JsonObject {
  if (!isRecord(value)) {
    throw new Error(`Expected ${label} to be an object`)
  }
  return value
}

function requireString(value: unknown, label: string): string {
  if (typeof value !== 'string') {
    throw new Error(`Expected ${label} to be a string`)
  }
  return value
}

function requireBoolean(value: unknown, label: string): boolean {
  if (typeof value !== 'boolean') {
    throw new Error(`Expected ${label} to be a boolean`)
  }
  return value
}

function isRecord(value: unknown): value is JsonObject {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
}

function serializeWorkbook(workbook: WorkPaper): string {
  return serializeWorkPaperDocument(exportWorkPaperDocument(workbook, { includeConfig: true }))
}

function createMemoryPersist(workbook: WorkPaper): () => FileBackedWorkPaperPersistResult {
  return () => {
    const serialized = serializeWorkbook(workbook)
    return {
      persisted: false,
      serializedBytes: Buffer.byteLength(serialized, 'utf8'),
    }
  }
}

function writeFileAtomically(path: string, contents: string): void {
  const tempPath = resolve(dirname(path), `.${basename(path)}.${process.pid.toString()}.tmp`)
  writeFileSync(tempPath, contents)
  renameSync(tempPath, path)
}

function emptySchema(): JsonObject {
  return {
    type: 'object',
    additionalProperties: false,
  }
}

function cellAddressSchema(required: string[]): JsonObject {
  return {
    type: 'object',
    required,
    properties: {
      sheetName: {
        type: 'string',
      },
      address: {
        type: 'string',
      },
    },
    additionalProperties: false,
  }
}

function readOnlyAnnotation(title: string): WorkPaperMcpToolAnnotations {
  return {
    title,
    readOnlyHint: true,
    destructiveHint: false,
    idempotentHint: true,
    openWorldHint: false,
  }
}

export {
  createFileBackedWorkPaperMcpToolServer,
  createFileBackedWorkPaperMcpToolServerFromFile,
  type FileBackedWorkPaperMcpOptions,
  type FileBackedWorkPaperPersistResult,
  type PersistedWorkPaperDocument,
}
