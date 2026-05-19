#!/usr/bin/env node

import { readFileSync, writeFileSync } from 'node:fs'
import { dirname, extname, join, basename } from 'node:path'

import type { RawCellContent } from '@bilig/headless'
import { recalculateXlsx, type XlsxFormulaRecalcEdit } from './index.js'

interface CliOptions {
  readonly inputPath: string
  readonly outputPath: string
  readonly edits: readonly XlsxFormulaRecalcEdit[]
  readonly reads: readonly string[]
  readonly json: boolean
}

try {
  const options = parseCliArgs(process.argv.slice(2))
  const result = recalculateXlsx(readFileSync(options.inputPath), {
    fileName: basename(options.inputPath),
    edits: options.edits,
    reads: options.reads,
  })
  writeFileSync(options.outputPath, result.xlsx)

  const summary = {
    input: options.inputPath,
    output: options.outputPath,
    edits: options.edits.length,
    reads: result.reads,
    warnings: result.warnings,
    verified: true,
  }

  if (options.json) {
    writeStdout(`${JSON.stringify(summary, null, 2)}\n`)
  } else {
    writeStdout(`Recalculated ${options.inputPath} -> ${options.outputPath}\n`)
    for (const [target, value] of Object.entries(result.reads)) {
      writeStdout(`${target}: ${JSON.stringify(value)}\n`)
    }
    if (result.warnings.length > 0) {
      writeStdout(`Warnings: ${result.warnings.length.toString()}\n`)
    }
  }
} catch (error) {
  process.exitCode = 1
  writeStderr(`${error instanceof Error ? error.message : String(error)}\n`)
}

function parseCliArgs(args: readonly string[]): CliOptions {
  if (args.length === 0 || args.includes('--help') || args.includes('-h')) {
    printHelp()
    process.exit(0)
  }

  const inputPath = args[0]
  if (!inputPath || inputPath.startsWith('-')) {
    throw new Error('Expected input XLSX path')
  }

  const edits: XlsxFormulaRecalcEdit[] = []
  const reads: string[] = []
  let outputPath: string | undefined
  let json = false

  for (let index = 1; index < args.length; index += 1) {
    const arg = args[index]
    if (arg === undefined) {
      throw new Error('Unexpected missing xlsx-recalc argument')
    }
    switch (arg) {
      case '--set':
        edits.push(parseEdit(requireNextArg(args, index, '--set')))
        index += 1
        break
      case '--read':
        reads.push(requireNextArg(args, index, '--read'))
        index += 1
        break
      case '--out':
      case '-o':
        outputPath = requireNextArg(args, index, arg)
        index += 1
        break
      case '--json':
        json = true
        break
      default:
        throw new Error(`Unknown xlsx-recalc option: ${arg}`)
    }
  }

  return {
    inputPath,
    outputPath: outputPath ?? defaultOutputPath(inputPath),
    edits,
    reads,
    json,
  }
}

function parseEdit(raw: string): XlsxFormulaRecalcEdit {
  const separator = raw.indexOf('=')
  if (separator <= 0) {
    throw new Error(`Expected --set value in Target=Value form, received: ${raw}`)
  }
  return {
    target: raw.slice(0, separator),
    value: parseRawCellContent(raw.slice(separator + 1)),
  }
}

function parseRawCellContent(raw: string): RawCellContent {
  if (raw === 'null') {
    return null
  }
  if (raw === 'true') {
    return true
  }
  if (raw === 'false') {
    return false
  }
  if (/^-?(?:\d+|\d*\.\d+)(?:e[+-]?\d+)?$/iu.test(raw)) {
    return Number(raw)
  }
  return raw
}

function requireNextArg(args: readonly string[], index: number, option: string): string {
  const value = args[index + 1]
  if (!value || value.startsWith('--')) {
    throw new Error(`Expected value after ${option}`)
  }
  return value
}

function defaultOutputPath(inputPath: string): string {
  const extension = extname(inputPath)
  const base = extension.length > 0 ? basename(inputPath, extension) : basename(inputPath)
  return join(dirname(inputPath), `${base}.recalculated${extension || '.xlsx'}`)
}

function printHelp(): void {
  writeStdout(`Usage: xlsx-recalc <input.xlsx> [options]

Options:
  --set <Sheet!A1=value>  Edit an input cell before recalculation. Repeatable.
  --read <Sheet!A1>       Read a recalculated cell after edits. Repeatable.
  --out, -o <path>        Output XLSX path. Defaults to <input>.recalculated.xlsx.
  --json                  Print a JSON summary.
  --help, -h              Show this help.
`)
}

function writeStdout(text: string): void {
  process.stdout.write(text)
}

function writeStderr(text: string): void {
  process.stderr.write(text)
}
