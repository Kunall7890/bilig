import path from 'node:path'

export type ProtocolEnumManifest = Record<string, readonly (readonly [string, number])[]> & {
  readonly Opcode: readonly (readonly [string, number])[]
}

export interface ProtocolBuiltinManifestEntry {
  readonly id: string
  readonly name: string
  readonly supportsWasm: boolean
}

export interface RunProtocolGeneratorArgs {
  readonly repoRoot: string
  readonly enumManifest: ProtocolEnumManifest
  readonly builtinManifest: readonly ProtocolBuiltinManifestEntry[]
}

const generatedHeader = `// GENERATED FILE. DO NOT EDIT DIRECTLY.\n// Source: scripts/gen-protocol.ts\n\n`

function renderEnum(name: string, entries: readonly (readonly [string, number])[]): string {
  const lines = entries.map(([key, value]) => `  ${key} = ${value}`)
  return `export enum ${name} {\n${lines.join(',\n')}\n}\n`
}

function renderProtocolEnums(enumManifest: ProtocolEnumManifest): string {
  return (
    generatedHeader +
    Object.entries(enumManifest)
      .map(([name, entries]) => renderEnum(name, entries))
      .join('\n')
  )
}

function renderOpcodeNames(enumManifest: ProtocolEnumManifest): string {
  return enumManifest.Opcode.map(([name]) => `  [Opcode.${name}]: "${name}"`).join(',\n')
}

function renderBuiltins(builtinManifest: readonly ProtocolBuiltinManifestEntry[]): string {
  return builtinManifest
    .map(({ id, name, supportsWasm }) => `  { id: BuiltinId.${id}, name: "${name}", supportsWasm: ${supportsWasm} }`)
    .join(',\n')
}

function renderOpcodesModule(enumManifest: ProtocolEnumManifest, builtinManifest: readonly ProtocolBuiltinManifestEntry[]): string {
  return `${generatedHeader}import { BuiltinId, Opcode } from "./enums.js";

export interface BuiltinDescriptor {
  readonly id: BuiltinId;
  readonly name: string;
  readonly supportsWasm: boolean;
}

export const OPCODE_NAMES: Record<Opcode, string> = {
${renderOpcodeNames(enumManifest)}
};

export const BUILTINS: BuiltinDescriptor[] = [
${renderBuiltins(builtinManifest)}
];
`
}

export async function runProtocolGenerator(args: RunProtocolGeneratorArgs): Promise<void> {
  const generatedFiles = [
    {
      path: path.join(args.repoRoot, 'packages/protocol/src/enums.ts'),
      contents: renderProtocolEnums(args.enumManifest),
    },
    {
      path: path.join(args.repoRoot, 'packages/protocol/src/opcodes.ts'),
      contents: renderOpcodesModule(args.enumManifest, args.builtinManifest),
    },
    {
      path: path.join(args.repoRoot, 'packages/wasm-kernel/assembly/protocol.ts'),
      contents: renderProtocolEnums(args.enumManifest),
    },
  ]

  const checkMode = Bun.argv.includes('--check')
  const staleFiles = (
    await Promise.all(
      generatedFiles.map(async (file) => {
        let existing = ''
        try {
          existing = await Bun.file(file.path).text()
        } catch {
          existing = ''
        }

        if (existing === file.contents) {
          return null
        }
        if (!checkMode) {
          await Bun.write(file.path, file.contents)
        }
        return path.relative(args.repoRoot, file.path)
      }),
    )
  ).filter((entry) => entry !== null)

  if (checkMode && staleFiles.length > 0) {
    console.error(`Protocol artifacts are stale:\n${staleFiles.map((entry) => `- ${entry}`).join('\n')}`)
    process.exitCode = 1
    return
  }

  if (!checkMode) {
    if (staleFiles.length === 0) {
      console.log('Protocol artifacts are already up to date.')
      return
    }
    console.log(`Updated protocol artifacts:\n${staleFiles.map((entry) => `- ${entry}`).join('\n')}`)
  }
}
