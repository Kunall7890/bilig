export const proofFieldNames = [
  'editedCell',
  'before',
  'after',
  'afterRestore',
  'persistedDocumentBytes',
  'verified',
  'limitations',
] as const

export interface AgentDiscoveryEvaluatorDoorInputs {
  readonly repositoryUrl: string
  readonly siteRoot: string
  readonly workpaperPackageSpec: string
}

export interface AgentDiscoveryEvaluatorDoor {
  readonly name: string
  readonly audience: string
  readonly docs: string
  readonly source: string
  readonly package: string
  readonly command: string
  readonly expected_result: string
  readonly proof_schema?: string
  readonly canonical_door?: string
}

export interface AgentDiscoveryProofContract {
  readonly rule: string
  readonly fields: readonly string[]
  readonly accepted_aliases: {
    readonly afterRestore: readonly string[]
    readonly persistedDocumentBytes: readonly string[]
  }
  readonly required_steps: readonly string[]
}

export type CompactStringArrayProperty = (json: string, propertyName: string, values: readonly string[], indent: string) => string

export function buildEvaluatorDoors(args: AgentDiscoveryEvaluatorDoorInputs): readonly AgentDiscoveryEvaluatorDoor[] {
  const siteRoot = args.siteRoot.replace(/\/+$/, '')
  return [
    {
      name: 'eval-xlsx-cache-doctor',
      audience: 'A repo, service, CI job, or agent has XLSX files that may contain stale cached formula values.',
      docs: `${siteRoot}/eval-xlsx-cache-doctor.html`,
      source: `${args.repositoryUrl}/blob/main/docs/eval-xlsx-cache-doctor.md`,
      package: '@bilig/xlsx-formula-recalc',
      command: 'npm exec --yes --package @bilig/xlsx-formula-recalc@latest -- bilig-evaluate --door xlsx-cache --json',
      expected_result:
        'bilig-evaluator.v1 JSON with stale cached formula count, suggested reads, recalculated value, and verified true without Excel or LibreOffice',
      proof_schema: 'bilig-evaluator.v1',
      canonical_door: 'xlsx-cache',
    },
    {
      name: 'eval-xlsx-recalc',
      audience: 'A service has an XLSX file with stale formula results after editing inputs in Node.js.',
      docs: `${siteRoot}/eval-xlsx-recalc.html`,
      source: `${args.repositoryUrl}/blob/main/docs/eval-xlsx-recalc.md`,
      package: '@bilig/xlsx-formula-recalc',
      command:
        'npx --yes --package @bilig/xlsx-formula-recalc xlsx-recalc pricing.xlsx --set Inputs!B2=48 --read Summary!B7 --out pricing.recalculated.xlsx --json',
      expected_result: 'input edit, recalculated readback, output workbook path, and verified true without Excel or LibreOffice',
    },
    {
      name: 'eval-workpaper-service',
      audience: 'A Node.js service, route, queue worker, or test needs workbook-shaped formula logic as JSON state.',
      docs: `${siteRoot}/eval-workpaper-service.html`,
      source: `${args.repositoryUrl}/blob/main/docs/eval-workpaper-service.md`,
      package: '@bilig/workpaper',
      command: `npm exec --yes --package ${args.workpaperPackageSpec} -- bilig-evaluate --door workpaper-service --json`,
      expected_result:
        'bilig-evaluator.v1 JSON with direct WorkPaper edit, calculated readback, JSON persistence, restore check, and verified true',
      proof_schema: 'bilig-evaluator.v1',
      canonical_door: 'workpaper-service',
    },
    {
      name: 'eval-agent-mcp',
      audience: 'A coding agent, MCP client, or tool host needs spreadsheet operations without UI automation.',
      docs: `${siteRoot}/eval-agent-mcp.html`,
      source: `${args.repositoryUrl}/blob/main/docs/eval-agent-mcp.md`,
      package: '@bilig/workpaper',
      command: `npm exec --yes --package ${args.workpaperPackageSpec} -- bilig-evaluate --door agent-mcp --json`,
      expected_result:
        'bilig-evaluator.v1 JSON with MCP tool discovery, cell edit, formula readback, export, restore, tool list, and verified true',
      proof_schema: 'bilig-evaluator.v1',
      canonical_door: 'agent-mcp',
    },
  ] as const
}

export function buildProofContract(): AgentDiscoveryProofContract {
  return {
    rule: 'Do not claim spreadsheet or workbook success from a write call alone.',
    fields: proofFieldNames,
    accepted_aliases: {
      afterRestore: ['afterRestart'],
      persistedDocumentBytes: ['persistence.serializedBytes', 'serializedBytes'],
    },
    required_steps: [
      'read the relevant range before editing',
      'write the target input or formula cell',
      'read the dependent calculated output after recalculation',
      'export or serialize the WorkPaper document',
      'restore or reimport when a file boundary matters',
      'return editedCell, before, after, afterRestore, persistedDocumentBytes, verified, and limitations',
    ],
  }
}

export function compactProofContractJsonArrays(json: string, compactStringArrayProperty: CompactStringArrayProperty): string {
  const proofContract = buildProofContract()
  const compactProofFields = compactStringArrayProperty(json, 'fields', proofContract.fields, '    ')
  const compactAfterRestoreAlias = compactStringArrayProperty(
    compactProofFields,
    'afterRestore',
    proofContract.accepted_aliases.afterRestore,
    '      ',
  )
  return compactStringArrayProperty(
    compactAfterRestoreAlias,
    'persistedDocumentBytes',
    proofContract.accepted_aliases.persistedDocumentBytes,
    '      ',
  )
}

export function requireAgentJsonDiscoveryContract(
  args: AgentDiscoveryEvaluatorDoorInputs & { readonly parsedAgentJson: unknown; readonly skillManifestUrl: string },
): void {
  const parsedAgentJson = args.parsedAgentJson
  if (typeof parsedAgentJson !== 'object' || parsedAgentJson === null || Array.isArray(parsedAgentJson)) {
    throw new Error('docs/.well-known/agent.json must be a JSON object')
  }

  const siteRoot = args.siteRoot.replace(/\/+$/, '')
  for (const [fieldName, expectedValue] of [
    ['name', 'bilig'],
    ['repository', args.repositoryUrl],
    ['llms_txt', `${siteRoot}/llms.txt`],
    ['llms_full', `${siteRoot}/llms-full.txt`],
    ['skill_file', args.skillManifestUrl],
    ['agent_instructions', `${siteRoot}/AGENTS.md`],
  ] as const) {
    if (Reflect.get(parsedAgentJson, fieldName) !== expectedValue) {
      throw new Error(`docs/.well-known/agent.json ${fieldName} must be ${expectedValue}`)
    }
  }

  const agentJsonEvaluatorDoors = Reflect.get(parsedAgentJson, 'evaluator_doors')
  if (!Array.isArray(agentJsonEvaluatorDoors) || !agentJsonEvaluatorDoors.every((door) => typeof door === 'object' && door !== null)) {
    throw new Error('docs/.well-known/agent.json evaluator_doors must be an object array')
  }

  for (const requiredDoor of buildEvaluatorDoors(args)) {
    const door = agentJsonEvaluatorDoors.find((candidate) => Reflect.get(candidate, 'name') === requiredDoor.name)
    if (door === undefined) {
      throw new Error(`docs/.well-known/agent.json evaluator_doors is missing ${requiredDoor.name}`)
    }
    const command = Reflect.get(door, 'command')
    if (Reflect.get(door, 'docs') !== requiredDoor.docs || command !== requiredDoor.command) {
      throw new Error(`docs/.well-known/agent.json evaluator_doors has invalid routing for ${requiredDoor.name}`)
    }
    if (requiredDoor.proof_schema !== undefined && Reflect.get(door, 'proof_schema') !== requiredDoor.proof_schema) {
      throw new Error(`docs/.well-known/agent.json evaluator_doors has invalid proof schema for ${requiredDoor.name}`)
    }
    if (requiredDoor.canonical_door !== undefined && Reflect.get(door, 'canonical_door') !== requiredDoor.canonical_door) {
      throw new Error(`docs/.well-known/agent.json evaluator_doors has invalid canonical door for ${requiredDoor.name}`)
    }
  }

  const agentJsonProofContract = Reflect.get(parsedAgentJson, 'proof_contract')
  if (typeof agentJsonProofContract !== 'object' || agentJsonProofContract === null || Array.isArray(agentJsonProofContract)) {
    throw new Error('docs/.well-known/agent.json must define proof_contract')
  }
  const agentJsonProofFields = Reflect.get(agentJsonProofContract, 'fields')
  if (!Array.isArray(agentJsonProofFields) || !agentJsonProofFields.every((field) => typeof field === 'string')) {
    throw new Error('docs/.well-known/agent.json proof_contract.fields must be a string array')
  }
  for (const requiredField of proofFieldNames) {
    if (!agentJsonProofFields.includes(requiredField)) {
      throw new Error(`docs/.well-known/agent.json proof_contract.fields is missing ${requiredField}`)
    }
  }

  const agentJsonProofAliases = Reflect.get(agentJsonProofContract, 'accepted_aliases')
  if (typeof agentJsonProofAliases !== 'object' || agentJsonProofAliases === null || Array.isArray(agentJsonProofAliases)) {
    throw new Error('docs/.well-known/agent.json proof_contract.accepted_aliases must be an object')
  }
  const afterRestoreAliases = Reflect.get(agentJsonProofAliases, 'afterRestore')
  if (!Array.isArray(afterRestoreAliases) || !afterRestoreAliases.includes('afterRestart')) {
    throw new Error('docs/.well-known/agent.json proof_contract.accepted_aliases.afterRestore must include afterRestart')
  }
  const persistedBytesAliases = Reflect.get(agentJsonProofAliases, 'persistedDocumentBytes')
  if (!Array.isArray(persistedBytesAliases) || !persistedBytesAliases.includes('persistence.serializedBytes')) {
    throw new Error(
      'docs/.well-known/agent.json proof_contract.accepted_aliases.persistedDocumentBytes must include persistence.serializedBytes',
    )
  }
}
