import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import { describe, expect, it } from 'vitest'

import { createN8nForecastProof } from '@bilig/headless'

const workflowDir = join(process.cwd(), 'examples', 'n8n-workpaper-formula-readback')
const n8nNodeDir = join(process.cwd(), 'integrations', 'n8n-nodes-workpaper')

describe('n8n WorkPaper formula readback workflow', () => {
  it('keeps the hosted workflow on the four-node built-in formula proof path', () => {
    const workflow = readWorkflow('bilig-workpaper-formula-readback.n8n.json')
    const executableNodes = workflow.nodes.filter((node) => node.type !== 'n8n-nodes-base.stickyNote')

    expect(executableNodes.map((node) => `${node.name}:${node.type}`)).toEqual([
      'Manual Trigger:n8n-nodes-base.manualTrigger',
      'Choose forecast input:n8n-nodes-base.code',
      'Call Bilig WorkPaper:n8n-nodes-base.httpRequest',
      'Verify formula readback:n8n-nodes-base.code',
    ])
    expect(workflow.nodes.every((node) => node.type.startsWith('n8n-nodes-base.'))).toBe(true)
    expect(Object.keys(workflow.connections)).toEqual(['Manual Trigger', 'Choose forecast input', 'Call Bilig WorkPaper'])

    const chooseInputCode = getRequiredCode(workflow, 'Choose forecast input')
    expect(chooseInputCode).toContain("baseUrl: 'https://bilig.proompteng.ai'")
    expect(chooseInputCode).toContain("sheetName: 'Inputs'")
    expect(chooseInputCode).toContain("address: 'B3'")
    expect(chooseInputCode).toContain('value: 0.4')

    const requestNode = getRequiredNode(workflow, 'Call Bilig WorkPaper')
    expect(readStringParameter(requestNode, 'method')).toBe('POST')
    expect(readStringParameter(requestNode, 'url')).toBe("={{ $json.baseUrl.replace(/\\/$/, '') + '/api/workpaper/n8n/forecast' }}")
    expect(readStringParameter(requestNode, 'jsonBody')).toContain('sheetName: $json.sheetName, address: $json.address, value: $json.value')

    expectVerificationCodeReturnsCompactProof(getRequiredCode(workflow, 'Verify formula readback'))
  })

  it('keeps the self-hosted workflow on the same compact proof shape with local fallback', () => {
    const workflow = readWorkflow('bilig-workpaper-formula-readback.self-hosted.n8n.json')

    expect(workflow.nodes.every((node) => node.type.startsWith('n8n-nodes-base.'))).toBe(true)
    expect(workflow.nodes.map((node) => `${node.name}:${node.type}`)).toContain('Already verified?:n8n-nodes-base.if')
    expect(getNodesByType(workflow, 'n8n-nodes-base.httpRequest').map((node) => node.name)).toEqual([
      'Call local Bilig WorkPaper',
      'Call localhost fallback',
    ])

    const chooseInputCode = getRequiredCode(workflow, 'Choose local forecast input')
    expect(chooseInputCode).toContain("baseUrl: 'http://host.docker.internal:4321'")
    expect(chooseInputCode).toContain("fallbackBaseUrl: 'http://localhost:4321'")

    for (const requestNodeName of ['Call local Bilig WorkPaper', 'Call localhost fallback']) {
      const requestNode = getRequiredNode(workflow, requestNodeName)
      expect(readStringParameter(requestNode, 'method')).toBe('POST')
      expect(readStringParameter(requestNode, 'url')).toBe("={{ $json.baseUrl.replace(/\\/$/, '') + '/api/workpaper/n8n/forecast' }}")
    }

    const verificationCode = getRequiredCode(workflow, 'Verify local formula readback')
    expectVerificationCodeReturnsCompactProof(verificationCode)
    expect(verificationCode).toContain('dataBoundary')
    expect(verificationCode).not.toContain('verified-local')
  })

  it('keeps the native community-node workflow on the Bilig node path', () => {
    const workflow = readWorkflow('bilig-workpaper-native-node.n8n.json')
    const executableNodes = workflow.nodes.filter((node) => node.type !== 'n8n-nodes-base.stickyNote')

    expect(executableNodes.map((node) => `${node.name}:${node.type}`)).toEqual([
      'Manual Trigger:n8n-nodes-base.manualTrigger',
      'Verify forecast formula:@bilig/n8n-nodes-workpaper.biligWorkpaper',
      'Check forecast proof:n8n-nodes-base.code',
      'Evaluate WorkPaper JSON:@bilig/n8n-nodes-workpaper.biligWorkpaper',
      'Check WorkPaper JSON proof:n8n-nodes-base.code',
    ])
    expect(workflow.nodes.some((node) => node.type === 'n8n-nodes-base.httpRequest')).toBe(false)
    expect(Object.keys(workflow.connections)).toEqual([
      'Manual Trigger',
      'Verify forecast formula',
      'Check forecast proof',
      'Evaluate WorkPaper JSON',
    ])

    const forecastNode = getRequiredNode(workflow, 'Verify forecast formula')
    expect(readStringParameter(forecastNode, 'resource')).toBe('forecast')
    expect(readStringParameter(forecastNode, 'operation')).toBe('verifyReadback')
    expect(readStringParameter(forecastNode, 'baseUrl')).toBe('https://bilig.proompteng.ai')
    expect(readStringParameter(forecastNode, 'sheetName')).toBe('Inputs')
    expect(readStringParameter(forecastNode, 'address')).toBe('B3')
    expect(readNumberParameter(forecastNode, 'value')).toBe(0.4)

    const documentNode = getRequiredNode(workflow, 'Evaluate WorkPaper JSON')
    expect(readStringParameter(documentNode, 'resource')).toBe('workpaper')
    expect(readStringParameter(documentNode, 'operation')).toBe('evaluateDocument')
    expect(readStringParameter(documentNode, 'document')).toContain('bilig.headless.work-paper.document.v1')
    expect(readStringParameter(documentNode, 'edits')).toBe('[{"cell":"Inputs!B2","value":0.4}]')
    expect(readStringParameter(documentNode, 'readCells')).toBe('Summary!B2')
    expect(readBooleanParameter(documentNode, 'includeUpdatedDocument')).toBe(true)

    const forecastCheckCode = getRequiredCode(workflow, 'Check forecast proof')
    expect(forecastCheckCode).toContain('proof.verified !== true')
    expect(forecastCheckCode).toContain("operation: 'forecast'")
    expect(forecastCheckCode).toContain('afterExpectedArr: proof.after.expectedArr')

    const documentCheckCode = getRequiredCode(workflow, 'Check WorkPaper JSON proof')
    expect(documentCheckCode).toContain('proof.verified !== true')
    expect(documentCheckCode).toContain("operation: 'workpaper-json'")
    expect(documentCheckCode).toContain("proof.updatedDocument?.format !== 'bilig.headless.work-paper.document.v1'")
  })

  it('keeps the template-library workflow use-case driven instead of a thin endpoint call', () => {
    const workflow = readWorkflow('bilig-workpaper-forecast-approval-guard.n8n.json')
    const executableNodes = workflow.nodes.filter((node) => node.type !== 'n8n-nodes-base.stickyNote')

    expect(executableNodes.map((node) => `${node.name}:${node.type}`)).toEqual([
      'Manual Trigger:n8n-nodes-base.manualTrigger',
      'Load forecast requests:n8n-nodes-base.code',
      'Validate request rows:n8n-nodes-base.code',
      'Build Bilig request:n8n-nodes-base.code',
      'Call Bilig WorkPaper:n8n-nodes-base.httpRequest',
      'Verify formula proof:n8n-nodes-base.code',
      'Apply approval policy:n8n-nodes-base.code',
      'Approved?:n8n-nodes-base.if',
      'Build approval record:n8n-nodes-base.code',
      'Build review record:n8n-nodes-base.code',
      'Write audit summary:n8n-nodes-base.code',
    ])
    expect(executableNodes.length).toBeGreaterThan(10)
    expect(Object.keys(workflow.connections)).toEqual([
      'Manual Trigger',
      'Load forecast requests',
      'Validate request rows',
      'Build Bilig request',
      'Call Bilig WorkPaper',
      'Verify formula proof',
      'Apply approval policy',
      'Approved?',
      'Build approval record',
      'Build review record',
    ])

    const loadRequestsCode = getRequiredCode(workflow, 'Load forecast requests')
    expect(loadRequestsCode).toContain("requestId: 'deal-1042-renewal'")
    expect(loadRequestsCode).toContain("account: 'Northstar Systems'")
    expect(loadRequestsCode).toContain("baseUrl = 'https://bilig.proompteng.ai'")

    const requestNode = getRequiredNode(workflow, 'Call Bilig WorkPaper')
    expect(readStringParameter(requestNode, 'method')).toBe('POST')
    expect(readStringParameter(requestNode, 'url')).toBe("={{ $json.baseUrl.replace(/\\/$/, '') + '/api/workpaper/n8n/forecast' }}")
    expect(readStringParameter(requestNode, 'jsonBody')).toBe('={{ JSON.stringify($json.biligPayload) }}')

    expect(getRequiredCode(workflow, 'Verify formula proof')).toContain('proof.verified !== true')
    expect(getRequiredCode(workflow, 'Apply approval policy')).toContain("decision = expectedArrOk && targetGapOk ? 'approve' : 'review'")
    expect(getRequiredCode(workflow, 'Write audit summary')).toContain("package: '@bilig/workpaper'")
  })

  it('keeps the community-node package metadata aligned with n8n install expectations', () => {
    const packageJson = readJsonRecord(join(n8nNodeDir, 'package.json'))
    const codexJson = readJsonRecord(join(n8nNodeDir, 'nodes', 'Workpaper', 'BiligWorkpaper.node.json'))
    const nodeSource = readFileSync(join(n8nNodeDir, 'nodes', 'Workpaper', 'BiligWorkpaper.node.ts'), 'utf8')
    const packageName = readRequiredString(packageJson['name'], 'package name')
    const keywords = readStringArray(packageJson['keywords'], 'package keywords')
    const n8n = readRecord(packageJson['n8n'], 'package n8n metadata')
    const peerDependencies = readRecord(packageJson['peerDependencies'], 'package peerDependencies')
    const devDependencies = readRecord(packageJson['devDependencies'], 'package devDependencies')
    const nodeCliVersion = readRequiredString(devDependencies['@n8n/node-cli'], '@n8n/node-cli version')

    expect(packageName).toBe('@bilig/n8n-nodes-workpaper')
    expect(packageJson['license']).toBe('MIT')
    expect(packageJson['dependencies']).toBeUndefined()
    expect(nodeCliVersion).not.toBe('*')
    expect(isAtLeastNodeCli023(nodeCliVersion)).toBe(true)
    expect(keywords).toContain('n8n-community-node-package')
    expect(n8n['strict']).toBe(true)
    expect(n8n['credentials']).toEqual([])
    expect(n8n['nodes']).toEqual(['dist/nodes/Workpaper/BiligWorkpaper.node.js'])
    expect(nodeSource).toContain("name: 'biligWorkpaper'")
    expect(codexJson['node']).toBe(`${packageName}.biligWorkpaper`)
    expect(peerDependencies['n8n-workflow']).toBe('*')
  })

  it('matches the workflow compact output to the live forecast proof contract', () => {
    const proof = createN8nForecastProof({
      sheetName: 'Inputs',
      address: 'B3',
      value: 0.4,
    })

    expect({
      verdict: 'verified',
      editedCell: proof.editedCell,
      beforeExpectedArr: proof.before.expectedArr,
      afterExpectedArr: proof.after.expectedArr,
      targetGap: proof.after.targetGap,
      checks: {
        formulasPersisted: proof.checks.formulasPersisted,
        restoredMatchesAfter: proof.checks.restoredMatchesAfter,
        computedOutputChanged: proof.checks.computedOutputChanged,
      },
    }).toEqual({
      verdict: 'verified',
      editedCell: 'Inputs!B3',
      beforeExpectedArr: 60000,
      afterExpectedArr: 96000,
      targetGap: 5600,
      checks: {
        formulasPersisted: true,
        restoredMatchesAfter: true,
        computedOutputChanged: true,
      },
    })
  })

  it('documents the importable workflow and expected compact proof shape', () => {
    const readme = readFileSync(join(workflowDir, 'README.md'), 'utf8')
    const nodeReadme = readFileSync(join(n8nNodeDir, 'README.md'), 'utf8')

    for (const required of [
      'bilig-workpaper-native-node.n8n.json',
      'bilig-workpaper-forecast-approval-guard.n8n.json',
      'bilig-workpaper-formula-readback.n8n.json',
      'bilig-workpaper-formula-readback.self-hosted.n8n.json',
      '@bilig/n8n-nodes-workpaper',
      'POST https://bilig.proompteng.ai/api/workpaper/n8n/forecast',
      '"verdict": "verified"',
      '"beforeExpectedArr": 60000',
      '"afterExpectedArr": 96000',
      '"targetGap": 5600',
      '"formulasPersisted": true',
      '"restoredMatchesAfter": true',
      '"computedOutputChanged": true',
    ]) {
      expect(readme).toContain(required)
    }
    expect(nodeReadme).toContain('bilig-workpaper-native-node.n8n.json')

    for (const forbidden of ['verified on\nn8n Cloud', 'This is intentionally not a custom n8n node yet']) {
      expect(readme).not.toContain(forbidden)
    }
  })
})

interface Workflow {
  readonly nodes: readonly WorkflowNode[]
  readonly connections: Record<string, unknown>
}

interface WorkflowNode {
  readonly name: string
  readonly type: string
  readonly parameters: Record<string, unknown>
}

function readWorkflow(fileName: string): Workflow {
  const parsed: unknown = JSON.parse(readFileSync(join(workflowDir, fileName), 'utf8'))
  if (!isObject(parsed)) {
    throw new Error(`${fileName} must contain a workflow object`)
  }

  const nodes = parsed['nodes']
  if (!Array.isArray(nodes)) {
    throw new Error(`${fileName} must contain a nodes array`)
  }

  return {
    nodes: nodes.map((node, index) => readWorkflowNode(node, `${fileName} node ${index}`)),
    connections: readRecord(parsed['connections'], `${fileName} connections`),
  }
}

function readWorkflowNode(value: unknown, context: string): WorkflowNode {
  if (!isObject(value)) {
    throw new Error(`${context} must be an object`)
  }
  const name = value['name']
  const type = value['type']
  if (typeof name !== 'string' || typeof type !== 'string') {
    throw new Error(`${context} must define string name and type`)
  }
  return {
    name,
    type,
    parameters: readRecord(value['parameters'], `${context} parameters`),
  }
}

function getRequiredNode(workflow: Workflow, name: string): WorkflowNode {
  const node = workflow.nodes.find((candidate) => candidate.name === name)
  if (!node) {
    throw new Error(`Workflow is missing node ${name}`)
  }
  return node
}

function getNodesByType(workflow: Workflow, type: string): WorkflowNode[] {
  return workflow.nodes.filter((node) => node.type === type)
}

function getRequiredCode(workflow: Workflow, name: string): string {
  const node = getRequiredNode(workflow, name)
  expect(node.type).toBe('n8n-nodes-base.code')
  return readStringParameter(node, 'jsCode')
}

function readStringParameter(node: WorkflowNode, key: string): string {
  const value = node.parameters[key]
  if (typeof value !== 'string') {
    throw new Error(`${node.name} parameter ${key} must be a string`)
  }
  return value
}

function readNumberParameter(node: WorkflowNode, key: string): number {
  const value = node.parameters[key]
  if (typeof value !== 'number') {
    throw new Error(`${node.name} parameter ${key} must be a number`)
  }
  return value
}

function readBooleanParameter(node: WorkflowNode, key: string): boolean {
  const value = node.parameters[key]
  if (typeof value !== 'boolean') {
    throw new Error(`${node.name} parameter ${key} must be a boolean`)
  }
  return value
}

function expectVerificationCodeReturnsCompactProof(code: string): void {
  for (const required of [
    'proof.verified !== true',
    'proof.checks?.formulasPersisted',
    'proof.checks?.restoredMatchesAfter',
    'proof.checks?.computedOutputChanged',
    "verdict: 'verified'",
    'editedCell: proof.editedCell',
    'beforeExpectedArr: proof.before.expectedArr',
    'afterExpectedArr: proof.after.expectedArr',
    'targetGap: proof.after.targetGap',
    'checks: proof.checks',
  ]) {
    expect(code).toContain(required)
  }
}

function readRecord(value: unknown, context: string): Record<string, unknown> {
  if (!isObject(value)) {
    throw new Error(`${context} must be an object`)
  }
  return value
}

function readJsonRecord(filePath: string): Record<string, unknown> {
  return readRecord(JSON.parse(readFileSync(filePath, 'utf8')), filePath)
}

function readRequiredString(value: unknown, context: string): string {
  if (typeof value !== 'string') {
    throw new Error(`${context} must be a string`)
  }
  return value
}

function isAtLeastNodeCli023(value: string): boolean {
  const match = value.match(/^[~^]?(\d+)\.(\d+)\.(\d+)$/)
  if (!match) {
    return false
  }
  const major = Number(match[1])
  const minor = Number(match[2])
  return major > 0 || (major === 0 && minor >= 23)
}

function readStringArray(value: unknown, context: string): string[] {
  if (!Array.isArray(value) || value.some((item) => typeof item !== 'string')) {
    throw new Error(`${context} must be a string array`)
  }
  return value
}

function isObject(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
}
