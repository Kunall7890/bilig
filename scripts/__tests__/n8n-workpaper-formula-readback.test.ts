import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import { describe, expect, it } from 'vitest'

import { createN8nForecastProof } from '@bilig/headless'

const workflowDir = join(process.cwd(), 'examples', 'n8n-workpaper-formula-readback')

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

    for (const required of [
      'bilig-workpaper-formula-readback.n8n.json',
      'bilig-workpaper-formula-readback.self-hosted.n8n.json',
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

function isObject(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
}
