import { describe, expect, it } from 'vitest'
import { checkWorkbookFeaturePlugin, defineWorkbookFeaturePlugin, normalizeWorkbookCommandDescriptor } from '../index.js'

class TablesPlugin {
  readonly id = 'tables'
  readonly version = '1.0.0'
  readonly commands = []
  readonly projectionInterceptors = []
  readonly uiContributions = []
}

class TablesCommandDescriptor {
  readonly id = 'tables.createFromSelection'
  readonly featureId = 'tables'
  readonly category = 'command'
  readonly label = 'Create table'
}

class TablesProjectionInterceptor {
  readonly id = 'tables.rangeChrome'
  readonly featureId = 'tables'
  readonly point = 'rangeChrome'
}

class TablesUiContribution {
  readonly id = 'tables.toolbar.create'
  readonly featureId = 'tables'
  readonly slot = 'toolbar'
  readonly label = 'Create table'
}

describe('@bilig/workbook feature plugin data boundary', () => {
  it('rejects custom-prototype plugin manifests before registration', () => {
    const plugin = new TablesPlugin()

    expect(checkWorkbookFeaturePlugin(plugin)).toEqual({
      status: 'invalid',
      issues: [
        {
          code: 'invalid_feature_plugin',
          path: 'plugin',
          message: 'Workbook feature plugin must be an object record',
        },
      ],
    })
    expect(() => defineWorkbookFeaturePlugin(plugin)).toThrowError('Workbook feature plugin must be an object record')
  })

  it('rejects custom-prototype command, projection, and UI contribution records', () => {
    expect(
      checkWorkbookFeaturePlugin({
        id: 'tables',
        version: '1.0.0',
        commands: [new TablesCommandDescriptor()],
        projectionInterceptors: [new TablesProjectionInterceptor()],
        uiContributions: [new TablesUiContribution()],
      }),
    ).toEqual({
      status: 'invalid',
      issues: [
        {
          code: 'invalid_feature_plugin',
          path: 'commands[0]',
          message: 'Workbook command must be an object record',
        },
        {
          code: 'invalid_feature_plugin',
          path: 'projectionInterceptors[0]',
          message: 'Workbook projection interceptor must be an object record',
        },
        {
          code: 'invalid_feature_plugin',
          path: 'uiContributions[0]',
          message: 'Workbook UI contribution must be an object record',
        },
      ],
    })
  })

  it('normalizes command descriptors only from object-record data properties', () => {
    expect(() => normalizeWorkbookCommandDescriptor(new TablesCommandDescriptor())).toThrowError(
      'Workbook command descriptor must be an object record',
    )

    let getterInvoked = false
    const descriptor = {
      id: 'tables.createFromSelection',
      get featureId(): string {
        getterInvoked = true
        throw new Error('getter must not run')
      },
      category: 'command',
      label: 'Create table',
    }

    expect(() => normalizeWorkbookCommandDescriptor(descriptor)).toThrowError(
      'Workbook command descriptor feature id must be a data property',
    )
    expect(getterInvoked).toBe(false)
  })
})
