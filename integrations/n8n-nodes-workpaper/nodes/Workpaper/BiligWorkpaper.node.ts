import { NodeConnectionTypes, type INodeType, type INodeTypeDescription } from 'n8n-workflow'

const forecastResource = {
  resource: ['forecast'],
}

const workpaperResource = {
  resource: ['workpaper'],
}

const verifyForecastReadback = {
  resource: ['forecast'],
  operation: ['verifyReadback'],
}

const evaluateWorkpaperDocument = {
  resource: ['workpaper'],
  operation: ['evaluateDocument'],
}

export class BiligWorkpaper implements INodeType {
  description: INodeTypeDescription = {
    displayName: 'Bilig WorkPaper',
    name: 'biligWorkpaper',
    icon: { light: 'file:workpaper.svg', dark: 'file:workpaper.dark.svg' },
    group: ['transform'],
    version: 1,
    subtitle: '={{$parameter["operation"]}}',
    description: 'Verify spreadsheet formula readback with Bilig WorkPaper',
    defaults: {
      name: 'Bilig WorkPaper',
    },
    usableAsTool: true,
    inputs: [NodeConnectionTypes.Main],
    outputs: [NodeConnectionTypes.Main],
    credentials: [],
    requestDefaults: {
      baseURL: '={{$parameter["baseUrl"].replace(/\\/$/, "")}}',
      headers: {
        Accept: 'application/json',
        'Content-Type': 'application/json',
      },
    },
    properties: [
      {
        displayName: 'Resource',
        name: 'resource',
        type: 'options',
        noDataExpression: true,
        options: [
          {
            name: 'Forecast',
            value: 'forecast',
          },
          {
            name: 'WorkPaper JSON',
            value: 'workpaper',
          },
        ],
        default: 'forecast',
      },
      {
        displayName: 'Operation',
        name: 'operation',
        type: 'options',
        noDataExpression: true,
        displayOptions: {
          show: forecastResource,
        },
        options: [
          {
            name: 'Verify Formula Readback',
            value: 'verifyReadback',
            action: 'Verify formula readback',
            description: 'Edit one forecast input cell and return recalculated formula proof',
            routing: {
              request: {
                method: 'POST',
                url: '/api/workpaper/n8n/forecast',
              },
            },
          },
        ],
        default: 'verifyReadback',
      },
      {
        displayName: 'Operation',
        name: 'operation',
        type: 'options',
        noDataExpression: true,
        displayOptions: {
          show: workpaperResource,
        },
        options: [
          {
            name: 'Evaluate Document',
            value: 'evaluateDocument',
            action: 'Evaluate workpaper document',
            description: 'Apply cell edits to a WorkPaper JSON document and return formula readback proof',
            routing: {
              request: {
                method: 'POST',
                url: '/api/workpaper/n8n/evaluate',
              },
            },
          },
        ],
        default: 'evaluateDocument',
      },
      {
        displayName: 'Bilig Base URL',
        name: 'baseUrl',
        type: 'string',
        default: 'https://bilig.proompteng.ai',
        required: true,
        description: 'Base URL for the Bilig app or hosted demo endpoint',
      },
      {
        displayName: 'Sheet Name',
        name: 'sheetName',
        type: 'string',
        default: 'Inputs',
        required: true,
        displayOptions: {
          show: verifyForecastReadback,
        },
        description: 'Forecast input sheet to edit',
        routing: {
          send: {
            type: 'body',
            property: 'sheetName',
          },
        },
      },
      {
        displayName: 'Cell',
        name: 'address',
        type: 'options',
        default: 'B3',
        required: true,
        displayOptions: {
          show: verifyForecastReadback,
        },
        options: [
          {
            name: 'B2 Qualified Opportunities',
            value: 'B2',
          },
          {
            name: 'B3 Win Rate',
            value: 'B3',
          },
          {
            name: 'B4 Average ARR',
            value: 'B4',
          },
          {
            name: 'B5 Expansion Multiplier',
            value: 'B5',
          },
        ],
        description: 'Editable forecast input cell',
        routing: {
          send: {
            type: 'body',
            property: 'address',
          },
        },
      },
      {
        displayName: 'Value',
        name: 'value',
        type: 'number',
        default: 0.4,
        required: true,
        displayOptions: {
          show: verifyForecastReadback,
        },
        description: 'Numeric value to write before formula readback',
        routing: {
          send: {
            type: 'body',
            property: 'value',
          },
        },
      },
      {
        displayName: 'Document JSON',
        name: 'document',
        type: 'json',
        default:
          '{"format":"bilig.headless.work-paper.document.v1","sheets":[{"name":"Inputs","content":[["Metric","Value"],["Win rate",0.25]]},{"name":"Summary","content":[["Metric","Value"],["Expected customers","=Inputs!B2*20"]]}],"namedExpressions":[]}',
        required: true,
        displayOptions: {
          show: evaluateWorkpaperDocument,
        },
        description: 'Bilig WorkPaper JSON document to evaluate',
        routing: {
          send: {
            type: 'body',
            property: 'document',
          },
        },
      },
      {
        displayName: 'Edits JSON',
        name: 'edits',
        type: 'json',
        default: '[{"cell":"Inputs!B2","value":0.4}]',
        required: true,
        displayOptions: {
          show: evaluateWorkpaperDocument,
        },
        description: 'Cell edits to apply before readback, for example [{"cell":"Inputs!B2","value":0.4}]',
        routing: {
          send: {
            type: 'body',
            property: 'edits',
          },
        },
      },
      {
        displayName: 'Read Cells',
        name: 'readCells',
        type: 'string',
        default: 'Summary!B2',
        required: true,
        displayOptions: {
          show: evaluateWorkpaperDocument,
        },
        description: 'Comma-separated cells to read before, after, and after JSON restore',
        routing: {
          send: {
            type: 'body',
            property: 'readCells',
          },
        },
      },
      {
        displayName: 'Include Updated Document',
        name: 'includeUpdatedDocument',
        type: 'boolean',
        default: true,
        displayOptions: {
          show: evaluateWorkpaperDocument,
        },
        description: 'Whether to include the updated WorkPaper JSON document in the output',
        routing: {
          send: {
            type: 'body',
            property: 'includeUpdatedDocument',
          },
        },
      },
    ],
  }
}
