import type { PivotAggregation } from '@bilig/protocol'

export function subtotalValue(value: PivotAggregation): string {
  switch (value) {
    case 'sum':
      return 'sum'
    case 'count':
      return 'count'
    case 'countNums':
      return 'countNums'
    case 'average':
      return 'average'
    case 'min':
      return 'min'
    case 'max':
      return 'max'
    case 'product':
      return 'product'
  }
}

export function defaultDataFieldVerb(value: PivotAggregation): string {
  switch (value) {
    case 'sum':
      return 'Sum'
    case 'count':
      return 'Count'
    case 'countNums':
      return 'Count Nums'
    case 'average':
      return 'Average'
    case 'min':
      return 'Min'
    case 'max':
      return 'Max'
    case 'product':
      return 'Product'
  }
}
