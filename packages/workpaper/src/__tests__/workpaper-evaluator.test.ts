import { describe, expect, it } from 'vitest'

import { listBiligEvaluatorDoors } from '../evaluator.js'

describe('@bilig/workpaper evaluator wrapper', () => {
  it('exposes only the WorkPaper evaluator doors', () => {
    expect(listBiligEvaluatorDoors().map((door) => door.door)).toEqual(['workpaper-service', 'agent-mcp'])
  })
})
