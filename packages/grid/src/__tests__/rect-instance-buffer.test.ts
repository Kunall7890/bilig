import { describe, expect, it } from 'vitest'
import { coalesceGridRectsV3, packGridRectBufferV3 } from '../renderer-v3/rect-instance-buffer.js'
import type { GridGpuColor, GridGpuRect } from '../gridGpuScene.js'

const black: GridGpuColor = Object.freeze({ a: 1, b: 0, g: 0, r: 0 })
const red: GridGpuColor = Object.freeze({ a: 1, b: 0, g: 0, r: 1 })

describe('rect-instance-buffer', () => {
  it('coalesces adjacent same-color grid-line rects before GPU packing', () => {
    const rects: GridGpuRect[] = [
      { color: black, height: 1, width: 10, x: 0, y: 4 },
      { color: black, height: 1, width: 10, x: 10, y: 4 },
      { color: black, height: 1, width: 10, x: 20, y: 4 },
      { color: black, height: 5, width: 1, x: 2, y: 0 },
      { color: black, height: 5, width: 1, x: 2, y: 5 },
    ]

    const coalesced = coalesceGridRectsV3(rects)

    expect(coalesced).toEqual([
      { color: black, height: 1, width: 30, x: 0, y: 4 },
      { color: black, height: 10, width: 1, x: 2, y: 0 },
    ])
  })

  it('keeps different colors as separate draw rects', () => {
    const coalesced = coalesceGridRectsV3([
      { color: black, height: 1, width: 10, x: 0, y: 4 },
      { color: red, height: 1, width: 10, x: 10, y: 4 },
    ])

    expect(coalesced).toHaveLength(2)
  })

  it('packs coalesced rects into the instance buffer', () => {
    const packed = packGridRectBufferV3(
      {
        borderRects: [
          { color: black, height: 1, width: 10, x: 0, y: 4 },
          { color: black, height: 1, width: 10, x: 10, y: 4 },
        ],
        fillRects: [],
      },
      { height: 20, width: 20 },
    )

    expect(packed.rectCount).toBe(1)
    expect(packed.borderRectCount).toBe(1)
    expect(packed.rectInstances[0]).toBe(0)
    expect(packed.rectInstances[2]).toBe(20)
  })
})
