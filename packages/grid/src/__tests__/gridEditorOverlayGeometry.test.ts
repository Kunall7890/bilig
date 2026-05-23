// @vitest-environment jsdom
import { afterEach, describe, expect, test } from 'vitest'
import { resolveEditorOverlayScreenBounds, snapEditorOverlayScreenBounds } from '../gridEditorOverlayGeometry.js'

const ORIGINAL_DEVICE_PIXEL_RATIO = window.devicePixelRatio

function setDevicePixelRatio(value: number): void {
  Object.defineProperty(window, 'devicePixelRatio', {
    configurable: true,
    value,
  })
}

function expectDprAligned(value: number, dpr: number): void {
  expect(value * dpr).toBeCloseTo(Math.round(value * dpr), 6)
}

describe('gridEditorOverlayGeometry', () => {
  afterEach(() => {
    setDevicePixelRatio(ORIGINAL_DEVICE_PIXEL_RATIO)
  })

  test('snaps editor overlay bounds to physical pixels on fractional device ratios', () => {
    const bounds = snapEditorOverlayScreenBounds(
      {
        height: 22,
        width: 104,
        x: 46,
        y: 86.1875,
      },
      1.25,
    )

    expect(bounds.x).toBeCloseTo(46.4)
    expect(bounds.y).toBeCloseTo(86.4)
    expect(bounds.width).toBeCloseTo(104)
    expect(bounds.height).toBeCloseTo(21.6)
    expectDprAligned(bounds.x, 1.25)
    expectDprAligned(bounds.y, 1.25)
    expectDprAligned(bounds.x + bounds.width, 1.25)
    expectDprAligned(bounds.y + bounds.height, 1.25)
  })

  test('returns snapped viewport bounds for the active cell editor', () => {
    setDevicePixelRatio(1.25)
    const host = document.createElement('div')
    host.getBoundingClientRect = () =>
      ({
        bottom: 866.1875,
        height: 780,
        left: 0,
        right: 960,
        top: 86.1875,
        width: 960,
        x: 0,
        y: 86.1875,
        toJSON: () => ({}),
      }) as DOMRect

    const bounds = resolveEditorOverlayScreenBounds({
      col: 0,
      geometry: null,
      getCellLocalBounds: () => ({
        height: 22,
        width: 104,
        x: 46,
        y: 24,
      }),
      hostElement: host,
      row: 0,
    })

    expect(bounds?.x).toBeCloseTo(46.4)
    expect(bounds?.y).toBeCloseTo(110.4)
    expect(bounds?.width).toBeCloseTo(104)
    expect(bounds?.height).toBeCloseTo(21.6)
  })
})
