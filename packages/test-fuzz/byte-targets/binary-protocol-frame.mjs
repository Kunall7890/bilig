import { decodeFrame } from '@bilig/binary-protocol'

export function fuzz(data) {
  try {
    decodeFrame(data)
  } catch (error) {
    if (!(error instanceof Error)) {
      throw error
    }
  }
}
