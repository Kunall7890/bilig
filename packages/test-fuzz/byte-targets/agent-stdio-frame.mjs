import { decodeAgentFrame, decodeStdioMessages } from '@bilig/agent-api'

export function fuzz(data) {
  try {
    decodeAgentFrame(data)
  } catch (error) {
    if (!(error instanceof Error)) {
      throw error
    }
  }

  try {
    decodeStdioMessages(data)
  } catch (error) {
    if (!(error instanceof Error)) {
      throw error
    }
  }
}
