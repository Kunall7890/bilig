import { Opcode } from '@bilig/protocol'

export const PUSH_CELL_OPCODE = Number(Opcode.PushCell)
export const PUSH_RANGE_OPCODE = Number(Opcode.PushRange)
export const PUSH_STRING_OPCODE = Number(Opcode.PushString)
export const EMPTY_RUNTIME_PROGRAM = new Uint32Array(0)
