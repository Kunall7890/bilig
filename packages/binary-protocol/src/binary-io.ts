const textEncoder = new TextEncoder()
const textDecoder = new TextDecoder()

export class BinaryProtocolError extends Error {
  constructor(message: string) {
    super(message)
    this.name = 'BinaryProtocolError'
  }
}

export class BinaryWriter {
  private readonly chunks: number[] = []

  u8(value: number): void {
    this.chunks.push(value & 0xff)
  }

  u32(value: number): void {
    const next = value >>> 0
    this.chunks.push(next & 0xff, (next >>> 8) & 0xff, (next >>> 16) & 0xff, (next >>> 24) & 0xff)
  }

  f64(value: number): void {
    const buffer = new ArrayBuffer(8)
    new DataView(buffer).setFloat64(0, value, true)
    this.bytes(new Uint8Array(buffer))
  }

  bool(value: boolean): void {
    this.u8(value ? 1 : 0)
  }

  string(value: string): void {
    this.bytes(textEncoder.encode(value))
  }

  stringArray(values: readonly string[]): void {
    this.u32(values.length)
    values.forEach((value) => this.string(value))
  }

  bytes(value: Uint8Array): void {
    this.u32(value.byteLength)
    value.forEach((byte) => this.chunks.push(byte))
  }

  finish(): Uint8Array {
    return Uint8Array.from(this.chunks)
  }
}

export class BinaryReader {
  private offset = 0

  constructor(private readonly bytes: Uint8Array) {}

  private ensure(size: number): void {
    if (this.offset + size > this.bytes.byteLength) {
      throw new BinaryProtocolError('Unexpected end of binary frame')
    }
  }

  u8(): number {
    this.ensure(1)
    const value = this.bytes[this.offset]!
    this.offset += 1
    return value
  }

  u32(): number {
    this.ensure(4)
    const view = new DataView(this.bytes.buffer, this.bytes.byteOffset + this.offset, 4)
    const value = view.getUint32(0, true)
    this.offset += 4
    return value
  }

  f64(): number {
    const buffer = this.bytesView()
    return new DataView(buffer.buffer, buffer.byteOffset, buffer.byteLength).getFloat64(0, true)
  }

  bool(): boolean {
    return this.u8() === 1
  }

  string(): string {
    return textDecoder.decode(this.bytesView())
  }

  stringArray(): string[] {
    const count = this.u32()
    const values: string[] = []
    for (let index = 0; index < count; index += 1) {
      values.push(this.string())
    }
    return values
  }

  bytesView(): Uint8Array {
    const length = this.u32()
    this.ensure(length)
    const slice = this.bytes.subarray(this.offset, this.offset + length)
    this.offset += length
    return slice
  }

  done(): boolean {
    return this.offset === this.bytes.byteLength
  }
}
