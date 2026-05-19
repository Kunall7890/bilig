import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'

export const TEST_FUZZ_PACKAGE_ROOT = dirname(dirname(fileURLToPath(import.meta.url)))
export const BYTE_FUZZ_TARGETS_DIR = join(TEST_FUZZ_PACKAGE_ROOT, 'byte-targets')
export const BYTE_FUZZ_DICTIONARY_PATH = join(TEST_FUZZ_PACKAGE_ROOT, 'dictionaries', 'workbook-byte.dict')
