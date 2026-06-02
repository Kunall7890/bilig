#!/usr/bin/env node
import { runAgentStartCli } from '@bilig/headless/cli'

process.exitCode = runAgentStartCli({
  argv: process.argv.slice(2),
})
