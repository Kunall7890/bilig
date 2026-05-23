#!/usr/bin/env node
import { runN8nForecastServerCli } from '@bilig/headless/cli'

process.exitCode = runN8nForecastServerCli({
  argv: process.argv.slice(2),
  env: process.env,
})
