#!/usr/bin/env node
import { runBiligEvaluatorCli } from './evaluator.js'

process.exitCode = await runBiligEvaluatorCli(process.argv.slice(2))
