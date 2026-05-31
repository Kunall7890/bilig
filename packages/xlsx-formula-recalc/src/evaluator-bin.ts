import { runBiligEvaluatorCli } from './evaluator-cli.js'

process.exitCode = runBiligEvaluatorCli(process.argv.slice(2))
