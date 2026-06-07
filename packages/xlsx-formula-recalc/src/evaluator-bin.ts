import { runBiligEvaluatorCli } from './evaluator-cli.js'

process.exitCode = await runBiligEvaluatorCli(process.argv.slice(2))
