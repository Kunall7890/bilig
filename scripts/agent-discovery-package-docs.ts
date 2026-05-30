interface WorkpaperAgentInstructionsInput {
  headlessPackageAgentInstructions: string
  headlessPackageSpec: string
  unscopedWorkpaperPackageSpec: string
  workpaperPackageSpec: string
}

interface WorkpaperSkillDocumentInput {
  skillDocument: string
  workpaperPackageSpec: string
  unscopedWorkpaperPackageSpec: string
}

export function buildWorkpaperPackageAgentInstructions(input: WorkpaperAgentInstructionsInput): string {
  return input.headlessPackageAgentInstructions
    .replace('# @bilig/headless agent notes', '# bilig-workpaper agent notes')
    .replace('agent inspecting `node_modules/@bilig/headless`', 'agent inspecting `node_modules/bilig-workpaper`')
    .replaceAll(input.headlessPackageSpec, input.unscopedWorkpaperPackageSpec)
    .replaceAll(input.workpaperPackageSpec, input.unscopedWorkpaperPackageSpec)
    .replace(/@bilig\/workpaper/g, 'bilig-workpaper')
    .replace(/@bilig\/headless/g, 'bilig-workpaper')
}

export function buildWorkpaperPackageSkillDocument(input: WorkpaperSkillDocumentInput): string {
  return input.skillDocument
    .replace(/@bilig\/workpaper/g, 'bilig-workpaper')
    .replaceAll(input.workpaperPackageSpec, input.unscopedWorkpaperPackageSpec)
    .replace(
      `"args": [
    "exec",
    "--package",
    "${input.unscopedWorkpaperPackageSpec}",
    "--",
    "bilig-formula-clinic",
    "./reduced.xlsx",
    "--cells",
    "Summary!B7,Inputs!B2"
  ]`,
      `"args": ["exec", "--package", "${input.unscopedWorkpaperPackageSpec}", "--", "bilig-formula-clinic", "./reduced.xlsx", "--cells", "Summary!B7,Inputs!B2"]`,
    )
    .replace(
      `## First Choice: MCP

Use MCP when the host can run a stdio server or call a Streamable HTTP server.
Configure stdio as an argument array, not a shell-concatenated string:

If the host supports installable skills, first check that the public skill
package is discoverable:

\`\`\`sh
npx --yes skills@latest add https://bilig.proompteng.ai --list
npx --yes skills@latest add proompteng/bilig --skill bilig-workpaper --list
\`\`\`

Before wiring a client, an agent can check the direct WorkPaper loop with:`,
      `## First Choice: Direct npm check or TypeScript

Use the package directly when the host can run npm or TypeScript. This is the
highest-traffic evaluator path because it meets developers where they already
search: npm, ExcelJS, SheetJS, xlsx-populate, and StackOverflow formula
recalculation problems.

If the host supports installable skills, first check that the public skill
package is discoverable:

\`\`\`sh
npx --yes skills@latest add https://bilig.proompteng.ai --list
npx --yes skills@latest add proompteng/bilig --skill bilig-workpaper --list
\`\`\`

Start by checking the direct WorkPaper loop:`,
    )
    .replace(
      'For the actual file-backed MCP path, run the package-owned challenge first:',
      'Use MCP only when the host specifically needs an MCP client boundary. For that path, run the package-owned challenge first:',
    )
    .replace('## Second Choice: Direct TypeScript', '## Direct TypeScript')
}
