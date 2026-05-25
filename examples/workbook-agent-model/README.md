# Named-Range Workbook Agent Model

This example shows the intended `@bilig/workbook` shape without a built-in
business model. The consumer defines the model. Bilig supplies generic selectors,
formula helpers, static verification, runtime requirements, and a proof-shaped
run result.

```sh
npm install
npm start
```

The model binds three named refs, writes a formula to the result ref, and
declares checks. The adapter in this folder is a tiny proof fixture for
`runWorkbookPlan(..., { strict: true })`: it returns plan id, revision,
preview/apply op proof, command receipts with matching resolved refs, formula
readback, and check proof. Real runtimes such as `@bilig/core` provide those
facts from an engine; this example only shows the shape.

The output answers the questions an agent needs before it trusts a workbook
mutation:

- what model and action were selected
- which refs the selectors bound
- which commands and low-level ops were planned
- whether the plan survives JSON transport and `verifyPlanData`
- which adapter capabilities are required
- whether the transported plan data can be run without the consumer's private
  `refs` object shape
- whether apply matched preview with strict command proof
- what changed, and whether undo evidence exists
- which checks passed
- what proof supported each check

No revenue, quote, forecast, or other domain-specific model is built into
`@bilig/workbook`.
