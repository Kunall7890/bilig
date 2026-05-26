from __future__ import annotations

from pathlib import Path

import dagster as dg

ROOT = Path(__file__).resolve().parents[1]
PROOF_PATH = ROOT / ".tmp" / "workpaper-proof.json"


@dg.asset(compute_kind="javascript")
def bilig_workpaper_quote_asset(
    context: dg.AssetExecutionContext,
    pipes_subprocess_client: dg.PipesSubprocessClient,
) -> dg.MaterializeResult:
    return pipes_subprocess_client.run(
        command=[
            "npx",
            "--no-install",
            "tsx",
            "workpaper-asset.ts",
            "--output",
            str(PROOF_PATH),
        ],
        context=context,
        extras={
            "quantity": 18,
            "previous_quantity": 12,
            "unit_price": 125,
            "discount_rate": 0.1,
            "tax_rate": 0.08,
            "unit_cost": 52,
        },
    ).get_materialize_result()


defs = dg.Definitions(
    assets=[bilig_workpaper_quote_asset],
    resources={"pipes_subprocess_client": dg.PipesSubprocessClient()},
)
