from __future__ import annotations

import json
import subprocess
from datetime import datetime
from pathlib import Path
from typing import Any

try:
    from airflow.sdk import dag, task
except ImportError:
    from airflow.decorators import dag, task

ROOT = Path(__file__).resolve().parents[1]
PROOF_PATH = ROOT / ".tmp" / "workpaper-proof.json"


@dag(
    dag_id="bilig_workpaper_quote",
    start_date=datetime(2026, 1, 1),
    schedule=None,
    catchup=False,
    tags=["bilig", "workpaper", "formula-readback"],
)
def bilig_workpaper_quote_dag() -> None:
    @task(retries=2)
    def calculate_quote_fields(
        quantity: int = 18,
        previous_quantity: int = 12,
        unit_price: float = 125,
        discount_rate: float = 0.1,
        tax_rate: float = 0.08,
        unit_cost: float = 52,
    ) -> dict[str, Any]:
        PROOF_PATH.parent.mkdir(parents=True, exist_ok=True)
        command = [
            "npx",
            "--no-install",
            "tsx",
            "workpaper-quote.ts",
            "--quantity",
            str(quantity),
            "--previous-quantity",
            str(previous_quantity),
            "--unit-price",
            str(unit_price),
            "--discount-rate",
            str(discount_rate),
            "--tax-rate",
            str(tax_rate),
            "--unit-cost",
            str(unit_cost),
            "--output",
            str(PROOF_PATH),
        ]

        completed = subprocess.run(command, cwd=ROOT, check=True, capture_output=True, text=True)
        result = json.loads(PROOF_PATH.read_text(encoding="utf-8"))

        if not result.get("proof", {}).get("verified"):
            raise RuntimeError(f"WorkPaper proof failed: {completed.stdout}")

        return {
            "patch": result["patch"],
            "proof": {
                "editedCell": result["proof"]["editedCell"],
                "beforeTotal": result["proof"]["before"]["total"],
                "afterTotal": result["proof"]["after"]["total"],
                "afterRestoreTotal": result["proof"]["afterRestore"]["total"],
                "persistedDocumentBytes": result["proof"]["persistedDocumentBytes"],
                "outputFile": result["proof"]["outputFile"],
                "verified": result["proof"]["verified"],
            },
        }

    @task
    def verify_formula_proof(result: dict[str, Any]) -> dict[str, Any]:
        proof = result["proof"]
        if (
            proof["editedCell"] != "Inputs!B2"
            or proof["beforeTotal"] != 1458
            or proof["afterTotal"] != 2187
            or proof["afterRestoreTotal"] != 2187
            or proof["persistedDocumentBytes"] <= 0
            or not proof["verified"]
        ):
            raise ValueError(f"Unexpected WorkPaper proof: {json.dumps(proof, sort_keys=True)}")

        return result

    verify_formula_proof(calculate_quote_fields())


bilig_workpaper_quote_dag()
