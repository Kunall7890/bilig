from __future__ import annotations

import json
import subprocess
from pathlib import Path
from typing import Any

from prefect import flow, task

ROOT = Path(__file__).resolve().parent
PROOF_PATH = ROOT / ".tmp" / "workpaper-proof.json"


@task(retries=2, retry_delay_seconds=5)
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

    return result


@flow(name="bilig-workpaper-quote")
def bilig_workpaper_quote_flow(
    quantity: int = 18,
    previous_quantity: int = 12,
    unit_price: float = 125,
    discount_rate: float = 0.1,
    tax_rate: float = 0.08,
    unit_cost: float = 52,
) -> dict[str, Any]:
    return calculate_quote_fields(
        quantity=quantity,
        previous_quantity=previous_quantity,
        unit_price=unit_price,
        discount_rate=discount_rate,
        tax_rate=tax_rate,
        unit_cost=unit_cost,
    )


if __name__ == "__main__":
    print(json.dumps(bilig_workpaper_quote_flow(), indent=2))

