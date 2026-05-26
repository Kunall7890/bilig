from __future__ import annotations

import argparse
import json
import re
import subprocess
from pathlib import Path
from typing import Any

from smolagents import Tool


DEFAULT_PACKAGE_SPEC = "@bilig/workpaper@latest"
PACKAGE_SPEC_PATTERN = re.compile(r"^@bilig/workpaper(?:@[0-9A-Za-z._~+-]+)?$")


class BiligWorkPaperFormulaProofTool(Tool):
    name = "verify_workpaper_formula_readback"
    description = (
        "Run a Bilig WorkPaper formula proof. The tool edits an input cell, "
        "recalculates a dependent formula cell, serializes and restores the "
        "WorkPaper JSON document, and returns structured verification data."
    )
    inputs = {
        "package_spec": {
            "type": "string",
            "description": "The npm package spec to run. Use @bilig/workpaper@latest unless testing a pinned release.",
        },
    }
    output_type = "object"

    def forward(self, package_spec: str) -> dict[str, Any]:
        if PACKAGE_SPEC_PATTERN.fullmatch(package_spec) is None:
            raise ValueError("package_spec must be @bilig/workpaper or @bilig/workpaper@<version-or-tag>")

        completed = subprocess.run(
            [
                "npm",
                "exec",
                "--yes",
                "--package",
                package_spec,
                "--",
                "bilig-agent-challenge",
                "--json",
            ],
            check=False,
            capture_output=True,
            text=True,
            timeout=120,
        )
        if completed.returncode != 0:
            raise RuntimeError(
                "Bilig WorkPaper proof command failed:\n"
                f"stdout:\n{completed.stdout}\n\nstderr:\n{completed.stderr}",
            )

        proof = json.loads(completed.stdout)
        if proof.get("verified") is not True:
            raise RuntimeError(f"Bilig WorkPaper proof was not verified: {json.dumps(proof, indent=2)}")

        return {
            "framework": "smolagents",
            "toolName": self.name,
            "packageSpec": package_spec,
            "verified": True,
            "editedCell": proof.get("editedCell"),
            "dependentCell": proof.get("dependentCell"),
            "before": proof.get("before"),
            "after": proof.get("after"),
            "afterRestore": proof.get("afterRestore"),
            "persistedDocumentBytes": proof.get("persistedDocumentBytes"),
            "sheets": proof.get("sheets"),
            "checks": proof.get("checks"),
            "limitations": proof.get("limitations"),
            "star": proof.get("star"),
            "watchReleases": proof.get("watchReleases"),
            "adoptionBlocker": proof.get("adoptionBlocker"),
        }


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Run a smolagents Tool wrapper around Bilig WorkPaper proof.")
    parser.add_argument("--package", default=DEFAULT_PACKAGE_SPEC, help="Bilig npm package spec to execute.")
    parser.add_argument("--output", type=Path, help="Optional JSON proof output path.")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    tool = BiligWorkPaperFormulaProofTool()
    proof = tool(package_spec=args.package)
    proof_json = json.dumps(proof, indent=2, sort_keys=True)

    if args.output is not None:
        args.output.parent.mkdir(parents=True, exist_ok=True)
        args.output.write_text(f"{proof_json}\n", encoding="utf-8")

    print(proof_json)


if __name__ == "__main__":
    main()
