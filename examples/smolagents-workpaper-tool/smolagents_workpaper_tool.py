from __future__ import annotations

import argparse
import json
import re
import subprocess
import urllib.error
import urllib.request
from pathlib import Path
from typing import Any

from smolagents import Tool


DEFAULT_PACKAGE_SPEC = "@bilig/workpaper@latest"
DEFAULT_SPACE_API_URL = "https://gregkonush-bilig-workpaper-mcp-readback.hf.space/gradio_api/call/v2/prove_workpaper_readback"
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
                "bilig-evaluate",
                "--door",
                "agent-mcp",
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

        evidence = proof.get("evidence", {})
        return {
            "framework": "smolagents",
            "toolName": self.name,
            "packageSpec": package_spec,
            "door": proof.get("door"),
            "verified": True,
            "editedCell": evidence.get("editedCell"),
            "dependentCell": evidence.get("dependentCell"),
            "before": evidence.get("before"),
            "after": evidence.get("after"),
            "afterRestore": evidence.get("afterRestore"),
            "afterRestart": evidence.get("afterRestart"),
            "persistedDocumentBytes": evidence.get("persistedDocumentBytes"),
            "tools": evidence.get("tools"),
            "checks": evidence.get("checks"),
            "limitations": proof.get("limitations"),
        }


class BiligWorkPaperSpaceReadbackTool(Tool):
    name = "read_workpaper_space_formula"
    description = (
        "Call the public Bilig Hugging Face Space fixture. The Space edits "
        "Inputs!B3, recalculates formula cells, restores WorkPaper JSON, and "
        "returns verified readback data without any model key."
    )
    inputs = {
        "win_rate": {
            "type": "number",
            "description": "New win-rate input for Inputs!B3. Use 0.4 for the public smoke test.",
            "nullable": False,
        },
    }
    output_type = "object"

    def forward(self, win_rate: float = 0.4) -> dict[str, Any]:
        if win_rate <= 0 or win_rate > 1:
            raise ValueError("win_rate must be greater than 0 and no more than 1")

        proof = call_workpaper_space(win_rate)
        if proof.get("verified") is not True:
            raise RuntimeError(f"Bilig WorkPaper Space readback was not verified: {json.dumps(proof, indent=2)}")

        return {
            "framework": "smolagents",
            "toolName": self.name,
            "space": "gregkonush/bilig-workpaper-mcp-readback",
            "verified": True,
            "editedCell": proof.get("editedCell"),
            "before": proof.get("before"),
            "after": proof.get("after"),
            "afterRestore": proof.get("afterRestore"),
            "persistedDocumentBytes": proof.get("persistedDocumentBytes"),
            "checks": proof.get("checks"),
            "limitations": proof.get("limitations"),
        }


def call_workpaper_space(win_rate: float) -> dict[str, Any]:
    payload = json.dumps({"win_rate": win_rate}).encode("utf-8")
    start_request = urllib.request.Request(
        DEFAULT_SPACE_API_URL,
        data=payload,
        headers={"content-type": "application/json"},
        method="POST",
    )

    try:
        with urllib.request.urlopen(start_request, timeout=30) as response:
            event_id = json.load(response).get("event_id")
    except urllib.error.URLError as exc:
        raise RuntimeError(f"Could not start Hugging Face Space readback: {exc}") from exc

    if not isinstance(event_id, str) or not event_id:
        raise RuntimeError("Hugging Face Space response did not include an event_id")

    event_url = f"{DEFAULT_SPACE_API_URL}/{event_id}"
    try:
        with urllib.request.urlopen(event_url, timeout=60) as response:
            event_stream = response.read().decode("utf-8")
    except urllib.error.URLError as exc:
        raise RuntimeError(f"Could not fetch Hugging Face Space readback event: {exc}") from exc

    for line in event_stream.splitlines():
        if line.startswith("data: "):
            data = json.loads(line.removeprefix("data: "))
            if isinstance(data, list) and data:
                proof = data[0]
                if isinstance(proof, dict):
                    return proof

    raise RuntimeError(f"Hugging Face Space event did not include JSON proof data:\n{event_stream}")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Run a smolagents Tool wrapper around Bilig WorkPaper proof.")
    parser.add_argument(
        "--mode",
        choices=("local", "space"),
        default="local",
        help="Run the local npm evaluator or call the public Hugging Face Space.",
    )
    parser.add_argument("--package", default=DEFAULT_PACKAGE_SPEC, help="Bilig npm package spec to execute.")
    parser.add_argument("--win-rate", type=float, default=0.4, help="Win rate input for the Hugging Face Space mode.")
    parser.add_argument("--output", type=Path, help="Optional JSON proof output path.")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    if args.mode == "space":
        tool = BiligWorkPaperSpaceReadbackTool()
        proof = tool(win_rate=args.win_rate)
    else:
        tool = BiligWorkPaperFormulaProofTool()
        proof = tool(package_spec=args.package)

    proof_json = json.dumps(proof, indent=2, sort_keys=True)

    if args.output is not None:
        args.output.parent.mkdir(parents=True, exist_ok=True)
        args.output.write_text(f"{proof_json}\n", encoding="utf-8")

    print(proof_json)


if __name__ == "__main__":
    main()
