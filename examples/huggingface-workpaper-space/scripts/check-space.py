from __future__ import annotations

import json
import subprocess
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]


def main() -> None:
    result = subprocess.run(
        ["python", "app.py", "--check"],
        cwd=ROOT,
        capture_output=True,
        check=True,
        text=True,
        timeout=120,
    )
    payload = json.loads(result.stdout)
    if payload.get("verified") is not True:
        raise SystemExit(f"expected verified=true, received: {payload}")
    if payload.get("packageVersion") != "0.164.0":
        raise SystemExit(f"unexpected package version: {payload.get('packageVersion')}")
    if payload.get("editedCell") != "Inputs!B3":
        raise SystemExit(f"unexpected edited cell: {payload.get('editedCell')}")
    if payload.get("after", {}).get("expectedArr") != 96000:
        raise SystemExit(f"unexpected expected ARR: {payload.get('after')}")


if __name__ == "__main__":
    main()
