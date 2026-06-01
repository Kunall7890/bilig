from __future__ import annotations

import json
import os
import subprocess
import sys
from pathlib import Path
from typing import Any

import gradio as gr


ROOT = Path(__file__).resolve().parent


def prove_workpaper_readback(win_rate: float = 0.4) -> dict[str, Any]:
    rate = float(win_rate)
    if rate < 0 or rate > 1:
        raise gr.Error("Win rate must be between 0 and 1.")

    result = subprocess.run(
        ["node", str(ROOT / "workpaper_proof.mjs"), str(rate)],
        cwd=ROOT,
        capture_output=True,
        check=False,
        text=True,
        timeout=90,
    )
    if result.returncode != 0:
        message = result.stderr.strip() or result.stdout.strip() or "WorkPaper readback failed."
        raise gr.Error(message)

    try:
        payload = json.loads(result.stdout)
    except json.JSONDecodeError as error:
        raise gr.Error(f"WorkPaper readback returned invalid JSON: {error}") from error

    if payload.get("verified") is not True:
        raise gr.Error("WorkPaper readback did not verify.")

    return payload


if "--check" in sys.argv:
    print(json.dumps(prove_workpaper_readback(0.4), indent=2, sort_keys=True))
    raise SystemExit(0)


demo = gr.Interface(
    fn=prove_workpaper_readback,
    inputs=gr.Number(label="Win rate", value=0.4, minimum=0, maximum=1),
    outputs=gr.JSON(label="WorkPaper readback"),
    title="Bilig WorkPaper readback",
    description=(
        "Edit one input cell, recalculate formulas, serialize JSON, restore it, "
        "and return the exact values Bilig read back."
    ),
)

demo.launch(
    server_name=os.getenv("GRADIO_SERVER_NAME", "0.0.0.0"),
    server_port=int(os.getenv("GRADIO_SERVER_PORT", "7860")),
    mcp_server=True,
)
