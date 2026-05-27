import json
from typing import Any
from urllib.parse import urljoin, urlparse
from urllib.request import Request, urlopen


DEFAULT_OPENAPI_BASE_URL = "https://bilig.proompteng.ai/openapi/workpaper"


def resolve_set_cell_endpoint(base_url: str) -> str:
    if not base_url.startswith(("http://", "https://")):
        raise ValueError("base_url must start with http:// or https://")

    parsed = urlparse(base_url)
    path = parsed.path.rstrip("/")
    if path.endswith("/set-cell-and-readback"):
        return base_url
    if path.endswith("/openapi/workpaper"):
        return urljoin(base_url.rstrip("/") + "/", "set-cell-and-readback")
    return urljoin(base_url.rstrip("/") + "/", "openapi/workpaper/set-cell-and-readback")


def call_bilig_set_cell_and_readback(
    *,
    base_url: str,
    address: str,
    value: Any,
    sheet_name: str = "Inputs",
    readback_range: str = "Summary!A1:B3",
    timeout: int = 30,
) -> dict[str, Any]:
    endpoint = resolve_set_cell_endpoint(base_url)
    payload = json.dumps(
        {
            "sheetName": sheet_name,
            "address": address,
            "value": value,
            "readbackRange": readback_range,
        }
    ).encode("utf-8")
    request = Request(
        endpoint,
        data=payload,
        headers={
            "content-type": "application/json",
            "accept": "application/json",
            "user-agent": "Bilig-Dify-WorkPaper-Plugin/0.1 (+https://github.com/proompteng/bilig)",
        },
        method="POST",
    )

    with urlopen(request, timeout=timeout) as response:
        body = response.read().decode("utf-8")

    parsed = json.loads(body)
    if not isinstance(parsed, dict):
        raise ValueError(f"Expected JSON object response, received {type(parsed).__name__}")
    checks = parsed.get("checks") if isinstance(parsed.get("checks"), dict) else {}
    if checks.get("readbackChanged") is not True:
        raise ValueError(f"Unchanged WorkPaper readback response: {parsed}")
    if checks.get("restoredReadbackMatchesAfter") is not True:
        raise ValueError(f"Unrestorable WorkPaper readback response: {parsed}")
    return parsed


def compact_proof(proof: dict[str, Any]) -> dict[str, Any]:
    checks = proof.get("checks") if isinstance(proof.get("checks"), dict) else {}
    before = proof.get("before") if isinstance(proof.get("before"), dict) else {}
    after = proof.get("after") if isinstance(proof.get("after"), dict) else {}
    before_readback = proof.get("beforeReadback") if isinstance(proof.get("beforeReadback"), dict) else {}
    after_readback = proof.get("afterReadback") if isinstance(proof.get("afterReadback"), dict) else {}
    persistence = proof.get("persistence") if isinstance(proof.get("persistence"), dict) else {}

    return {
        "verified": checks.get("readbackChanged") is True
        and checks.get("restoredReadbackMatchesAfter") is True,
        "editedCell": proof.get("editedCell"),
        "readbackRange": proof.get("readbackRange"),
        "before": {
            "input": before.get("serialized"),
            "inputDisplay": before.get("displayValue"),
            "expectedCustomers": _read_numeric_cell(before_readback, row=1, column=1),
            "expectedArr": _read_numeric_cell(before_readback, row=2, column=1),
        },
        "after": {
            "input": after.get("serialized"),
            "inputDisplay": after.get("displayValue"),
            "expectedCustomers": _read_numeric_cell(after_readback, row=1, column=1),
            "expectedArr": _read_numeric_cell(after_readback, row=2, column=1),
        },
        "checks": {
            "readbackChanged": checks.get("readbackChanged") is True,
            "restoredReadbackMatchesAfter": checks.get("restoredReadbackMatchesAfter") is True,
            "persisted": checks.get("persisted") is True,
            "previousSerialized": checks.get("previousSerialized"),
            "newSerialized": checks.get("newSerialized"),
        },
        "persistence": {
            "serializedBytes": persistence.get("serializedBytes"),
            "requestLocal": persistence.get("persisted") is not True,
        },
        "source": "Bilig WorkPaper OpenAPI",
        "github": "https://github.com/proompteng/bilig",
    }


def _read_numeric_cell(readback: dict[str, Any], *, row: int, column: int) -> Any:
    values = readback.get("values")
    if not isinstance(values, list) or row >= len(values):
        return None
    row_values = values[row]
    if not isinstance(row_values, list) or column >= len(row_values):
        return None
    cell = row_values[column]
    if isinstance(cell, dict):
        return cell.get("value")
    return cell
