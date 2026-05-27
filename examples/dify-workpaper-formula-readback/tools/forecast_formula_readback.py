import json
from typing import Any, Generator

from dify_plugin import Tool
from dify_plugin.entities.tool import ToolInvokeMessage

from tools.bilig_openapi_client import (
    DEFAULT_OPENAPI_BASE_URL,
    call_bilig_set_cell_and_readback,
    compact_proof,
)


class ForecastFormulaReadbackTool(Tool):
    def _invoke(self, tool_parameters: dict[str, Any]) -> Generator[ToolInvokeMessage]:
        base_url = self.runtime.credentials.get("base_url") or DEFAULT_OPENAPI_BASE_URL
        address = str(tool_parameters.get("address") or "B3").upper()
        sheet_name = str(tool_parameters.get("sheet_name") or "Inputs")
        readback_range = str(tool_parameters.get("readback_range") or "Summary!A1:B3")
        value = tool_parameters.get("value", 0.4)

        try:
            proof = call_bilig_set_cell_and_readback(
                base_url=base_url,
                sheet_name=sheet_name,
                address=address,
                value=value,
                readback_range=readback_range,
            )
            yield self.create_json_message(json=compact_proof(proof))
        except Exception as error:
            yield self.create_text_message(text=f"Bilig WorkPaper forecast readback failed: {error}")
