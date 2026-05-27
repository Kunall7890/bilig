from typing import Any
from urllib.error import HTTPError, URLError

from dify_plugin import ToolProvider

from tools.bilig_openapi_client import DEFAULT_OPENAPI_BASE_URL, call_bilig_set_cell_and_readback


class BiligProvider(ToolProvider):
    def _validate_credentials(self, credentials: dict[str, Any]) -> None:
        base_url = credentials.get("base_url") or DEFAULT_OPENAPI_BASE_URL
        try:
            call_bilig_set_cell_and_readback(
                base_url=base_url,
                address="B3",
                value=0.4,
                readback_range="Summary!A1:B3",
                timeout=10,
            )
        except (HTTPError, URLError, TimeoutError, ValueError) as error:
            raise ValueError(f"Bilig forecast readback validation failed: {error}") from error
