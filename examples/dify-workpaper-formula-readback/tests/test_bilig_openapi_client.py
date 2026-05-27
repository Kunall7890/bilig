import unittest

from tools.bilig_openapi_client import compact_proof, resolve_set_cell_endpoint


class BiligOpenApiClientTest(unittest.TestCase):
    def test_resolves_hosted_tool_base(self) -> None:
        self.assertEqual(
            resolve_set_cell_endpoint("https://bilig.proompteng.ai/openapi/workpaper"),
            "https://bilig.proompteng.ai/openapi/workpaper/set-cell-and-readback",
        )

    def test_resolves_app_root(self) -> None:
        self.assertEqual(
            resolve_set_cell_endpoint("https://bilig.proompteng.ai"),
            "https://bilig.proompteng.ai/openapi/workpaper/set-cell-and-readback",
        )

    def test_compacts_hosted_readback_proof(self) -> None:
        proof = {
            "editedCell": "Inputs!B3",
            "readbackRange": "Summary!A1:B3",
            "before": {"serialized": 0.25, "displayValue": "0.25"},
            "after": {"serialized": 0.4, "displayValue": "0.4"},
            "beforeReadback": {
                "values": [
                    [{"value": "Metric"}, {"value": "Value"}],
                    [{"value": "Expected customers"}, {"value": 5}],
                    [{"value": "Expected ARR"}, {"value": 60000}],
                ]
            },
            "afterReadback": {
                "values": [
                    [{"value": "Metric"}, {"value": "Value"}],
                    [{"value": "Expected customers"}, {"value": 8}],
                    [{"value": "Expected ARR"}, {"value": 96000}],
                ]
            },
            "persistence": {"persisted": False, "serializedBytes": 1162},
            "checks": {
                "persisted": False,
                "readbackChanged": True,
                "restoredReadbackMatchesAfter": True,
                "previousSerialized": 0.25,
                "newSerialized": 0.4,
            },
        }

        self.assertEqual(
            compact_proof(proof),
            {
                "verified": True,
                "editedCell": "Inputs!B3",
                "readbackRange": "Summary!A1:B3",
                "before": {
                    "input": 0.25,
                    "inputDisplay": "0.25",
                    "expectedCustomers": 5,
                    "expectedArr": 60000,
                },
                "after": {
                    "input": 0.4,
                    "inputDisplay": "0.4",
                    "expectedCustomers": 8,
                    "expectedArr": 96000,
                },
                "checks": {
                    "readbackChanged": True,
                    "restoredReadbackMatchesAfter": True,
                    "persisted": False,
                    "previousSerialized": 0.25,
                    "newSerialized": 0.4,
                },
                "persistence": {"serializedBytes": 1162, "requestLocal": True},
                "source": "Bilig WorkPaper OpenAPI",
                "github": "https://github.com/proompteng/bilig",
            },
        )


if __name__ == "__main__":
    unittest.main()
