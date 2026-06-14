import io
import os
import unittest
from unittest.mock import patch

import pandas as pd
from fastapi.testclient import TestClient

from api.index import _build_preview_payload, app


def make_xlsx_bytes(sheets: dict[str, list[dict[str, object]]]) -> bytes:
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        for sheet_name, rows in sheets.items():
            pd.DataFrame(rows).to_excel(writer, sheet_name=sheet_name, index=False)
    return buffer.getvalue()


class BuildPreviewPayloadTests(unittest.TestCase):
    @patch("api.index.inject_variables", side_effect=lambda template, row: template.replace("{{name}}", str(row["name"])))
    @patch("api.index.convert_docx_to_html", return_value="<p>{{name}}</p>")
    def test_build_preview_payload_exposes_sheet_names_and_default_selection(self, _convert_docx_to_html, _inject_variables):
        xlsx_content = make_xlsx_bytes(
            {
                "名單一": [{"name": "Alice", "email": "alice@example.com", "subject": "Hello"}],
                "名單二": [{"name": "Bob", "email": "bob@example.com", "subject": "World"}],
            }
        )

        payload = _build_preview_payload(
            docx_content=b"unused",
            xlsx_content=xlsx_content,
            sheet=None,
            font=None,
            cache_info=None,
        )

        self.assertEqual(payload["sheet_names"], ["名單一", "名單二"])
        self.assertEqual(payload["selected_sheet"], "名單一")
        self.assertEqual(payload["first_row"], {"name": "Alice", "email": "alice@example.com", "subject": "Hello"})

    @patch("api.index.inject_variables", side_effect=lambda template, row: template.replace("{{name}}", str(row["name"])))
    @patch("api.index.convert_docx_to_html", return_value="<p>{{name}}</p>")
    def test_build_preview_payload_uses_selected_sheet(self, _convert_docx_to_html, _inject_variables):
        xlsx_content = make_xlsx_bytes(
            {
                "第一批": [{"name": "Alice", "email": "alice@example.com", "subject": "Hello"}],
                "第二批": [{"name": "Bob", "email": "bob@example.com", "subject": "World"}],
            }
        )

        payload = _build_preview_payload(
            docx_content=b"unused",
            xlsx_content=xlsx_content,
            sheet="第二批",
            font=None,
            cache_info=None,
        )

        self.assertEqual(payload["selected_sheet"], "第二批")
        self.assertEqual(payload["first_row"], {"name": "Bob", "email": "bob@example.com", "subject": "World"})


class DriveFilesEndpointTests(unittest.TestCase):
    def setUp(self) -> None:
        self.client = TestClient(app)
        self.client.post("/api/dev/login")

    @patch("api.index.load_user_credentials", return_value=object())
    @patch("api.index.build")
    def test_drive_files_lists_supported_recipient_files(self, mock_build, _load_user_credentials):
        class FakeListRequest:
            def execute(self):
                return {
                    "files": [
                        {
                            "id": "xlsx-file",
                            "name": "名單.xlsx",
                            "mimeType": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            "modifiedTime": "2026-06-14T10:00:00Z",
                        },
                        {
                            "id": "sheet-file",
                            "name": "Sheet 名單",
                            "mimeType": "application/vnd.google-apps.spreadsheet",
                            "modifiedTime": "2026-06-13T10:00:00Z",
                        },
                        {
                            "id": "doc-file",
                            "name": "模板.docx",
                            "mimeType": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            "modifiedTime": "2026-06-12T10:00:00Z",
                        },
                    ]
                }

        class FakeFilesResource:
            def list(self, **_kwargs):
                return FakeListRequest()

        class FakeDriveService:
            def files(self):
                return FakeFilesResource()

        mock_build.return_value = FakeDriveService()

        response = self.client.get("/api/drive/files", params={"kind": "xlsx"})

        self.assertEqual(response.status_code, 200)
        payload = response.json()
        self.assertEqual(
            [item["name"] for item in payload["files"]],
            ["名單.xlsx", "Sheet 名單"],
        )
        self.assertTrue(all(item["kind"] == "xlsx" for item in payload["files"]))


class PickerConfigEndpointTests(unittest.TestCase):
    def setUp(self) -> None:
        self.client = TestClient(app)

    def test_picker_config_returns_enabled_payload_when_configured(self):
        with patch.dict(
            os.environ,
            {
                "GOOGLE_CLIENT_ID": "708359400952-example.apps.googleusercontent.com",
                "GOOGLE_PICKER_API_KEY": "picker-key",
            },
            clear=False,
        ):
            response = self.client.get("/api/google/picker/config")

        self.assertEqual(response.status_code, 200)
        payload = response.json()
        self.assertTrue(payload["enabled"])
        self.assertEqual(payload["client_id"], "708359400952-example.apps.googleusercontent.com")
        self.assertEqual(payload["api_key"], "picker-key")
        self.assertEqual(payload["app_id"], "708359400952")
        self.assertIsNone(payload["login_hint"])

    def test_picker_config_returns_disabled_when_api_key_missing(self):
        with patch.dict(
            os.environ,
            {
                "GOOGLE_CLIENT_ID": "708359400952-example.apps.googleusercontent.com",
                "GOOGLE_PICKER_API_KEY": "",
            },
            clear=False,
        ):
            response = self.client.get("/api/google/picker/config")

        self.assertEqual(response.status_code, 200)
        payload = response.json()
        self.assertFalse(payload["enabled"])


if __name__ == "__main__":
    unittest.main()
