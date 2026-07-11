import io
import os
import unittest
from unittest.mock import patch
from pathlib import Path

import pandas as pd
from fastapi.testclient import TestClient
from docx import Document

from api.index import (
    _build_preview_payload,
    _find_attachment_headers,
    _resolve_row_attachments,
    _split_attachment_names,
    app,
    session_info,
)
from api.core.processor import export_html_to_docx_bytes


def make_xlsx_bytes(sheets: dict[str, list[dict[str, object]]]) -> bytes:
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        for sheet_name, rows in sheets.items():
            pd.DataFrame(rows).to_excel(writer, sheet_name=sheet_name, index=False)
    return buffer.getvalue()


class BuildPreviewPayloadTests(unittest.TestCase):
    def test_attachment_detection_uses_single_attachment_column(self):
        self.assertEqual(_find_attachment_headers(["email", "附件", "附件1", "附件2"]), ["附件"])

    def test_split_attachment_names_accepts_common_cell_separators(self):
        self.assertEqual(
            _split_attachment_names("簡章.pdf、報名表.pdf\n地圖.pdf, 合約.pdf；說明.pdf"),
            ["簡章.pdf", "報名表.pdf", "地圖.pdf", "合約.pdf", "說明.pdf"],
        )

    def test_resolve_row_attachments_uses_multiple_names_from_single_cell(self):
        selected_local = {
            "簡章.pdf": {"filename": "簡章.pdf", "content": b"intro", "mime_type": "application/pdf"},
            "報名表.pdf": {"filename": "報名表.pdf", "content": b"form", "mime_type": "application/pdf"},
        }

        attachments, errors = _resolve_row_attachments(
            row={"附件": "簡章.pdf、報名表.pdf"},
            row_index=1,
            attachment_headers_raw=["附件"],
            selected_drive_attachments={},
            selected_local_attachments=selected_local,
            base_dir=Path("/path/not/exist"),
            creds=None,
            drive_service=None,
            attachment_drive_folder_id=None,
        )

        self.assertEqual(errors, [])
        self.assertEqual([item["filename"] for item in attachments], ["簡章.pdf", "報名表.pdf"])

    @patch("api.index.inject_variables", side_effect=lambda template, row: template.replace("{{name}}", str(row["name"])))
    @patch("api.index.resolve_template_html", return_value="<p>{{name}}</p>")
    def test_build_preview_payload_exposes_sheet_names_and_default_selection(self, _resolve_template_html, _inject_variables):
        xlsx_content = make_xlsx_bytes(
            {
                "名單一": [{"name": "Alice", "email": "alice@example.com", "subject": "Hello"}],
                "名單二": [{"name": "Bob", "email": "bob@example.com", "subject": "World"}],
            }
        )

        payload = _build_preview_payload(
            docx_content=b"unused",
            template_html=None,
            xlsx_content=xlsx_content,
            sheet=None,
            font=None,
            cache_info=None,
            preview_page=1,
            preview_page_size=20,
        )

        self.assertEqual(payload["sheet_names"], ["名單一", "名單二"])
        self.assertEqual(payload["selected_sheet"], "名單一")
        self.assertEqual(payload["first_row"], {"name": "Alice", "email": "alice@example.com", "subject": "Hello"})

    @patch("api.index.inject_variables", side_effect=lambda template, row: template.replace("{{name}}", str(row["name"])))
    @patch("api.index.resolve_template_html", return_value="<p>{{name}}</p>")
    def test_build_preview_payload_uses_selected_sheet(self, _resolve_template_html, _inject_variables):
        xlsx_content = make_xlsx_bytes(
            {
                "第一批": [{"name": "Alice", "email": "alice@example.com", "subject": "Hello"}],
                "第二批": [{"name": "Bob", "email": "bob@example.com", "subject": "World"}],
            }
        )

        payload = _build_preview_payload(
            docx_content=b"unused",
            template_html=None,
            xlsx_content=xlsx_content,
            sheet="第二批",
            font=None,
            cache_info=None,
            preview_page=1,
            preview_page_size=20,
        )

        self.assertEqual(payload["selected_sheet"], "第二批")
        self.assertEqual(payload["first_row"], {"name": "Bob", "email": "bob@example.com", "subject": "World"})

    @patch("api.index.inject_variables", side_effect=lambda template, row: template.replace("{{name}}", str(row["name"])))
    def test_build_preview_payload_prefers_template_html_when_provided(self, _inject_variables):
        xlsx_content = make_xlsx_bytes(
            {
                "名單": [{"name": "Alice", "email": "alice@example.com", "subject": "Hello"}],
            }
        )

        payload = _build_preview_payload(
            docx_content=None,
            template_html="<p>{{name}}</p>",
            xlsx_content=xlsx_content,
            sheet=None,
            font=None,
            cache_info=None,
            preview_page=1,
            preview_page_size=20,
        )

        self.assertEqual(
            payload["template_html"],
            '<div style="font-family: Arial, Helvetica, sans-serif;"><p>{{name}}</p></div>',
        )
        self.assertEqual(
            payload["preview_first_row"],
            '<div style="font-family: Arial, Helvetica, sans-serif;"><p>Alice</p></div>',
        )

    @patch("api.index.inject_variables", side_effect=lambda template, row: template.replace("{{name}}", str(row["name"])))
    @patch("api.index.resolve_template_html", return_value="<p>{{name}}</p>")
    def test_build_preview_payload_paginates_preview_rows(self, _resolve_template_html, _inject_variables):
        rows = [
            {
                "name": f"User {index}",
                "email": "" if index == 0 else f"user{index}@example.com",
                "subject": f"Subject {index}",
            }
            for index in range(55)
        ]
        xlsx_content = make_xlsx_bytes({"名單": rows})

        payload = _build_preview_payload(
            docx_content=b"unused",
            template_html=None,
            xlsx_content=xlsx_content,
            sheet=None,
            font=None,
            cache_info=None,
            preview_page=2,
            preview_page_size=20,
        )

        self.assertEqual(payload["total_records"], 55)
        self.assertEqual(payload["preview_pagination"], {"page": 2, "page_size": 20, "total_pages": 3})
        self.assertEqual(len(payload["preview_rows"]), 20)
        self.assertEqual(payload["preview_rows"][0]["name"], "User 20")
        self.assertEqual(payload["preview_rows"][19]["subject"], "Subject 39")

    @patch("api.index._resolve_row_attachments")
    @patch("api.index.inject_variables", side_effect=lambda template, row: template.replace("{{name}}", str(row["name"])))
    @patch("api.index.resolve_template_html", return_value="<p>{{name}}</p>")
    def test_build_preview_payload_exposes_first_row_attachments(
        self,
        _resolve_template_html,
        _inject_variables,
        mock_resolve_row_attachments,
    ):
        xlsx_content = make_xlsx_bytes(
            {
                "名單": [
                    {"name": "Alice", "email": "alice@example.com", "subject": "Hello", "附件": "簡章.pdf"},
                ],
            }
        )
        mock_resolve_row_attachments.return_value = (
            [
                {
                    "filename": "簡章.pdf",
                    "mime_type": "application/pdf",
                    "source": "local_upload",
                    "content": b"pdf-bytes",
                }
            ],
            [],
        )

        payload = _build_preview_payload(
            docx_content=None,
            template_html="<p>{{name}}</p>",
            xlsx_content=xlsx_content,
            sheet=None,
            font=None,
            cache_info={"cache_id": "a" * 32, "docx_name": "編輯器內容", "xlsx_name": "名單.xlsx"},
            preview_page=1,
            preview_page_size=20,
            attachment_context={
                "selected_drive_attachments": {},
                "selected_local_attachments": {},
                "attachments_dir": None,
                "attachment_drive_folder_id": None,
                "creds": None,
                "drive_service": None,
            },
        )

        self.assertEqual(
            payload["first_row_attachments"],
            [
                {
                    "name": "簡章.pdf",
                    "mime_type": "application/pdf",
                    "source": "local_upload",
                    "previewable": True,
                    "attachment_index": 0,
                    "row_index": 0,
                }
            ],
        )

    def test_export_html_to_docx_preserves_font_face_from_font_tag(self):
        content = export_html_to_docx_bytes('<p><font face="Garamond">Hello</font></p>')
        document = Document(io.BytesIO(content))
        self.assertEqual(document.paragraphs[0].runs[0].font.name, "Garamond")


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

    @patch("api.index.load_user_credentials", return_value=object())
    @patch("api.index.build")
    def test_drive_files_lists_recent_attachment_files_without_folders(self, mock_build, _load_user_credentials):
        class FakeListRequest:
            def execute(self):
                return {
                    "files": [
                        {
                            "id": "pdf-file",
                            "name": "簡章.pdf",
                            "mimeType": "application/pdf",
                            "modifiedTime": "2026-06-14T10:00:00Z",
                        },
                        {
                            "id": "folder-file",
                            "name": "附件資料夾",
                            "mimeType": "application/vnd.google-apps.folder",
                            "modifiedTime": "2026-06-13T10:00:00Z",
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

        response = self.client.get("/api/drive/files", params={"kind": "attachment"})

        self.assertEqual(response.status_code, 200)
        payload = response.json()
        self.assertEqual(
            payload["files"],
            [
                {
                    "id": "pdf-file",
                    "name": "簡章.pdf",
                    "mime_type": "application/pdf",
                    "modified_time": "2026-06-14T10:00:00Z",
                    "kind": "attachment",
                    "is_google_workspace": False,
                }
            ],
        )


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


class SessionProfileTests(unittest.TestCase):
    def setUp(self) -> None:
        self.client = TestClient(app)

    def test_session_info_returns_name_and_email_from_session(self):
        request = type(
            "FakeRequest",
            (),
            {"session": {"user_key": "dev_user", "user_email": "alice@example.com", "user_name": "Alice"}},
        )()
        response = session_info(request)
        self.assertEqual(
            response,
            {"user_key": "dev_user", "email": "alice@example.com", "name": "Alice"},
        )

    @patch("api.index.load_user_credentials", return_value=object())
    @patch("api.index._fetch_google_profile", return_value={"email": "alice@example.com", "name": "Alice"})
    def test_session_restore_fetches_and_returns_name_and_email(self, _fetch_google_profile, _load_user_credentials):
        response = self.client.post(
            "/api/session/restore",
            json={"user_key": "aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa"},
        )

        self.assertEqual(response.status_code, 200)
        self.assertEqual(
            response.json(),
            {"ok": True, "restored": True, "email": "alice@example.com", "name": "Alice"},
        )


class TemplateExportEndpointTests(unittest.TestCase):
    def setUp(self) -> None:
        self.client = TestClient(app)

    def test_template_export_returns_docx_attachment(self):
        response = self.client.post(
            "/api/template/export",
            json={"template_html": "<div><p>Hello</p><p><strong>World</strong></p></div>", "filename": "sample"},
        )

        self.assertEqual(response.status_code, 200)
        self.assertEqual(
            response.headers["content-type"],
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )
        self.assertIn('filename="sample.docx"', response.headers["content-disposition"])
        self.assertGreater(len(response.content), 0)


class BatchDraftAttachmentSelectionTests(unittest.TestCase):
    def setUp(self) -> None:
        self.client = TestClient(app)
        self.client.post("/api/dev/login")

    @patch("api.index.create_draft", return_value={"id": "draft-1"})
    @patch("api.index.resolve_template_html", return_value="<p>Hello</p>")
    @patch("api.index.load_user_credentials", return_value=object())
    @patch("api.index.build")
    @patch("api.index._download_drive_file")
    def test_selected_drive_attachments_are_used_before_name_lookup(
        self,
        mock_download_drive_file,
        mock_build,
        _load_user_credentials,
        _resolve_template_html,
        mock_create_draft,
    ):
        xlsx_content = make_xlsx_bytes(
            {
                "名單": [
                    {
                        "email": "alice@example.com",
                        "subject": "Hello",
                        "附件": "簡章.pdf",
                    }
                ]
            }
        )

        mock_download_drive_file.return_value = {
            "file_id": "drive-attachment-1",
            "name": "簡章.pdf",
            "mime_type": "application/pdf",
            "content": b"pdf-bytes",
        }
        mock_build.return_value = object()

        response = self.client.post(
            "/api/drafts/batch",
            params={"font": "Sans Serif"},
            data={
                "attachment_drive_file_ids_json": '["drive-attachment-1"]',
            },
            files={
                "docx_file": (
                    "template.docx",
                    b"unused-docx",
                    "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                ),
                "xlsx_file": (
                    "list.xlsx",
                    xlsx_content,
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                ),
            },
        )

        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.json()["draft_count"], 1)
        mock_download_drive_file.assert_called_once()
        attachments = mock_create_draft.call_args.kwargs["attachments"]
        self.assertEqual(
            attachments,
            [
                {
                    "filename": "簡章.pdf",
                    "content": b"pdf-bytes",
                    "mime_type": "application/pdf",
                }
            ],
        )

    @patch("api.index.create_draft", return_value={"id": "draft-1"})
    @patch("api.index.resolve_template_html", return_value="<p>Hello</p>")
    @patch("api.index.load_user_credentials", return_value=object())
    @patch("api.index.get_attachment_content")
    def test_invalid_attachment_drive_file_ids_payload_returns_400(
        self,
        _get_attachment_content,
        _load_user_credentials,
        _resolve_template_html,
        _create_draft,
    ):
        xlsx_content = make_xlsx_bytes(
            {
                "名單": [
                    {
                        "email": "alice@example.com",
                        "subject": "Hello",
                    }
                ]
            }
        )

        response = self.client.post(
            "/api/drafts/batch",
            data={"attachment_drive_file_ids_json": '{"bad":"payload"}'},
            files={
                "docx_file": (
                    "template.docx",
                    b"unused-docx",
                    "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                ),
                "xlsx_file": (
                    "list.xlsx",
                    xlsx_content,
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                ),
            },
        )

        self.assertEqual(response.status_code, 400)
        self.assertEqual(response.json()["detail"]["error"], "invalid_attachment_drive_file_ids")

    @patch("api.index.create_draft", return_value={"id": "draft-1"})
    @patch("api.index.resolve_template_html", return_value="<p>Hello</p>")
    @patch("api.index.load_user_credentials", return_value=object())
    @patch("api.index.build")
    @patch("api.index.get_attachment_content")
    def test_attachment_folder_id_is_passed_to_drive_lookup(
        self,
        mock_get_attachment_content,
        mock_build,
        _load_user_credentials,
        _resolve_template_html,
        _create_draft,
    ):
        xlsx_content = make_xlsx_bytes(
            {
                "名單": [
                    {
                        "email": "alice@example.com",
                        "subject": "Hello",
                        "附件": "簡章.pdf",
                    }
                ]
            }
        )

        mock_get_attachment_content.return_value = {
            "filename": "簡章.pdf",
            "content": b"pdf-bytes",
            "mime_type": "application/pdf",
        }
        mock_build.return_value = object()

        response = self.client.post(
            "/api/drafts/batch",
            data={"attachment_drive_folder_id": "folder-123"},
            files={
                "docx_file": (
                    "template.docx",
                    b"unused-docx",
                    "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                ),
                "xlsx_file": (
                    "list.xlsx",
                    xlsx_content,
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                ),
            },
        )

        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.json()["draft_count"], 1)
        mock_get_attachment_content.assert_called_once()
        self.assertEqual(mock_get_attachment_content.call_args.kwargs["parent_folder_id"], "folder-123")


class AttachmentPreviewEndpointTests(unittest.TestCase):
    def setUp(self) -> None:
        self.client = TestClient(app)
        self.client.post("/api/dev/login")

    @patch("api.index.inject_variables", side_effect=lambda template, row: template.replace("{{name}}", str(row["name"])))
    def test_attachment_preview_reads_cached_local_attachment(self, _inject_variables):
        xlsx_content = make_xlsx_bytes(
            {
                "名單": [
                    {"name": "Alice", "email": "alice@example.com", "subject": "Hello", "附件": "簡章.pdf"},
                ]
            }
        )

        response = self.client.post(
            "/api/process",
            params={"font": "Sans Serif"},
            data={"template_html": "<p>{{name}}</p>"},
            files=[
                (
                    "xlsx_file",
                    (
                        "list.xlsx",
                        xlsx_content,
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    ),
                ),
                (
                    "attachment_local_files",
                    ("簡章.pdf", b"pdf-bytes", "application/pdf"),
                ),
            ],
        )

        self.assertEqual(response.status_code, 200)
        payload = response.json()
        self.assertEqual(payload["first_row_attachments"][0]["name"], "簡章.pdf")

        preview_response = self.client.get(
            "/api/attachments/content",
            params={
                "cache_id": payload["cache_id"],
                "row": 0,
                "attachment_index": 0,
            },
        )

        self.assertEqual(preview_response.status_code, 200)
        self.assertEqual(preview_response.content, b"pdf-bytes")
        self.assertEqual(preview_response.headers["content-type"], "application/pdf")

    @patch("api.index.create_draft", return_value={"id": "draft-1"})
    @patch("api.index.resolve_template_html", return_value="<p>Hello</p>")
    @patch("api.index.load_user_credentials", return_value=object())
    @patch("api.index.get_attachment_content")
    def test_uploaded_local_attachment_files_are_used_before_drive_lookup(
        self,
        mock_get_attachment_content,
        _load_user_credentials,
        _resolve_template_html,
        mock_create_draft,
    ):
        xlsx_content = make_xlsx_bytes(
            {
                "名單": [
                    {
                        "email": "alice@example.com",
                        "subject": "Hello",
                        "附件": "簡章.pdf",
                    }
                ]
            }
        )

        response = self.client.post(
            "/api/drafts/batch",
            files=[
                (
                    "docx_file",
                    (
                        "template.docx",
                        b"unused-docx",
                        "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    ),
                ),
                (
                    "xlsx_file",
                    (
                        "list.xlsx",
                        xlsx_content,
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    ),
                ),
                (
                    "attachment_local_files",
                    (
                        "簡章.pdf",
                        b"local-pdf-bytes",
                        "application/pdf",
                    ),
                ),
            ],
        )

        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.json()["draft_count"], 1)
        mock_get_attachment_content.assert_not_called()
        attachments = mock_create_draft.call_args.kwargs["attachments"]
        self.assertEqual(attachments[0]["filename"], "簡章.pdf")
        self.assertEqual(attachments[0]["content"], b"local-pdf-bytes")


if __name__ == "__main__":
    unittest.main()
