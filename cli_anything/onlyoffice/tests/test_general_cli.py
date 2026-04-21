import io
import json
import unittest
from contextlib import redirect_stdout
from unittest import mock

from cli_anything.onlyoffice.core import cli as cli_module
from cli_anything.onlyoffice.core import general_cli


class OnlyOfficeGeneralCLITests(unittest.TestCase):
    def test_normalize_command_alias_maps_known_namespace(self):
        command, meta = general_cli.normalize_command_alias("document.open")

        self.assertEqual(command, "open")
        self.assertEqual(meta["requested_command"], "document.open")
        self.assertEqual(meta["resolved_command"], "open")
        self.assertEqual(meta["command_namespace"], "document")

    def test_normalize_command_alias_leaves_unknown_namespace_unchanged(self):
        command, meta = general_cli.normalize_command_alias("random.open")

        self.assertEqual(command, "random.open")
        self.assertIsNone(meta)

    def test_handle_general_command_routes_editor_session_flags(self):
        printed = []
        doc_server = mock.Mock()
        doc_server.editor_session.return_value = {"success": True, "session": "abc"}

        handled = general_cli.handle_general_command(
            "editor-session",
            ["/tmp/report.docx", "--open", "--wait", "4.5", "--activate"],
            json_output=True,
            alias_meta=None,
            print_result=lambda payload, json_output: printed.append((payload, json_output)),
            doc_server=doc_server,
            docx_available=True,
            openpyxl_available=True,
            pptx_available=True,
        )

        self.assertTrue(handled)
        doc_server.editor_session.assert_called_once_with(
            "/tmp/report.docx",
            open_if_needed=True,
            wait_seconds=4.5,
            activate=True,
        )
        self.assertEqual(printed, [({"success": True, "session": "abc"}, True)])

    def test_handle_general_command_routes_backup_prune_flags(self):
        printed = []
        doc_server = mock.Mock()
        doc_server.prune_backups.return_value = {"success": True, "deleted": 3}

        handled = general_cli.handle_general_command(
            "backup-prune",
            ["--file", "/tmp/report.docx", "--keep", "7", "--days", "14"],
            json_output=False,
            alias_meta=None,
            print_result=lambda payload, json_output: printed.append((payload, json_output)),
            doc_server=doc_server,
            docx_available=True,
            openpyxl_available=True,
            pptx_available=True,
        )

        self.assertTrue(handled)
        doc_server.prune_backups.assert_called_once_with(
            file_path="/tmp/report.docx",
            keep=7,
            older_than_days=14,
        )
        self.assertEqual(printed, [({"success": True, "deleted": 3}, False)])

    def test_handle_general_command_reports_usage_from_registry(self):
        printed = []

        handled = general_cli.handle_general_command(
            "open",
            [],
            json_output=True,
            alias_meta=None,
            print_result=lambda payload, json_output: printed.append((payload, json_output)),
            doc_server=mock.Mock(),
            docx_available=True,
            openpyxl_available=True,
            pptx_available=True,
        )

        self.assertTrue(handled)
        self.assertEqual(
            printed,
            [({"success": False, "error": general_cli.command_usage("open")}, True)],
        )

    def test_handle_general_command_returns_false_for_unknown_command(self):
        handled = general_cli.handle_general_command(
            "not-a-real-command",
            [],
            json_output=False,
            alias_meta=None,
            print_result=lambda payload, json_output: None,
            doc_server=mock.Mock(),
            docx_available=True,
            openpyxl_available=True,
            pptx_available=True,
        )

        self.assertFalse(handled)

    def test_handle_general_command_backup_list_handles_missing_client(self):
        printed = []

        handled = general_cli.handle_general_command(
            "backup-list",
            ["/tmp/report.docx"],
            json_output=True,
            alias_meta=None,
            print_result=lambda payload, json_output: printed.append((payload, json_output)),
            doc_server=None,
            docx_available=True,
            openpyxl_available=True,
            pptx_available=True,
        )

        self.assertTrue(handled)
        self.assertEqual(
            printed,
            [({"success": False, "error": "Client not available"}, True)],
        )

    def test_cli_backup_prune_dispatches_via_general_handler(self):
        stdout = io.StringIO()

        with mock.patch.object(
            cli_module.doc_server,
            "prune_backups",
            return_value={"success": True, "deleted": 2},
        ) as prune_backups:
            with mock.patch(
                "sys.argv",
                [
                    "cli-anything-onlyoffice",
                    "backup-prune",
                    "--file",
                    "/tmp/report.docx",
                    "--keep",
                    "5",
                    "--days",
                    "21",
                    "--json",
                ],
            ):
                with redirect_stdout(stdout):
                    cli_module.main()

        payload = json.loads(stdout.getvalue())
        self.assertTrue(payload["success"])
        prune_backups.assert_called_once_with(
            file_path="/tmp/report.docx",
            keep=5,
            older_than_days=21,
        )

    def test_cli_backup_restore_dispatches_via_general_handler(self):
        stdout = io.StringIO()

        with mock.patch.object(
            cli_module.doc_server,
            "restore_backup",
            return_value={"success": True, "restored": False, "dry_run": True},
        ) as restore_backup:
            with mock.patch(
                "sys.argv",
                [
                    "cli-anything-onlyoffice",
                    "backup-restore",
                    "/tmp/report.docx",
                    "--latest",
                    "--dry-run",
                    "--json",
                ],
            ):
                with redirect_stdout(stdout):
                    cli_module.main()

        payload = json.loads(stdout.getvalue())
        self.assertTrue(payload["success"])
        restore_backup.assert_called_once_with(
            file_path="/tmp/report.docx",
            backup=None,
            latest=True,
            dry_run=True,
        )
