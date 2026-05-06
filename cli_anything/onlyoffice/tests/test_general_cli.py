import io
import json
import subprocess
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

    def test_detect_conversion_capability_reports_docker_and_x2t(self):
        calls = []

        def fake_run(args, **kwargs):
            calls.append((args, kwargs))
            return subprocess.CompletedProcess(args, 0, stdout=b"", stderr=b"")

        result = general_cli.detect_conversion_capability(
            which=lambda name: "/usr/bin/docker" if name == "docker" else None,
            run=fake_run,
        )

        self.assertTrue(result["available"])
        self.assertTrue(result["docker"]["available"])
        self.assertEqual(result["docker"]["path"], "/usr/bin/docker")
        self.assertTrue(result["x2t"]["available"])
        self.assertEqual(result["x2t"]["container"], "onlyoffice-documentserver")
        self.assertEqual(calls[0][0][0], "/usr/bin/docker")

    def test_cmd_status_includes_conversion_metadata(self):
        printed = []
        doc_server = mock.Mock()
        doc_server.check_health.return_value = False
        conversion = {
            "available": True,
            "docker": {"available": True, "path": "/usr/bin/docker"},
            "x2t": {
                "available": True,
                "container": "onlyoffice-documentserver",
                "path": general_cli.ONLYOFFICE_X2T_PATH,
                "checked": True,
            },
        }

        with mock.patch.object(
            general_cli, "detect_conversion_capability", return_value=conversion
        ):
            general_cli.cmd_status(
                json_output=True,
                print_result=lambda payload, json_output: printed.append((payload, json_output)),
                doc_server=doc_server,
                docx_available=True,
                openpyxl_available=True,
                pptx_available=True,
            )

        payload, json_output = printed[0]
        self.assertTrue(json_output)
        self.assertEqual(payload["conversion"], conversion)
        self.assertTrue(payload["dependencies"]["docker"])
        self.assertTrue(payload["dependencies"]["onlyoffice_x2t"])
        self.assertIn("install_check", payload)
        self.assertTrue(payload["install_check"]["external_dependencies_ok"])
        self.assertTrue(payload["capabilities"]["office_to_pdf"])
        self.assertTrue(payload["capabilities"]["docx_to_pdf"])

    def test_build_installation_check_reports_missing_required_dependencies(self):
        conversion = {
            "available": False,
            "docker": {"available": False, "path": None},
            "x2t": {
                "available": False,
                "container": general_cli.ONLYOFFICE_DOCKER_CONTAINER,
                "path": general_cli.ONLYOFFICE_X2T_PATH,
                "checked": False,
            },
        }
        python_dependencies = [
            {
                "key": "pyshacl",
                "package": "pyshacl",
                "import_name": "pyshacl",
                "required": True,
                "available": False,
                "installed_version": None,
                "minimum_version": "0.25.0",
                "version_satisfied": False,
                "status": "fail",
                "install_requirement": "pyshacl>=0.25.0",
            }
        ]

        payload = general_cli.build_installation_check(
            doc_server=mock.Mock(),
            docx_available=True,
            openpyxl_available=True,
            pptx_available=True,
            conversion_detector=lambda: conversion,
            python_detector=lambda: python_dependencies,
        )

        self.assertFalse(payload["success"])
        self.assertFalse(payload["install_ready"])
        self.assertEqual(payload["missing_python"], ["pyshacl>=0.25.0"])
        self.assertIn("docker", payload["missing_external"])
        self.assertIn("onlyoffice_x2t", payload["missing_external"])
        self.assertGreaterEqual(len(payload["install_hints"]), 2)

    def test_cmd_setup_check_uses_strict_installation_report(self):
        printed = []
        expected = {
            "success": True,
            "install_ready": True,
            "python_dependencies_ok": True,
            "external_dependencies_ok": True,
        }

        with mock.patch.object(
            general_cli,
            "build_installation_check",
            return_value=expected,
        ) as build_check:
            general_cli.cmd_setup_check(
                json_output=True,
                print_result=lambda payload, json_output: printed.append((payload, json_output)),
                doc_server=mock.Mock(),
                docx_available=True,
                openpyxl_available=True,
                pptx_available=True,
            )

        self.assertEqual(printed, [(expected, True)])
        build_check.assert_called_once()

    def test_build_installation_check_can_request_live_smoke(self):
        conversion = {
            "available": True,
            "docker": {"available": True, "path": "/usr/bin/docker"},
            "x2t": {
                "available": True,
                "container": general_cli.ONLYOFFICE_DOCKER_CONTAINER,
                "path": general_cli.ONLYOFFICE_X2T_PATH,
                "checked": True,
            },
        }

        with mock.patch.object(
            general_cli,
            "run_live_docx_pdf_smoke",
            return_value={"success": True, "checks": {"conversion_success": True}},
        ) as live_smoke:
            payload = general_cli.build_installation_check(
                doc_server=mock.Mock(),
                docx_available=True,
                openpyxl_available=True,
                pptx_available=True,
                live_smoke=True,
                conversion_detector=lambda: conversion,
                python_detector=lambda: [],
            )

        self.assertTrue(payload["success"])
        self.assertTrue(payload["live_smoke_requested"])
        self.assertTrue(payload["live_smoke"]["success"])
        live_smoke.assert_called_once()
