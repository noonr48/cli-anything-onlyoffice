import os
import subprocess
import tempfile
import unittest
from pathlib import Path
from unittest import mock

from cli_anything.onlyoffice.utils.docserver import DocumentServerClient


class OnlyOfficeDocserverSafetyTests(unittest.TestCase):
    def _client(self, base: str) -> DocumentServerClient:
        env = {
            "ONLYOFFICE_BACKUP_DIR": str(Path(base) / "backups"),
            "ONLYOFFICE_LOCK_DIR": str(Path(base) / "locks"),
        }
        with mock.patch.dict(os.environ, env, clear=False):
            return DocumentServerClient()

    def test_default_backup_and_lock_dirs_are_private_under_home(self):
        with tempfile.TemporaryDirectory(prefix="oo-home-") as home:
            with mock.patch.dict(os.environ, {"HOME": home}, clear=False):
                os.environ.pop("ONLYOFFICE_BACKUP_DIR", None)
                os.environ.pop("ONLYOFFICE_LOCK_DIR", None)
                client = DocumentServerClient()

            expected_root = Path(home) / ".cli-anything"
            self.assertEqual(client.backup_dir, expected_root / "backups")
            self.assertEqual(client.lock_dir, expected_root / "locks")
            self.assertEqual(client.backup_dir.stat().st_mode & 0o777, 0o700)
            self.assertEqual(client.lock_dir.stat().st_mode & 0o777, 0o700)

    def test_symlinked_backup_and_lock_dirs_are_rejected(self):
        with tempfile.TemporaryDirectory(prefix="oo-state-symlink-") as base:
            base_path = Path(base)
            real_backup = base_path / "real-backup"
            real_backup.mkdir()
            backup_link = base_path / "backup-link"
            backup_link.symlink_to(real_backup, target_is_directory=True)

            with mock.patch.dict(
                os.environ,
                {
                    "ONLYOFFICE_BACKUP_DIR": str(backup_link),
                    "ONLYOFFICE_LOCK_DIR": str(base_path / "locks"),
                },
                clear=False,
            ):
                with self.assertRaisesRegex(RuntimeError, "backup directory must not be a symlink"):
                    DocumentServerClient()

            real_lock = base_path / "real-lock"
            real_lock.mkdir()
            lock_link = base_path / "lock-link"
            lock_link.symlink_to(real_lock, target_is_directory=True)

            with mock.patch.dict(
                os.environ,
                {
                    "ONLYOFFICE_BACKUP_DIR": str(base_path / "backups"),
                    "ONLYOFFICE_LOCK_DIR": str(lock_link),
                },
                clear=False,
            ):
                with self.assertRaisesRegex(RuntimeError, "lock directory must not be a symlink"):
                    DocumentServerClient()

    def test_file_lock_uses_stable_private_lock_file(self):
        with tempfile.TemporaryDirectory(prefix="oo-lock-") as base:
            client = self._client(base)
            target = Path(base) / "document.docx"
            target.write_bytes(b"fixture")

            with client._file_lock(str(target)):
                lock_path = client._lock_path(str(target))
                self.assertTrue(lock_path.exists())
                self.assertEqual(lock_path.parent, client.lock_dir)
                self.assertEqual(lock_path.stat().st_mode & 0o777, 0o600)

            self.assertTrue(lock_path.exists())

    def test_restore_rejects_backup_outside_backup_dir(self):
        with tempfile.TemporaryDirectory(prefix="oo-restore-outside-") as base:
            client = self._client(base)
            target = Path(base) / "target.docx"
            target.write_bytes(b"current")
            outside_dir = Path(base) / "outside"
            outside_dir.mkdir()
            outside_backup = outside_dir / (
                f"{target.name}.{client._file_key(str(target))}."
                "20260505T010203000000Z.bak.docx"
            )
            outside_backup.write_bytes(b"outside")

            result = client.restore_backup(str(target), backup=str(outside_backup))

            self.assertFalse(result["success"])
            self.assertIn("under the configured backup_dir", result["error"])
            self.assertEqual(target.read_bytes(), b"current")

    def test_restore_rejects_backup_for_different_target(self):
        with tempfile.TemporaryDirectory(prefix="oo-restore-mismatch-") as base:
            client = self._client(base)
            target = Path(base) / "target.docx"
            other = Path(base) / "other.docx"
            target.write_bytes(b"current")
            backup = client.backup_dir / (
                f"{other.name}.{client._file_key(str(other))}."
                "20260505T010203000000Z.bak.docx"
            )
            backup.write_bytes(b"other")

            result = client.restore_backup(str(target), backup=backup.name)

            self.assertFalse(result["success"])
            self.assertIn("does not match", result["error"])
            self.assertEqual(target.read_bytes(), b"current")

    def test_snapshot_backup_writes_final_backup_content(self):
        with tempfile.TemporaryDirectory(prefix="oo-snapshot-") as base:
            client = self._client(base)
            target = Path(base) / "paper.docx"
            target.write_bytes(b"version one")

            backup = Path(client._snapshot_backup(str(target)))

            self.assertEqual(backup.parent, client.backup_dir)
            self.assertTrue(backup.exists())
            self.assertEqual(backup.read_bytes(), b"version one")
            self.assertFalse(list(client.backup_dir.glob("*.tmp")))

    def test_office_to_pdf_rejects_same_input_output_and_non_pdf_output(self):
        with tempfile.TemporaryDirectory(prefix="oo-pdf-reject-") as base:
            client = self._client(base)
            pdf_input = Path(base) / "source.pdf"
            docx_input = Path(base) / "source.docx"
            pdf_input.write_bytes(b"%PDF-1.4")
            docx_input.write_bytes(b"docx")

            with mock.patch("subprocess.run") as run:
                same = client._office_to_pdf(str(pdf_input), output_path=str(pdf_input))
                non_pdf = client._office_to_pdf(
                    str(docx_input), output_path=str(Path(base) / "out.txt")
                )

            self.assertFalse(same["success"])
            self.assertIn("same as input", same["error"])
            self.assertFalse(non_pdf["success"])
            self.assertIn(".pdf", non_pdf["error"])
            run.assert_not_called()

            normalized = client._resolve_pdf_output_path(
                str(docx_input), str(Path(base) / "normalized")
            )
            self.assertEqual(normalized, Path(base).resolve() / "normalized.pdf")

    def test_office_to_pdf_stages_output_then_replaces_with_backup(self):
        with tempfile.TemporaryDirectory(prefix="oo-pdf-stage-") as base:
            client = self._client(base)
            source = Path(base) / "source.docx"
            output = Path(base) / "converted.pdf"
            source.write_bytes(b"docx")
            output.write_bytes(b"old pdf")
            converted_bytes = b"%PDF-1.4\n% staged fake\n"
            calls = []

            def fake_run(args, **kwargs):
                calls.append(list(args))
                if args[:2] == ["docker", "cp"] and args[2] == str(source.resolve()):
                    self.assertTrue(args[3].startswith("onlyoffice-documentserver:/tmp/convert_"))
                    return subprocess.CompletedProcess(args, 0, stdout=b"", stderr=b"")

                if args[:2] == ["docker", "cp"] and str(args[2]).endswith(".xml"):
                    xml_path = Path(args[2])
                    self.assertTrue(xml_path.exists())
                    xml = xml_path.read_text(encoding="utf-8")
                    self.assertIn("<TaskQueueDataConvert>", xml)
                    self.assertNotIn("XMLEOF", xml)
                    self.assertTrue(args[3].startswith("onlyoffice-documentserver:/tmp/convert_"))
                    return subprocess.CompletedProcess(args, 0, stdout=b"", stderr=b"")

                if (
                    args[:3] == ["docker", "exec", "onlyoffice-documentserver"]
                    and str(args[3]).endswith("/x2t")
                ):
                    return subprocess.CompletedProcess(args, 0, stdout=b"", stderr=b"")

                if (
                    args[:2] == ["docker", "cp"]
                    and str(args[2]).startswith("onlyoffice-documentserver:/tmp/convert_")
                    and str(args[2]).endswith(".pdf")
                ):
                    staged = Path(args[3])
                    self.assertNotEqual(staged.resolve(), output.resolve())
                    self.assertEqual(staged.parent, output.parent)
                    staged.write_bytes(converted_bytes)
                    return subprocess.CompletedProcess(args, 0, stdout=b"", stderr=b"")

                if args[:4] == ["docker", "exec", "onlyoffice-documentserver", "rm"]:
                    return subprocess.CompletedProcess(args, 0, stdout=b"", stderr=b"")

                self.fail(f"unexpected subprocess call: {args}")

            with mock.patch("subprocess.run", side_effect=fake_run):
                result = client._office_to_pdf(str(source), output_path=str(output))

            self.assertTrue(result["success"], result)
            self.assertEqual(output.read_bytes(), converted_bytes)
            self.assertEqual(result["output_file"], str(output.resolve()))
            self.assertEqual(result["file_size"], len(converted_bytes))
            self.assertIsNotNone(result["backup"])
            self.assertEqual(Path(result["backup"]).read_bytes(), b"old pdf")
            self.assertFalse(any("bash" in part for call in calls for part in call))

    def test_office_to_pdf_rejects_empty_staged_output_without_replacing(self):
        with tempfile.TemporaryDirectory(prefix="oo-pdf-empty-") as base:
            client = self._client(base)
            source = Path(base) / "source.docx"
            output = Path(base) / "converted.pdf"
            source.write_bytes(b"docx")

            def fake_run(args, **kwargs):
                if (
                    args[:2] == ["docker", "cp"]
                    and str(args[2]).startswith("onlyoffice-documentserver:/tmp/convert_")
                    and str(args[2]).endswith(".pdf")
                ):
                    Path(args[3]).write_bytes(b"")
                return subprocess.CompletedProcess(args, 0, stdout=b"", stderr=b"")

            with mock.patch("subprocess.run", side_effect=fake_run):
                result = client._office_to_pdf(str(source), output_path=str(output))

            self.assertFalse(result["success"])
            self.assertIn("empty output", result["error"])
            self.assertFalse(output.exists())


if __name__ == "__main__":
    unittest.main()
