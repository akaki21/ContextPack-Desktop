from __future__ import annotations

import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]


class InstallerSourceTests(unittest.TestCase):
    def test_inno_setup_has_beginner_install_features(self) -> None:
        source = (ROOT / "installer" / "ContextPack.iss").read_text(encoding="utf-8-sig")
        self.assertIn("PrivilegesRequired=lowest", source)
        self.assertIn("Name: \"desktopicon\"", source)
        self.assertIn("{group}\\ContextPack Desktop", source)
        self.assertIn("{uninstallexe}", source)
        self.assertIn("bootstrap.ps1", source)
        self.assertNotIn("createallsubdirs", source)

    def test_bootstrap_checks_dependencies_and_writes_logs(self) -> None:
        source = (ROOT / "installer" / "bootstrap.ps1").read_text(encoding="utf-8-sig")
        self.assertIn("Python.Python.3.12", source)
        self.assertIn("check-environment.ps1", source)
        self.assertIn("install.log", source)
        self.assertIn("last-install-error.txt", source)

    def test_setup_can_use_detected_python_safely(self) -> None:
        source = (ROOT / "setup.ps1").read_text(encoding="utf-8-sig")
        self.assertIn("CONTEXTPACK_PYTHON", source)
        self.assertIn("ReparsePoint", source)

    def test_installer_assets_exist(self) -> None:
        for relative in (
            "assets/contextpack.ico",
            "assets/contextpack-icon.png",
            "installer/build-installer.ps1",
        ):
            path = ROOT / relative
            self.assertTrue(path.is_file(), f"Missing installer asset: {relative}")
            self.assertGreater(path.stat().st_size, 0)


if __name__ == "__main__":
    unittest.main()
