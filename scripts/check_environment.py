#!/usr/bin/env python
"""Check runtime dependencies for the experiment report grading skill.

This script only reports missing dependencies. It never installs packages by
itself, because installation may need user approval, network access, or a
project-specific Python environment.
"""

from __future__ import annotations

import argparse
import importlib.util
import shutil
import sys
from pathlib import Path


CORE_PACKAGES = [
    ("python-docx", "docx"),
    ("openpyxl", "openpyxl"),
]
OCR_PACKAGES = [
    ("Pillow", "PIL"),
    ("pytesseract", "pytesseract"),
]


def missing_packages(packages: list[tuple[str, str]]) -> list[str]:
    missing = []
    for pip_name, module_name in packages:
        if importlib.util.find_spec(module_name) is None:
            missing.append(pip_name)
    return missing


def rel_requirements(name: str) -> str:
    return str((Path(__file__).resolve().parents[1] / name).resolve())


def print_install_hint(requirements_file: str) -> None:
    print("Install after user approval with:")
    print(f"  {sys.executable} -m pip install -r \"{requirements_file}\"")


def main() -> int:
    parser = argparse.ArgumentParser(description="Check skill runtime dependencies.")
    parser.add_argument("--ocr", action="store_true", help="Also check optional OCR dependencies")
    args = parser.parse_args()

    failed = False
    if sys.version_info < (3, 10):
        failed = True
        print(f"Python >= 3.10 is recommended; current: {sys.version.split()[0]}")

    core_missing = missing_packages(CORE_PACKAGES)
    if core_missing:
        failed = True
        print("Missing required Python packages: " + ", ".join(core_missing))
        print_install_hint(rel_requirements("requirements.txt"))
    else:
        print("Required Python packages are available.")

    if args.ocr:
        ocr_missing = missing_packages(OCR_PACKAGES)
        tesseract = shutil.which("tesseract")
        if ocr_missing:
            failed = True
            print("Missing optional OCR Python packages: " + ", ".join(ocr_missing))
            print_install_hint(rel_requirements("requirements-ocr.txt"))
        else:
            print("Optional OCR Python packages are available.")
        if not tesseract:
            failed = True
            print("Tesseract executable was not found on PATH.")
            print("Install Tesseract OCR separately and add it to PATH before OCR tasks.")
        else:
            print(f"Tesseract executable: {tesseract}")

    return 1 if failed else 0


if __name__ == "__main__":
    raise SystemExit(main())
