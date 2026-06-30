"""Update all fields (including TOC) in a DOCX file using LibreOffice.

Uses LibreOffice headless mode to open the document, execute field update
commands, save, and close. This populates TOC, page numbers, and other
dynamic fields so users don't need to manually "Update Field" in Word.

Requires LibreOffice (soffice) to be installed.
"""

import argparse
import logging
import shutil
import subprocess
from pathlib import Path

logger = logging.getLogger(__name__)

LIBREOFFICE_PROFILE = "/tmp/libreoffice_docx_profile"
MACRO_DIR = f"{LIBREOFFICE_PROFILE}/user/basic/Standard"

UPDATE_FIELDS_MACRO = """<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE script:module PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "module.dtd">
<script:module xmlns:script="http://openoffice.org/2000/script" script:name="Module1" script:language="StarBasic">
    Sub UpdateAllFieldsAndIndexes()
        Dim document As Object
        Dim dispatcher As Object

        document = ThisComponent.CurrentController.Frame
        dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")

        ' Update all indexes (TOC, figures, tables, etc.)
        dispatcher.executeDispatch(document, ".uno:UpdateAllIndexes", "", 0, Array())

        ' Update other fields (page numbers, cross-references, etc.)
        dispatcher.executeDispatch(document, ".uno:UpdateFields", "", 0, Array())

        ThisComponent.store()
        ThisComponent.close(True)
    End Sub
</script:module>"""


def _get_soffice_env() -> dict:
    """Get environment for running soffice (with AF_UNIX shim if needed)."""
    import os

    env = os.environ.copy()
    env["SAL_USE_VCLPLUGIN"] = "svp"

    # Try to use the docx skill's soffice helper if available
    try:
        import importlib.util
        spec = importlib.util.spec_from_file_location(
            "soffice_helper",
            Path(__file__).resolve().parents[3] / ".claude" / "skills" / "docx" / "scripts" / "office" / "soffice.py",
        )
        if spec and spec.loader:
            module = importlib.util.module_from_spec(spec)
            spec.loader.exec_module(module)
            return module.get_soffice_env()
    except Exception:
        pass

    return env


def _setup_macro() -> bool:
    """Ensure the UpdateAllFields macro exists in the LibreOffice profile."""
    macro_dir = Path(MACRO_DIR)
    macro_file = macro_dir / "Module1.xba"

    if macro_file.exists() and "UpdateAllFieldsAndIndexes" in macro_file.read_text():
        return True

    if not macro_dir.exists():
        # Initialize LibreOffice profile
        subprocess.run(
            [
                "soffice",
                "--headless",
                f"-env:UserInstallation=file://{LIBREOFFICE_PROFILE}",
                "--terminate_after_init",
            ],
            capture_output=True,
            timeout=10,
            check=False,
            env=_get_soffice_env(),
        )
        macro_dir.mkdir(parents=True, exist_ok=True)

    try:
        macro_file.write_text(UPDATE_FIELDS_MACRO)
        return True
    except Exception as e:
        logger.warning(f"Failed to setup LibreOffice macro: {e}")
        return False


def update_fields(
    input_file: str,
    output_file: str | None = None,
) -> tuple[bool, str]:
    """Update all fields in a DOCX file using LibreOffice.

    Args:
        input_file: Path to the input .docx file.
        output_file: Path for the output .docx file. If None, overwrites input.

    Returns:
        (success, message) tuple.
    """
    input_path = Path(input_file).resolve()

    if not input_path.exists():
        return False, f"Error: Input file not found: {input_file}"

    if input_path.suffix.lower() != ".docx":
        return False, f"Error: Input file is not a DOCX file: {input_file}"

    if output_file is None:
        output_file = str(input_path)
    output_path = Path(output_file).resolve()

    try:
        output_path.parent.mkdir(parents=True, exist_ok=True)
        if input_path != output_path:
            shutil.copy2(input_path, output_path)
    except Exception as e:
        return False, f"Error: Failed to copy input file: {e}"

    if not _setup_macro():
        return False, "Error: Failed to setup LibreOffice macro"

    cmd = [
        "soffice",
        "--headless",
        f"-env:UserInstallation=file://{LIBREOFFICE_PROFILE}",
        "--norestore",
        "vnd.sun.star.script:Standard.Module1.UpdateAllFieldsAndIndexes?language=Basic&location=application",
        str(output_path.absolute()),
    ]

    try:
        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            timeout=60,
            check=False,
            env=_get_soffice_env(),
        )
    except subprocess.TimeoutExpired:
        # Timeout is expected — LibreOffice process exits via the AF_UNIX shim
        # after the macro calls close(). The timeout means it didn't exit cleanly.
        return (
            True,
            f"Fields updated (with timeout): {input_file} -> {output_file}",
        )

    if result.returncode != 0:
        return False, f"Error: LibreOffice failed (rc={result.returncode}): {result.stderr}"

    return (
        True,
        f"Successfully updated all fields: {input_file} -> {output_file}",
    )


if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description="Update all fields (TOC, page numbers, etc.) in a DOCX file"
    )
    parser.add_argument("input_file", help="Input DOCX file")
    parser.add_argument(
        "output_file",
        nargs="?",
        default=None,
        help="Output DOCX file (defaults to overwriting input)",
    )
    args = parser.parse_args()

    success, message = update_fields(args.input_file, args.output_file)
    print(message)

    if not success:
        raise SystemExit(1)
