from __future__ import annotations

import subprocess
import sys


def main() -> None:
    command = [
        sys.executable,
        "-m",
        "PyInstaller",
        "--noconfirm",
        "--clean",
        "--windowed",
        "--name",
        "GeradorEbookFiscal",
        "--add-data",
        r"ebook_fiscal\templates\Ebook_Fiscal_Template.docx;ebook_fiscal\templates",
        "--hidden-import",
        "openpyxl.cell._writer",
        "gerador_ebook_fiscal.py",
    ]
    subprocess.run(command, check=True)
    print(r"Executavel criado em: dist\GeradorEbookFiscal\GeradorEbookFiscal.exe")


if __name__ == "__main__":
    main()
