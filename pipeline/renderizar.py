"""Renderiza .docx em PDF com o LibreOffice (headless), para QA visual e verificar_paginacao.py.

O LibreOffice não é o programa do cliente: o teste final continua sendo no Word ou no
Google Docs (06, seção 5). Este PDF serve para conferir paginação, sumário e quebras.

Uso:
  python3 -I pipeline/renderizar.py arquivo.docx [pasta_saida]
"""
from __future__ import annotations

import shutil
import subprocess
import sys
import tempfile
from pathlib import Path


def renderizar(docx: Path, saida: Path) -> Path:
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        raise SystemExit("LibreOffice não encontrado (apt install libreoffice-writer).")
    saida.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory() as perfil:  # perfil isolado: evita travar em sessões paralelas
        subprocess.run([soffice, f"-env:UserInstallation=file://{perfil}", "--headless",
                        "--convert-to", "pdf", "--outdir", str(saida), str(docx)],
                       check=True, capture_output=True, timeout=300)
    pdf = saida / (docx.stem + ".pdf")
    if not pdf.exists():
        raise SystemExit("A conversão não gerou o PDF.")
    return pdf


if __name__ == "__main__":
    if len(sys.argv) < 2:
        raise SystemExit(__doc__)
    docx = Path(sys.argv[1]).resolve()
    destino = Path(sys.argv[2]).resolve() if len(sys.argv) > 2 else docx.parent
    print(renderizar(docx, destino))
