"""Ponto de entrada do DT 3.0.

O código legado permanece em ``python/dt30_legacy.py`` nesta primeira etapa
de organização. Este arquivo mantém o nome usado pelo empacotamento e delega
a execução ao legado sem alterar o fluxo interativo/automático.
"""

from pathlib import Path
import runpy


LEGACY_PATH = Path(__file__).resolve().parent / "python" / "dt30_legacy.py"


if __name__ == "__main__":
    runpy.run_path(str(LEGACY_PATH), run_name="__main__")
