"""
PASTAS_DT.py
Gerador automático de pastas para DriveTest (SSV)
Lê RELATORIO_SSV_COMPLETO.txt e cria a estrutura local em DT3.0\\PASTAS PADRÃO

Lógica do parser:
  - Instaladora / Site / Cidade  → extraídos da seção "Nome - LOGS"
  - UF                           → extraída da linha de log SE presente (ex: _..._MG_)
                                   caso contrário, buscada no cabeçalho do bloco (CIDADE - MG)
  - Tecnologias / Frequências    → extraídas da linha "(Frequências):" do bloco
  - Nome da pasta: SEM underscore no início e no fim
        Ex: EOLEN_SSV_MG-UVS_4G_700_PARACATU_MG
"""

import os
import re
import ctypes
import argparse
import sys
from pathlib import Path
import tkinter as tk
from tkinter import messagebox

# ─────────────────────────────────────────────
# CONFIGURAÇÕES DE CAMINHOS
# ─────────────────────────────────────────────
def obter_pasta_projeto():
    """Localiza a pasta DT3.0 sem depender de caminhos de outro computador."""
    if getattr(sys, "frozen", False):
        pasta_execucao = Path(sys.executable).resolve().parent
        if pasta_execucao.name.lower() == "dist" and pasta_execucao.parent.name.upper() == "GERAR PASTAS - DT":
            return pasta_execucao.parent.parent
        return pasta_execucao
    return Path(__file__).resolve().parent.parent


PASTA_PROJETO = obter_pasta_projeto()
RELATORIO_PATH = str(PASTA_PROJETO / "OUT" / "RELATORIO_SSV_COMPLETO.txt")
STANDBY_PATH   = str(PASTA_PROJETO / "PASTAS PADRÃO")
ICON_PATH      = str(PASTA_PROJETO / "GERAR PASTAS - DT" / "ICO")

# ─────────────────────────────────────────────
# MAPEAMENTO DE ÍCONES
# ─────────────────────────────────────────────
def get_ico_path(tech: str, freq: str) -> str:
    """
    - 4G : usa a frequência (700.ico, 1800.ico...). Qualquer 2600* → 2600.ico
    - 5G : sempre 3500.ico
    - 3G : sempre 3G.ico  (independente da frequência: 850, 2100...)
    - 2G : sempre 2G.ico  (independente da frequência: 900, 1800, 1900...)
    """
    if tech == "4G":
        ico_file = "2600.ico" if freq.upper().startswith("2600") else f"{freq}.ico"
    elif tech == "5G":
        ico_file = "3500.ico"
    elif tech == "3G":
        ico_file = "3G.ico"
    elif tech == "2G":
        ico_file = "2G.ico"
    else:
        ico_file = ""
    return os.path.join(ICON_PATH, ico_file)

# ─────────────────────────────────────────────
# APLICAR ÍCONE NA PASTA  (Windows desktop.ini)
# ─────────────────────────────────────────────
def apply_folder_icon(folder_path: str, ico_path: str):
    if not ico_path or not os.path.exists(ico_path):
        print(f"  [AVISO] Ícone não encontrado: {ico_path}")
        return
    desktop_ini = os.path.join(folder_path, "desktop.ini")
    ini_content = (
        "[.ShellClassInfo]\n"
        f"IconResource={ico_path},0\n"
        "[ViewState]\n"
        "Mode=\n"
        "Vid=\n"
        "FolderType=Generic\n"
    )
    try:
        with open(desktop_ini, "w", encoding="utf-8") as f:
            f.write(ini_content)
        ctypes.windll.kernel32.SetFileAttributesW(desktop_ini, 0x02)
        ctypes.windll.kernel32.SetFileAttributesW(folder_path, 0x01 | 0x20)
    except Exception as e:
        print(f"  [AVISO] Não foi possível aplicar ícone em '{folder_path}': {e}")

# ─────────────────────────────────────────────
# ESTRUTURAS DE SUBPASTAS POR TECNOLOGIA
# ─────────────────────────────────────────────
def create_subfolders_4G(base: str):
    os.makedirs(os.path.join(base, "ROTA"), exist_ok=True)
    for s in ["S1", "S2", "S3"]:
        for sub in ["DL", "UL", "VOZ"]:
            os.makedirs(os.path.join(base, "STATICO", s, sub), exist_ok=True)

def create_subfolders_5G(base: str):
    for mode in ["NSA", "SA"]:
        os.makedirs(os.path.join(base, "ROTA", mode), exist_ok=True)
    for s in ["S1", "S2", "S3"]:
        for sub in ["DL", "UL"]:
            os.makedirs(os.path.join(base, "STATICO", "NSA", s, sub), exist_ok=True)
        for sub in ["DL", "UL", "EPSFB"]:
            os.makedirs(os.path.join(base, "STATICO", "SA", s, sub), exist_ok=True)

def create_subfolders_3G(base: str):
    os.makedirs(os.path.join(base, "ROTA - A"), exist_ok=True)
    os.makedirs(os.path.join(base, "ROTA - B"), exist_ok=True)

def create_subfolders_2G(base: str):
    os.makedirs(os.path.join(base, "ROTA"), exist_ok=True)

SUBFOLDER_CREATORS = {
    "4G": create_subfolders_4G,
    "5G": create_subfolders_5G,
    "3G": create_subfolders_3G,
    "2G": create_subfolders_2G,
}

# ─────────────────────────────────────────────
# PARSER DO RELATÓRIO
# ─────────────────────────────────────────────

# Log COM UF no final:  _EOLEN_SSV_MG-UVS_4G_700_PARACATU_MG_
LOG_RE_COM_UF = re.compile(
    r"_(TELEQUIPE|EOLEN|STEIN)_SSV_([A-Z0-9\-]+)_(4G|5G|3G|2G)_([A-Z0-9]+)_((?:[A-Z0-9]+_)+)([A-Z]{2})_",
    re.IGNORECASE
)

# Log SEM UF no final:  _TELEQUIPE_SSV_MG-VJG_4G_700_VESPASIANO__
LOG_RE_SEM_UF = re.compile(
    r"_(TELEQUIPE|EOLEN|STEIN)_SSV_([A-Z0-9\-]+)_(4G|5G|3G|2G)_([A-Z0-9]+)_((?:[A-Z0-9]+_?)+)__",
    re.IGNORECASE
)

# UF no cabeçalho do bloco: (CIDADE - MG) ou (CIDADE - SP )
UF_HEADER_RE = re.compile(r"\([^()]+\s*-\s*([A-Z]{2})\s*\)", re.IGNORECASE)

# Frequências: 4G:700/1800/2100  ou  5G: 3500
FREQ_RE = re.compile(r"(4G|5G|3G|2G)\s*:\s*([\d/A-Za-z]+)", re.IGNORECASE)


def parse_relatorio(path: str) -> list[dict]:
    """
    Para cada bloco do relatório:
      1. Seção 'Nome - LOGS' → instaladora, site, cidade
      2. UF → retirada do log se presente; senão do cabeçalho (CIDADE - UF)
      3. Linha '(Frequências):' → expande TODAS as frequências por tecnologia

    Retorna lista de dicts: { tech, freq, nome_pasta }
    """
    with open(path, "r", encoding="utf-8") as f:
        content = f.read()

    partes = re.split(r"--+\s*Nome\s*-\s*LOGS\s*--+", content)
    entradas = []

    for i in range(len(partes) - 1):
        corpo = partes[i]
        logs  = partes[i + 1]

        # ── 1. Tenta log COM UF ──────────────────────────────────────
        lm = LOG_RE_COM_UF.search(logs)
        if lm:
            instaladora = lm.group(1).upper()
            site        = lm.group(2).upper()
            cidade      = lm.group(5).rstrip("_").upper()
            uf          = lm.group(6).upper()
        else:
            # ── 2. Fallback: log SEM UF + UF do cabeçalho ───────────
            lm = LOG_RE_SEM_UF.search(logs)
            if not lm:
                continue
            instaladora = lm.group(1).upper()
            site        = lm.group(2).upper()
            cidade      = lm.group(5).rstrip("_").upper()
            uf_match    = UF_HEADER_RE.search(corpo)
            uf          = uf_match.group(1).upper() if uf_match else "XX"

        # ── 3. Frequências do bloco ──────────────────────────────────
        fb = re.search(
            r"\(Frequências\)\s*:(.*?)(?:OBS\.|NGSOLUTION|ERICSSON|$)",
            corpo, re.DOTALL | re.IGNORECASE
        )
        ft = fb.group(1) if fb else corpo

        for tm in FREQ_RE.finditer(ft):
            tech  = tm.group(1).upper()
            freqs = [f.strip().upper() for f in tm.group(2).split("/") if f.strip()]
            for freq in freqs:
                nome_pasta = f"{instaladora}_SSV_{site}_{tech}_{freq}_{cidade}_{uf}"
                entradas.append({
                    "tech":       tech,
                    "freq":       freq,
                    "nome_pasta": nome_pasta,
                })

    return entradas

# ─────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────
def main():
    global RELATORIO_PATH, ICON_PATH
    parser = argparse.ArgumentParser(add_help=False)
    parser.add_argument("--relatorio")
    parser.add_argument("--icon-path")
    args, _ = parser.parse_known_args()
    if args.relatorio:
        RELATORIO_PATH = args.relatorio
    if args.icon_path:
        ICON_PATH = args.icon_path

    root = tk.Tk()
    root.withdraw()

    print("=" * 65)
    print("  GERADOR DE PASTAS DT — SSV")
    print("=" * 65)

    if not os.path.exists(RELATORIO_PATH):
        msg = f"Relatório não encontrado:\n{RELATORIO_PATH}"
        print(f"\n[ERRO] {msg}")
        messagebox.showerror("PASTAS DT — ERRO", msg)
        return

    os.makedirs(STANDBY_PATH, exist_ok=True)
    existentes = set(os.listdir(STANDBY_PATH))

    print(f"\nLendo relatório: {RELATORIO_PATH}\n")
    entradas = parse_relatorio(RELATORIO_PATH)

    if not entradas:
        msg = "Nenhuma entrada encontrada no relatório.\nVerifique o arquivo RELATORIO_SSV_COMPLETO.txt."
        print(f"[AVISO] {msg}")
        messagebox.showwarning("PASTAS DT — AVISO", msg)
        return

    criadas    = []
    ja_existia = []

    for entry in entradas:
        nome = entry["nome_pasta"]
        dest = os.path.join(STANDBY_PATH, nome)

        if nome in existentes:
            ja_existia.append(nome)
            print(f"  [JÁ EXISTE]  {nome}")
        else:
            os.makedirs(dest, exist_ok=True)

            creator = SUBFOLDER_CREATORS.get(entry["tech"])
            if creator:
                creator(dest)

            ico = get_ico_path(entry["tech"], entry["freq"])
            apply_folder_icon(dest, ico)

            criadas.append(nome)
            print(f"  [CRIADA]     {nome}")

    # ── Resumo console ────────────────────────────────────────────────
    print()
    print("=" * 65)
    print("  RESUMO FINAL")
    print("=" * 65)
    print(f"  ✔ Pastas criadas:  {len(criadas)}")
    print(f"  ✦ Já existiam:     {len(ja_existia)}")

    # ── Resumo janela pop-up ──────────────────────────────────────────
    linhas = [
        f"✔ Pastas criadas:  {len(criadas)}",
        f"✦ Já existiam:     {len(ja_existia)}",
        "",
    ]
    if criadas:
        linhas.append("─── Criadas: ───────────────────────────────────────")
        for p in criadas:
            linhas.append(f"  {p}")
        linhas.append("")
    if ja_existia:
        linhas.append("─── Já existiam (ignoradas): ───────────────────────")
        for p in ja_existia:
            linhas.append(f"  {p}")

    messagebox.showinfo("PASTAS DT — Concluído", "\n".join(linhas))


if __name__ == "__main__":
    main()
