"""
DT 3.0 - Automação Drive Test
Organiza atividades, otimiza rota, atualiza Google Sheets,
gera relatórios e mapa em um único processo.
"""

import os
import sys
import re
import math
import unicodedata
import requests
import pandas as pd
import folium
import gspread
from pathlib import Path
from datetime import datetime
from google.oauth2.service_account import Credentials


# ============================================================
# CONFIGURAÇÃO
# ============================================================

def get_base_dir():
    if getattr(sys, 'frozen', False):
        return Path(os.path.dirname(sys.executable))
    return Path(os.path.abspath(__file__)).parent


BASE_DIR        = get_base_dir()
PASTA_NOVAS     = BASE_DIR / "novas atividades"
PASTA_OUT       = BASE_DIR / "out"
BASE_4G_PATH    = BASE_DIR / "Base_4G.xlsx"
CREDS_PATH      = BASE_DIR / "credentials.json"
SHEET_ID        = "1gPrzFOvPG6bF88H54ChXoyUmTWm84XU_ifFVO9X3rPE"
GITHUB_TOKEN_PATH  = BASE_DIR / "github_token.txt"
HOTEIS_PATH        = BASE_DIR / "HOTEIS.xlsx"          # fallback local (opcional)
HOTEIS_SHEET_ID    = "1Vw1cgSppfxezM8MGRv56E88pnD-im5ThNX2LkJTyDsk"
HOTEIS_GID         = 965678690                          # gid da aba correta

# Status da planilha
ST_CONCLUIDO    = "✓ Atividade concluída"
ST_IMPRODUTIVO  = "IMPRODUTIVO"
ST_DESLOCAMENTO = ">> EM DESLOCAMENTO"
ST_AGUARDANDO   = "Aguardando para deslocar"
ST_NOVA         = "Nova Atividade"
STATUS_FIXOS    = {ST_CONCLUIDO, ST_IMPRODUTIVO}

# Colunas da planilha (índice 0)
CI = {
    "DEMANDA":    0,  # A
    "INTEGRACAO": 1,  # B
    "SITE":       2,  # C
    "UF":         3,  # D
    "TEC":        4,  # E
    "LAT":        5,  # F
    "LON":        6,  # G
    "CIDADE":     7,  # H
    "TEC_4G":     8,  # I
    "TEC_5G":     9,  # J
    "STATUS":    10,  # K
    "CONCLUIDO": 11,  # L
    "HOTEL":     12,  # M
}

DATA_START_ROW = 3   # 1-based, primeira linha de dados

# Ordem de exibição das bandas 4G
ORDEM_BANDA_4G = [700, 850, 900, 1800, 2100, 2300, 2600]


# ============================================================
# UTILITÁRIOS
# ============================================================

def remove_acentos(texto):
    return unicodedata.normalize('NFKD', str(texto)).encode('ASCII', 'ignore').decode('ASCII')


def safe_float(value):
    try:
        if pd.isna(value):
            return 0.0
        return float(str(value).replace(',', '.'))
    except Exception:
        return 0.0


def float_close(a, b, eps=1e-4):
    try:
        return abs(float(a) - float(b)) <= eps
    except Exception:
        return False


def unique_preserve_order(seq):
    seen, out = set(), []
    for x in seq:
        if x not in seen:
            seen.add(x)
            out.append(str(x).strip())
    return out


def normalizar_colunas(df):
    df.columns = (
        df.columns.str.strip().str.upper()
        .str.normalize('NFKD')
        .str.encode('ascii', errors='ignore')
        .str.decode('utf-8')
    )
    return df


def formatar_site(site):
    s = str(site).strip().upper()
    if len(s) >= 5:
        return f"{s[:2]}-{s[-3:]}"
    return s


# ============================================================
# NORMALIZAÇÃO DE TECNOLOGIA (TEC)
# ============================================================

# Tabela de substituição: chave = valor bruto normalizado → valor padronizado
# Ordem importa: mais específico primeiro (ex: LTE|NR antes de LTE)
_MAPA_TEC = [
    # 4 tecnologias
    (r"^2G[|/]3G[|/]LTE[|/]NR$",  "2G|3G|4G|5G"),
    # 3 tecnologias
    (r"^3G[|/]LTE[|/]NR$",         "3G|4G|5G"),
    (r"^2G[|/]3G[|/]LTE$",         "2G|3G|4G"),
    (r"^2G[|/]3G[|/]NR$",          "2G|3G|5G"),
    # 2 tecnologias
    (r"^LTE[|/]NR$",               "4G|5G"),
    (r"^3G[|/]LTE$",               "3G|4G"),
    (r"^3G[|/]NR$",                "3G|5G"),
    (r"^2G[|/]LTE$",               "2G|4G"),
    # 1 tecnologia
    (r"^LTE$",                     "4G"),
    (r"^NR$",                      "5G"),
]

# Valores já no padrão correto — não tocar
_TEC_PADRAO = {"2G", "3G", "4G", "5G",
               "2G|3G", "2G|4G", "2G|5G",
               "3G|4G", "3G|5G", "4G|5G",
               "2G|3G|4G", "2G|3G|5G", "2G|4G|5G", "3G|4G|5G",
               "2G|3G|4G|5G"}


def normalizar_tec(valor):
    """
    Converte valores de TEC fora do padrão para o padrão DT 3.0.

    Exemplos:
      'LTE'          → '4G'
      'LTE|NR'       → '4G|5G'
      '3G|LTE|NR'   → '3G|4G|5G'
      '2G|3G|LTE|NR' → '2G|3G|4G|5G'
      '4G'           → '4G'   (já correto, sem alteração)
      ''             → ''     (vazio, sem alteração)
    """
    if not valor:
        return valor
    s = str(valor).strip().upper()
    if not s or s in ("NAN", "NONE"):
        return ""
    # Já está no padrão — retorna sem tocar
    if s in _TEC_PADRAO:
        return s
    # Normalizar separadores: / → |  e remover espaços
    s_norm = re.sub(r"\s*[/]\s*", "|", s)
    for padrao, substituto in _MAPA_TEC:
        if re.match(padrao, s_norm, re.IGNORECASE):
            return substituto
    # Não reconhecido — retorna o original sem alterar (nunca inventa)
    return valor.strip()


# ============================================================
# NORMALIZAÇÃO DE FREQUÊNCIA
# ============================================================

def _ordenar_nums_4g(nums_str):
    """
    Recebe lista de strings tipo ['2600', '700', '1800RS']
    Retorna reordenada por ORDEM_BANDA_4G (700, 850, 900, 1800, 2100, 2300, 2600).
    Sufixos como RS são preservados.
    """
    def chave(x):
        num = int(re.search(r'\d+', x).group()) if re.search(r'\d+', x) else 9999
        return ORDEM_BANDA_4G.index(num) if num in ORDEM_BANDA_4G else 99
    return sorted(nums_str, key=chave)


def _reordenar_bloco_4g(bloco):
    """
    Recebe '4G:2600/1800/700' ou '4G:700/1800/2100/2300/2600'
    Retorna '4G:700/1800/2600' (ordenado).
    """
    if ':' not in bloco:
        return bloco
    tec, nums_raw = bloco.split(':', 1)
    nums = [n.strip() for n in nums_raw.split('/') if n.strip()]
    nums_ordenados = _ordenar_nums_4g(nums)
    return f"{tec}:{'/'.join(nums_ordenados)}"


def normalizar_frequencia(freq_raw):
    """
    Normaliza qualquer formato de frequência para (freq_234g, freq_5g).

    Exemplos:
      'N3500|L2600/L2100/L1800/L700'          → ('4G: 700/1800/2100/2600', '5G: 3500')
      '3G:850|4G:700/1800/2100/2600|5G:3500'  → ('3G:850|4G:700/1800/2100/2600', '5G:3500')
      '4G:2600/1800/700|5G:3500'               → ('4G:700/1800/2600', '5G:3500')
      '4G:700/2600RS'                          → ('4G:700/2600RS', '')
    """
    if not isinstance(freq_raw, str) or not freq_raw.strip():
        return "", ""

    freq = freq_raw.strip()

    # ---- Formato antigo: NR?NNNN e LNNNN sem 'G:' ----
    if re.search(r'\b(?:NR?|L)\d{3,4}', freq, re.IGNORECASE) and 'G:' not in freq:
        freq_5g = ""
        nums_4g = []

        for m in re.finditer(r'NR?(\d{3,4})', freq, re.IGNORECASE):
            freq_5g = f"5G: {m.group(1)}"

        for m in re.finditer(r'L(\d{3,4})', freq, re.IGNORECASE):
            nums_4g.append(m.group(1))

        if nums_4g:
            freq_4g = "4G: " + "/".join(_ordenar_nums_4g(nums_4g))
        else:
            freq_4g = ""

        return freq_4g, freq_5g

    # ---- Formato já limpo: separar 5G e reordenar 4G ----
    partes = freq.split("|")
    ate_4g, freq_5g = [], ""

    for p in partes:
        p = p.strip()
        if not p:
            continue
        if p.upper().startswith("5G"):
            freq_5g = p
        elif re.match(r'^4G:', p, re.IGNORECASE):
            ate_4g.append(_reordenar_bloco_4g(p))
        else:
            ate_4g.append(p)

    return "|".join(ate_4g), freq_5g


# ============================================================
# DISTÂNCIA / OTIMIZAÇÃO DE ROTA
# ============================================================

def haversine(lat1, lon1, lat2, lon2):
    R = 6371
    dLat = math.radians(lat2 - lat1)
    dLon = math.radians(lon2 - lon1)
    a = (math.sin(dLat / 2) ** 2
         + math.cos(math.radians(lat1))
         * math.cos(math.radians(lat2))
         * math.sin(dLon / 2) ** 2)
    return R * 2 * math.atan2(math.sqrt(a), math.sqrt(1 - a))


def distancia_total(df, lat0, lon0):
    total, la, lo = 0, lat0, lon0
    for _, r in df.iterrows():
        total += haversine(la, lo, r["LAT"], r["LONG"])
        la, lo = r["LAT"], r["LONG"]
    return total


def vizinho_mais_proximo(df, lat0, lon0):
    pontos = df.copy()
    rota, la, lo = [], lat0, lon0
    while len(pontos):
        pontos["_d"] = pontos.apply(
            lambda r: haversine(la, lo, r["LAT"], r["LONG"]), axis=1
        )
        idx = pontos["_d"].idxmin()
        prox = pontos.loc[idx]
        rota.append(prox)
        la, lo = prox["LAT"], prox["LONG"]
        pontos = pontos.drop(idx)
    return pd.DataFrame(rota).drop(columns=["_d"], errors="ignore")


def _rota_para_coords(df):
    """Converte DataFrame de rota em lista de (lat, lon) para cálculos rápidos."""
    return [(r["LAT"], r["LONG"]) for _, r in df.iterrows()]


def _dist_coords(coords, lat0, lon0):
    """Calcula distância total de uma lista de coords."""
    total, la, lo = 0, lat0, lon0
    for lat, lon in coords:
        total += haversine(la, lo, lat, lon)
        la, lo = lat, lon
    return total


def two_opt(df, lat0, lon0):
    """2-OPT clássico — inverte segmentos."""
    rota = df.reset_index(drop=True)
    coords = _rota_para_coords(rota)
    melhor_d = _dist_coords(coords, lat0, lon0)
    melhorou = True
    while melhorou:
        melhorou = False
        for i in range(len(coords) - 1):
            for j in range(i + 2, len(coords)):
                nova = coords[:i+1] + coords[i+1:j+1][::-1] + coords[j+1:]
                d = _dist_coords(nova, lat0, lon0)
                if d < melhor_d - 1e-6:
                    coords = nova
                    melhor_d = d
                    melhorou = True
    # Reconstruir DataFrame na nova ordem
    idx_map = {(r["LAT"], r["LONG"]): i for i, (_, r) in enumerate(rota.iterrows())}
    nova_ordem = []
    for lat, lon in coords:
        nova_ordem.append(idx_map[(lat, lon)])
    return rota.iloc[nova_ordem].reset_index(drop=True), melhor_d


def three_opt(df, lat0, lon0):
    """
    3-OPT: testa todas as combinações de 3 arestas e aplica a melhor
    reconexão possível (8 variantes por tripla de segmentos).
    Resolve cruzamentos que o 2-OPT não consegue desfazer.
    """
    rota = df.reset_index(drop=True)
    coords = _rota_para_coords(rota)
    n = len(coords)
    melhor_d = _dist_coords(coords, lat0, lon0)
    melhorou = True

    # Pré-calcular todas as distâncias entre pares (cache)
    def d(a, b):
        return haversine(a[0], a[1], b[0], b[1])

    # Ponto de origem como coordenada
    origem = (lat0, lon0)

    def dist_total(seq):
        return _dist_coords(seq, lat0, lon0)

    while melhorou:
        melhorou = False
        for i in range(n - 2):
            for j in range(i + 1, n - 1):
                for k in range(j + 1, n):
                    # Segmentos: [0..i] [i+1..j] [j+1..k] [k+1..n-1]
                    A = coords[:i+1]
                    B = coords[i+1:j+1]
                    C = coords[j+1:k+1]
                    D = coords[k+1:]

                    # 7 reconexões possíveis além da original
                    candidatos = [
                        A + B[::-1] + C       + D,   # 2-opt AB
                        A + B       + C[::-1] + D,   # 2-opt BC
                        A + B[::-1] + C[::-1] + D,   # 2-opt AB+BC
                        A + C       + B       + D,   # 3-opt move B após C
                        A + C       + B[::-1] + D,   # 3-opt + inv B
                        A + C[::-1] + B       + D,   # 3-opt + inv C
                        A + C[::-1] + B[::-1] + D,   # 3-opt + inv ambos
                    ]

                    for cand in candidatos:
                        d_cand = dist_total(cand)
                        if d_cand < melhor_d - 1e-6:
                            coords = cand
                            melhor_d = d_cand
                            melhorou = True

    # Reconstruir DataFrame
    idx_map = {}
    for i, (_, r) in enumerate(rota.iterrows()):
        idx_map[(r["LAT"], r["LONG"])] = i
    nova_ordem = [idx_map[(lat, lon)] for lat, lon in coords]
    return rota.iloc[nova_ordem].reset_index(drop=True), melhor_d


def or_opt(df, lat0, lon0, tamanho_seg=None):
    """
    Or-Opt: move segmentos de 1, 2 ou 3 cidades para outra posição da rota.
    Resolve o caso clássico de cidades 'puladas' que ficam no caminho.
    """
    rota = df.reset_index(drop=True)
    coords = _rota_para_coords(rota)
    n = len(coords)
    melhor_d = _dist_coords(coords, lat0, lon0)

    tamanhos = tamanho_seg if tamanho_seg else [1, 2, 3]
    melhorou = True

    while melhorou:
        melhorou = False
        for seg_len in tamanhos:
            for i in range(n - seg_len + 1):
                segmento = coords[i:i + seg_len]
                resto = coords[:i] + coords[i + seg_len:]

                for j in range(len(resto) + 1):
                    if j == i:
                        continue
                    nova = resto[:j] + segmento + resto[j:]
                    d_nova = _dist_coords(nova, lat0, lon0)
                    if d_nova < melhor_d - 1e-6:
                        coords = nova
                        melhor_d = d_nova
                        melhorou = True

    # Reconstruir DataFrame
    idx_map = {}
    for i, (_, r) in enumerate(rota.iterrows()):
        idx_map[(r["LAT"], r["LONG"])] = i
    nova_ordem = [idx_map[(lat, lon)] for lat, lon in coords]
    return rota.iloc[nova_ordem].reset_index(drop=True), melhor_d


def otimizar_rota(df_pool, lat0, lon0):
    """
    Pipeline completo de otimização:
    1. Múltiplos pontos de partida (vizinho mais próximo a partir de cada cidade)
    2. 2-OPT em cada candidato
    3. Or-Opt (move segmentos de 1-3 cidades)
    4. 3-OPT (resolve cruzamentos complexos)
    5. Or-Opt final para polish
    """
    n = len(df_pool)
    print(f"   → Gerando candidatos ({n} pontos de partida alternativos)...")

    melhor_rota = None
    melhor_d = float("inf")

    # Candidato 1: ponto de partida real
    pontos_inicio = [(lat0, lon0)]

    # Candidatos adicionais: cada cidade como ponto de partida fictício
    for _, r in df_pool.iterrows():
        pontos_inicio.append((r["LAT"], r["LONG"]))

    for i, (la, lo) in enumerate(pontos_inicio):
        rota_c = vizinho_mais_proximo(df_pool, la, lo)
        d_c = distancia_total(rota_c, lat0, lon0)  # distância sempre medida da origem real
        if d_c < melhor_d:
            melhor_d = d_c
            melhor_rota = rota_c

    d_pos_nn = melhor_d
    print(f"   📏 Melhor vizinho mais próximo: {d_pos_nn:.1f} km")

    print("   → 2-OPT...")
    melhor_rota, d_2opt = two_opt(melhor_rota, lat0, lon0)
    print(f"   📏 Após 2-OPT  : {d_2opt:.1f} km  (Δ {d_pos_nn - d_2opt:.1f} km)")

    print("   → Or-Opt (segmentos 1-3)...")
    melhor_rota, d_oropt = or_opt(melhor_rota, lat0, lon0)
    print(f"   📏 Após Or-Opt : {d_oropt:.1f} km  (Δ {d_2opt - d_oropt:.1f} km)")

    print("   → 3-OPT...")
    melhor_rota, d_3opt = three_opt(melhor_rota, lat0, lon0)
    print(f"   📏 Após 3-OPT  : {d_3opt:.1f} km  (Δ {d_oropt - d_3opt:.1f} km)")

    print("   → Or-Opt final (polish)...")
    melhor_rota, d_final = or_opt(melhor_rota, lat0, lon0)

    print(f"\nDistância inicial : {d_pos_nn:.1f} km")
    print(f"Distância final   : {d_final:.1f} km")
    print(f"Economia total    : {d_pos_nn - d_final:.1f} km ({(d_pos_nn - d_final) / d_pos_nn * 100:.1f}%)")

    return melhor_rota


# ============================================================
# GOOGLE SHEETS
# ============================================================

def conectar_sheets():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = Credentials.from_service_account_file(str(CREDS_PATH), scopes=scopes)
    return gspread.authorize(creds)


def obter_aba_vigente(sheet):
    meses = [
        "JANEIRO", "FEVEREIRO", "MARÇO", "ABRIL", "MAIO", "JUNHO",
        "JULHO", "AGOSTO", "SETEMBRO", "OUTUBRO", "NOVEMBRO", "DEZEMBRO",
    ]
    now = datetime.now()
    nome = f"{meses[now.month - 1]}/{now.year}"

    try:
        ws = sheet.worksheet(nome)
        print(f"   Aba encontrada: {nome}")
        return ws
    except gspread.exceptions.WorksheetNotFound:
        print(f"   ⚠️  Aba '{nome}' não existe. Criando...")
        abas = sheet.worksheets()
        if abas:
            sheet.duplicate_sheet(abas[-1].id, new_sheet_name=nome)
            ws = sheet.worksheet(nome)
            # Limpar dados mantendo cabeçalhos (linhas 1 e 2)
            all_vals = ws.get_all_values()
            if len(all_vals) >= DATA_START_ROW:
                last = len(all_vals)
                ws.batch_clear([f"A{DATA_START_ROW}:M{last}"])
        else:
            ws = sheet.add_worksheet(nome, rows=200, cols=20)
        return ws


COLUNAS_SHEETS = [
    "ROW_SHEET", "DEMANDA", "INTEGRAÇÃO", "SITE", "UF", "TEC",
    "LAT", "LONG", "CIDADE", "2G|3G|4G", "5G", "STATUS", "CONCLUIDO", "HOTEL",
]


def ler_atividades_sheets(ws):
    """Lê todas as linhas de dados e retorna DataFrame com colunas sempre definidas."""
    dados = ws.get_all_values()
    linhas = dados[DATA_START_ROW - 1:]   # 0-indexed
    registros = []

    for i, linha in enumerate(linhas):
        while len(linha) <= max(CI.values()):
            linha.append("")

        # Ignorar linhas completamente vazias
        if not any(linha[:8]):
            continue

        try:
            lat = float(str(linha[CI["LAT"]]).replace(',', '.')) if linha[CI["LAT"]] else None
            lon = float(str(linha[CI["LON"]]).replace(',', '.')) if linha[CI["LON"]] else None
        except Exception:
            lat = lon = None

        registros.append({
            "ROW_SHEET": DATA_START_ROW + i,
            "DEMANDA":    linha[CI["DEMANDA"]],
            "INTEGRAÇÃO": linha[CI["INTEGRACAO"]],
            "SITE":       linha[CI["SITE"]],
            "UF":         linha[CI["UF"]],
            "TEC":        normalizar_tec(linha[CI["TEC"]]),
            "LAT":        lat,
            "LONG":       lon,
            "CIDADE":     linha[CI["CIDADE"]],
            "2G|3G|4G":  linha[CI["TEC_4G"]],
            "5G":         linha[CI["TEC_5G"]],
            "STATUS":     linha[CI["STATUS"]].strip(),
            "CONCLUIDO":  linha[CI["CONCLUIDO"]],
            "HOTEL":      linha[CI["HOTEL"]],
        })

    if not registros:
        # Retorna DataFrame vazio mas com todas as colunas definidas
        return pd.DataFrame(columns=COLUNAS_SHEETS)

    return pd.DataFrame(registros)


def atualizar_sheets(ws, df_fixas, df_rota, sites_originais, df_aguardando=None):
    """
    Grava a planilha atualizada:
    - Linhas fixas (concluídas/improdutivas) permanecem na ordem original
    - Demais linhas vêm em seguida na ordem otimizada
    - "Aguardando para deslocar" vai ao final, separado por linha em branco
    - "Nova Atividade" é limpo de atividades antigas; adicionado às novas
    """
    print("   Montando linhas...")

    def _coord(v):
        """Converte coordenada para string com vírgula (padrão BR do Sheets)."""
        try:
            return str(float(v)).replace(".", ",")
        except Exception:
            return str(v) if v else ""

    def formatar_linha(row_dict, status_override=None):
        status = status_override if status_override is not None else row_dict.get("STATUS", "")
        return [
            str(row_dict.get("DEMANDA",    "")),
            str(row_dict.get("INTEGRAÇÃO", "")),
            str(row_dict.get("SITE",       "")),
            str(row_dict.get("UF",         "")),
            str(row_dict.get("TEC",        "")),
            _coord(row_dict.get("LAT",     "")),
            _coord(row_dict.get("LONG",    "")),
            str(row_dict.get("CIDADE",     "")),
            str(row_dict.get("2G|3G|4G",   "")),
            str(row_dict.get("5G",         "")),
            str(status),
            str(row_dict.get("CONCLUIDO",  "") or ""),
            str(row_dict.get("HOTEL",      "") or ""),
        ]

    todas_linhas = []

    # 1. Atividades fixas (ordem original)
    for _, row in df_fixas.iterrows():
        todas_linhas.append(formatar_linha(row.to_dict()))

    # 2. Atividades otimizadas
    for _, row in df_rota.iterrows():
        row_d = row.to_dict()
        site = str(row_d.get("SITE", "")).upper()

        if site in sites_originais:
            # Já existia: limpa "Nova Atividade", mantém status atual
            status = row_d.get("STATUS", "")
            if status == ST_NOVA:
                status = ""
        else:
            # Nova: marca com "Nova Atividade"
            status = ST_NOVA

        todas_linhas.append(formatar_linha(row_d, status_override=status))

    # 3. Adicionar "Aguardando para deslocar" no final (separado por linha em branco)
    if df_aguardando is not None and not df_aguardando.empty:
        todas_linhas.append([""] * 13)   # linha em branco separadora
        for _, row in df_aguardando.iterrows():
            todas_linhas.append(formatar_linha(row.to_dict()))

    # 4. Gravar
    end_row = DATA_START_ROW + len(todas_linhas) - 1
    range_ref = f"A{DATA_START_ROW}:M{end_row}"
    ws.update(values=todas_linhas, range_name=range_ref, value_input_option="USER_ENTERED")

    # 5. Limpar apenas linhas excedentes vazias (nunca apaga atividades reais)
    # Lê de volta para saber quantas linhas realmente têm conteúdo após o que gravamos
    dados_atuais = ws.get_all_values()
    ultima_linha_com_dados = len(dados_atuais)
    if ultima_linha_com_dados > end_row:
        # Só limpa se as linhas de fato existem e estão além do nosso bloco
        ws.batch_clear([f"A{end_row + 1}:M{ultima_linha_com_dados}"])

    print(f"   ✅ {len(todas_linhas)} linhas gravadas no Google Sheets.")


# ============================================================
# LOCALIZAÇÃO INICIAL
# ============================================================

def _reverse_geocode(lat, lon):
    try:
        url = (
            f"https://nominatim.openstreetmap.org/reverse"
            f"?format=json&lat={lat}&lon={lon}&zoom=10"
        )
        r = requests.get(url, headers={"User-Agent": "DT3.0"}, timeout=6)
        addr = r.json().get("address", {})
        return (
            addr.get("city") or addr.get("town")
            or addr.get("village") or addr.get("state") or "Local"
        )
    except Exception:
        return "Local"


def determinar_ponto_inicio(df_sheets):
    """
    Determina ponto de partida baseado na última atividade concluída
    ou deslocamento ativo, com opção de entrada manual.
    """
    # Planilha pode estar vazia (aba nova) — protege o acesso à coluna STATUS
    if not df_sheets.empty and "STATUS" in df_sheets.columns:
        concluidas = df_sheets[df_sheets["STATUS"] == ST_CONCLUIDO]

        if not concluidas.empty:
            ultima = concluidas.iloc[-1]
            cidade = ultima["CIDADE"]
            lat_u = ultima["LAT"]
            lon_u = ultima["LONG"]
            print(f"\n📍 Última atividade concluída: {ultima['SITE']} — {cidade}")

            if lat_u and lon_u:
                resp = input(f"   Você ainda está em {cidade}? [S/N]: ").strip().upper()
                if resp == "S":
                    return float(lat_u), float(lon_u), str(cidade)

        # Verificar deslocamento ativo
        desloc = df_sheets[
            df_sheets["STATUS"].str.contains(
                "EM DESLOCAMENTO|Aguardando", na=False, case=False
            )
        ]
        if not desloc.empty:
            prox = desloc.iloc[0]
            print(f"\n   Próxima atividade pendente: {prox['SITE']} — {prox['CIDADE']}")
    else:
        print("   Planilha sem atividades anteriores.")

    print("\n   Informe sua localização atual no formato: -00.000000 -00.000000")
    print("   (ou ENTER para detectar automaticamente por IP)")
    entrada = input("   >>> ").strip()

    if entrada:
        partes = entrada.split()
        if len(partes) >= 2:
            try:
                lat = float(partes[0].replace(",", "."))
                lon = float(partes[1].replace(",", "."))
                cidade = _reverse_geocode(lat, lon)
                print(f"   Localização manual: {cidade} ({lat}, {lon})")
                return lat, lon, cidade
            except Exception:
                print("   ⚠️  Entrada inválida. Tentando detecção por IP...")

    # Detecção por IP
    try:
        r = requests.get("http://ip-api.com/json/", timeout=5)
        d = r.json()
        if d.get("status") == "success":
            lat = d["lat"]
            lon = d["lon"]
            cidade = d.get("city") or d.get("regionName") or "Local detectado"
            print(f"   Localização detectada: {cidade} ({lat}, {lon})")
            return lat, lon, cidade
    except Exception:
        pass

    print("   ⚠️  Usando Belo Horizonte como padrão.")
    return -19.95, -43.93, "Belo Horizonte"


# ============================================================
# LEITURA DAS NOVAS ATIVIDADES
# ============================================================

# Mapeamento de campos internos para variações de nome de coluna
# Adicione aqui novos sinônimos se seu chefe mudar algum nome
_MAPA_COLUNAS = {
    "SITE":        ["SITE", "SITES", "SITE_ID", "COD_SITE", "NOME_SITE", "CODIGO"],
    "CIDADE":      ["CIDADE", "MUNICIPIO", "CITY", "LOCALIDADE"],
    "UF":          ["UF", "ESTADO", "STATE", "REGIONAL", "UF_SITE"],
    "LATITUDE":    ["LATITUDE", "LAT", "LATIT"],       # espaços removidos na normalização
    "LONGITUDE":   ["LONGITUDE", "LON", "LONG", "LONGIT"],
    "FREQUENCIA":  ["FREQUENCIA", "FREQUÊNCIA", "FREQ", "FREQUENCIAS", "FREQUENCY"],
    "VENDOR":      ["VENDOR", "EQUIPE_RF", "EMPRESA_RF", "FORNECEDOR", "INTEGRADOR_RF"],
    "INTEGRADORA": ["INTEGRADORA", "DEMANDANTE", "CLIENTE", "OPERADORA"],
    "TECNOLOGIA":  ["TECNOLOGIA", "TEC", "TECNOLOGIAS", "TECH", "TECHNOLOGY"],
}


def _resolver_colunas(df_raw):
    """
    Normaliza os nomes de colunas do DataFrame recebido e
    retorna um dict {campo_interno: nome_original_no_arquivo}.
    """
    # Normalizar: strip + upper + sem acento
    cols_norm = {}
    for col_orig in df_raw.columns:
        col_n = (unicodedata.normalize("NFKD", str(col_orig).strip().upper())
                 .encode("ascii", errors="ignore").decode("utf-8"))
        cols_norm[col_n] = col_orig

    mapeado = {}
    nao_encontrado = []
    for campo, candidatos in _MAPA_COLUNAS.items():
        col_orig = next((cols_norm[c] for c in candidatos if c in cols_norm), None)
        if col_orig:
            mapeado[campo] = col_orig
        else:
            nao_encontrado.append(campo)

    if nao_encontrado:
        print(f"   ⚠️  Colunas nao encontradas: {nao_encontrado}")
        print(f"      Colunas disponiveis: {list(df_raw.columns)}")

    return mapeado


def _limpar_uf(val):
    """Retorna UF válida (2 letras) ou string vazia."""
    s = str(val).strip().upper()
    if s in ("", "NAN", "NONE", "0") or not s.isalpha():
        return ""
    return s if len(s) == 2 else ""


def _limpar_cidade(val):
    """Retorna cidade em maiúsculas ou string vazia."""
    s = str(val).strip().upper()
    return "" if s in ("", "NAN", "NONE") else s


def _formatar_site_robusto(site_raw):
    """
    Formata código de site para padrão UF-XXX.
    Trata sufixos como _B, _A, _C (ex: MGGCR_B → MG-GCR).
    """
    s = str(site_raw).strip().upper()
    # Remover sufixo _X (letra única no final)
    s = re.sub(r"_[A-Z]$", "", s)
    if len(s) >= 5:
        return f"{s[:2]}-{s[-3:]}"
    return s


# ─────────────────────────────────────────────────────────────────────────────
# AJUSTE CIRÚRGICO 1 — _buscar_na_base4g
# Nova função auxiliar centralizada: busca match na Base_4G por lat/lon,
# com tolerância configurável (padrão 0.002 ≈ ~220m).
# Usada tanto em processar_arquivo quanto em gerar_relatorio.
# ─────────────────────────────────────────────────────────────────────────────
def _buscar_na_base4g(df_base, lat, lon, tolerancia=0.002):
    """
    Retorna o subconjunto de df_base cujas coordenadas estão dentro de
    `tolerancia` graus de (lat, lon).  Tenta primeiro com cidade (se
    disponível no chamador), mas aqui retorna tudo — o filtro por cidade
    fica a cargo do chamador quando necessário.
    """
    if df_base is None or df_base.empty:
        return pd.DataFrame()
    mask = (
        (abs(df_base["LATITUDE"]  - lat) <= tolerancia) &
        (abs(df_base["LONGITUDE"] - lon) <= tolerancia)
    )
    return df_base[mask]


def processar_arquivo(df_raw, nome_arquivo="", df_base=None):
    """
    Converte um DataFrame bruto de qualquer planilha de atividades
    para o formato interno do DT 3.0.

    Proteções implementadas:
    - Nomes de colunas com espaço, case, acentos, sinônimos
    - UF = 0 ou inválida (marcada vazia para busca na Base_4G)
    - Cidade vazia/NaN (marcada vazia para busca por coordenadas)
    - Sufixos no código do site (_B, _A, etc.)
    - Tecnologia ausente (NaN → string vazia, não bloqueia)
    - Linhas sem coordenadas são descartadas com aviso
    - Linhas sem SITE são descartadas com aviso
    """
    prefixo = f"   [{nome_arquivo}]" if nome_arquivo else "  "

    # ── 1. Resolver colunas ─────────────────────────────────────────────
    col = _resolver_colunas(df_raw)

    def _get(campo, default=""):
        """Retorna valor da coluna ou default se coluna não existir."""
        if campo not in col:
            return default
        v = df_raw[col[campo]]
        return v

    registros = []
    descartados = 0

    for idx, row in df_raw.iterrows():
        linha_num = idx + 2   # +2 porque idx começa em 0 e linha 1 é cabeçalho

        # ── SITE — obrigatório ──────────────────────────────────────────
        site_raw = str(row[col["SITE"]]).strip() if "SITE" in col else ""
        if not site_raw or site_raw.upper() in ("NAN", "NONE", ""):
            print(f"{prefixo} ⚠️  Linha {linha_num}: SITE vazio — ignorada")
            descartados += 1
            continue
        site = _formatar_site_robusto(site_raw)

        # ── COORDENADAS — obrigatórias ──────────────────────────────────
        lat  = safe_float(row[col["LATITUDE"]])  if "LATITUDE"  in col else 0.0
        lon  = safe_float(row[col["LONGITUDE"]]) if "LONGITUDE" in col else 0.0
        if lat == 0.0 and lon == 0.0:
            print(f"{prefixo} ⚠️  Linha {linha_num} ({site}): coordenadas zeradas — ignorada")
            descartados += 1
            continue

        # ── UF ──────────────────────────────────────────────────────────
        uf_raw = row[col["UF"]] if "UF" in col else ""
        uf = _limpar_uf(uf_raw)
        if not uf:
            print(f"{prefixo} ℹ️  {site}: UF ausente/invalida ({repr(uf_raw)}) — sera buscada na Base_4G")

        # ── CIDADE ──────────────────────────────────────────────────────
        cidade_raw = row[col["CIDADE"]] if "CIDADE" in col else ""
        cidade = _limpar_cidade(cidade_raw)
        if not cidade:
            print(f"{prefixo} ℹ️  {site}: CIDADE vazia — sera buscada por coordenadas")

        # ── TECNOLOGIA — opcional ───────────────────────────────────────
        tec_raw = row[col["TECNOLOGIA"]] if "TECNOLOGIA" in col else ""
        tec = "" if pd.isna(tec_raw) else normalizar_tec(str(tec_raw).strip())

        # ── FREQUÊNCIA ──────────────────────────────────────────────────
        freq_raw = row[col["FREQUENCIA"]] if "FREQUENCIA" in col else ""
        if pd.isna(freq_raw) or str(freq_raw).strip().upper() in ("NAN", "NONE", ""):
            freq_234, freq_5g = "", ""
        else:
            freq_234, freq_5g = normalizar_frequencia(str(freq_raw).strip())

        # ── DEMANDA / VENDOR ────────────────────────────────────────────
        def _s(campo):
            if campo not in col: return ""
            v = row[col[campo]]
            return "" if pd.isna(v) else str(v).strip()

        registros.append({
            "DEMANDA":    _s("INTEGRADORA"),
            "INTEGRAÇÃO": _s("VENDOR"),
            "SITE":       site,
            "UF":         uf,
            "TEC":        tec,
            "LAT":        lat,
            "LONG":       lon,
            "CIDADE":     cidade,
            "2G|3G|4G":  freq_234,
            "5G":         freq_5g,
            "STATUS":     ST_NOVA,
            "CONCLUIDO":  "",
            "HOTEL":      "",
        })

    if descartados:
        print(f"{prefixo} ⚠️  {descartados} linha(s) descartada(s).")

    df_out = pd.DataFrame(registros)

    # ── Recuperar campos ausentes via Base_4G (por lat/lon) ─────────────
    # AJUSTE CIRÚRGICO 1b: além de CIDADE e UF, complementa TEC e FREQUÊNCIA
    # quando ausentes. STATUS, CONCLUIDO e HOTEL nunca são tocados aqui.
    if df_base is not None and not df_out.empty:
        for idx, row in df_out.iterrows():
            precisa_cidade = not row["CIDADE"]
            precisa_uf     = not row["UF"]
            precisa_tec    = not row["TEC"]
            precisa_4g     = not row["2G|3G|4G"]
            precisa_5g     = not row["5G"]

            nada_precisa = not any([precisa_cidade, precisa_uf,
                                    precisa_tec, precisa_4g, precisa_5g])
            if nada_precisa or row["LAT"] == 0.0:
                continue

            try:
                # Usar função auxiliar centralizada
                matches = _buscar_na_base4g(df_base, row["LAT"], row["LONG"])

                if matches.empty:
                    continue

                m = matches.iloc[0]

                if precisa_cidade:
                    val = str(m.get("CIDADE", "")).strip().upper()
                    if val and val not in ("", "NAN", "NONE"):
                        df_out.at[idx, "CIDADE"] = val
                        print(f"{prefixo} ℹ️  {row['SITE']}: CIDADE recuperada da Base_4G: {val}")

                if precisa_uf:
                    val = str(m.get("[P]UF", "")).strip().upper()
                    if val and val not in ("", "NAN", "NONE") and len(val) == 2:
                        df_out.at[idx, "UF"] = val
                        print(f"{prefixo} ℹ️  {row['SITE']}: UF recuperada da Base_4G: {val}")

                # ── AJUSTE: complementar TEC quando ausente ──────────────
                if precisa_tec:
                    # Tentar coluna TEC ou TECNOLOGIA na Base_4G
                    for col_tec in ("TEC", "TECNOLOGIA", "TECHNOLOGY"):
                        if col_tec in matches.columns:
                            val = str(m.get(col_tec, "")).strip()
                            if val and val.upper() not in ("", "NAN", "NONE"):
                                df_out.at[idx, "TEC"] = val
                                print(f"{prefixo} ℹ️  {row['SITE']}: TEC recuperada da Base_4G: {val}")
                                break

                # ── AJUSTE: complementar frequências quando ausentes ─────
                if precisa_4g or precisa_5g:
                    for col_freq in ("FREQUENCIA", "FREQUÊNCIA", "FREQ", "FREQUENCY"):
                        if col_freq in matches.columns:
                            freq_raw_base = str(m.get(col_freq, "")).strip()
                            if freq_raw_base and freq_raw_base.upper() not in ("", "NAN", "NONE"):
                                f4g, f5g = normalizar_frequencia(freq_raw_base)
                                if precisa_4g and f4g:
                                    df_out.at[idx, "2G|3G|4G"] = f4g
                                    print(f"{prefixo} ℹ️  {row['SITE']}: 2G|3G|4G recuperada da Base_4G: {f4g}")
                                if precisa_5g and f5g:
                                    df_out.at[idx, "5G"] = f5g
                                    print(f"{prefixo} ℹ️  {row['SITE']}: 5G recuperada da Base_4G: {f5g}")
                                break

            except Exception as e:
                print(f"{prefixo} ⚠️  Erro ao buscar Base_4G para {row['SITE']}: {e}")

        # Avisar sobre campos ainda ausentes após todas as tentativas
        for _, row in df_out.iterrows():
            if not row["CIDADE"]:
                print(f"{prefixo} ⚠️  {row['SITE']}: CIDADE nao encontrada em nenhuma fonte.")
            if not row["UF"]:
                print(f"{prefixo} ⚠️  {row['SITE']}: UF nao encontrada — sera solicitada no relatorio.")

    return df_out


# Manter alias para compatibilidade com chamadas existentes
def processar_felipe(df_raw, df_base=None):
    return processar_arquivo(df_raw, df_base=df_base)


def encontrar_arquivos_novas():
    """
    Retorna lista com todos os .xlsx/.xls encontrados na pasta 'novas atividades'.
    Independente do nome do arquivo — qualquer planilha na pasta sera processada.
    """
    PASTA_NOVAS.mkdir(parents=True, exist_ok=True)
    arquivos = [
        arq for arq in PASTA_NOVAS.iterdir()
        if arq.suffix.lower() in (".xlsx", ".xls") and arq.is_file()
    ]
    if arquivos:
        for arq in arquivos:
            print(f"   Arquivo encontrado: {arq.name}")
    else:
        print("   Nenhum arquivo .xlsx encontrado na pasta 'novas atividades'.")
    return arquivos


# ============================================================
# GERAÇÃO DE RELATÓRIO DE TEXTO
# ============================================================

def _menor_banda_4g(freq_234):
    """Retorna a menor banda 4G presente na string de frequências."""
    import re as _re
    ORDEM = [700, 850, 900, 1800, 2100, 2300, 2600]
    nums = [int(n) for n in _re.findall(r"\b(\d{3,4})(?:RS)?\b", freq_234)]
    for banda in ORDEM:
        if banda in nums:
            return str(banda)
    return None


def gerar_relatorio(df_atividades, df_base, out_dir):
    out_dir.mkdir(parents=True, exist_ok=True)
    mes_ano = datetime.now().strftime("/%m/%Y")
    all_texts = []

    # Cache de UF já perguntadas nesta execução (evita perguntar duas vezes o mesmo site)
    _uf_cache = {}

    for idx, row in df_atividades.iterrows():
        try:
            demanda    = str(row.get("DEMANDA",    "")).strip().upper()
            integracao = str(row.get("INTEGRAÇÃO", "")).strip().upper()
            site       = str(row.get("SITE",       "")).strip().upper()
            uf         = str(row.get("UF",         "")).strip().upper()
            cidade     = str(row.get("CIDADE",     "")).strip().upper()
            lat        = safe_float(row.get("LAT",  0))
            lon        = safe_float(row.get("LONG", 0))
            freq_234   = str(row.get("2G|3G|4G",   "")).strip()
            freq_5g    = str(row.get("5G",          "")).strip()

            # ─────────────────────────────────────────────────────────────
            # AJUSTE CIRÚRGICO 2 — busca PCI/AZIMUTH/UF na Base_4G
            #
            # Estratégia em dois passos:
            #   Passo A (restrito): lat + lon + cidade — match exato, mais confiável
            #   Passo B (relaxado): só lat + lon       — fallback quando cidade diverge
            #
            # Isso resolve o caso mais comum de falha: cidade com grafia diferente
            # (ex: "SAO PAULO" vs "SÃO PAULO", "RIBEIRAO" vs "RIBEIRÃO PRETO").
            # ─────────────────────────────────────────────────────────────

            # Passo A: match com cidade
            matches = pd.DataFrame()
            if cidade:
                mask_a = df_base.apply(
                    lambda b: (
                        float_close(lat, b["LATITUDE"])
                        and float_close(lon, b["LONGITUDE"])
                        and remove_acentos(cidade).lower() == remove_acentos(str(b["CIDADE"])).strip().lower()
                    ),
                    axis=1,
                )
                matches = df_base[mask_a]

            # Passo B: fallback só por lat/lon (tolerância 0.002°)
            if matches.empty:
                matches = _buscar_na_base4g(df_base, lat, lon, tolerancia=0.002)

            # ── Extrair PCI e AZIMUTH ────────────────────────────────────
            pci_list = [str(int(v)) for v in matches["PCI"].dropna().unique()]    if "PCI"     in matches.columns else []
            az_list  = [str(int(v)) for v in matches["AZIMUTH"].dropna().unique()] if "AZIMUTH" in matches.columns else []
            pci_str  = "/".join(unique_preserve_order(pci_list))
            az_str   = "/".join(unique_preserve_order(az_list))

            # ── Aviso quando PCI/AZIMUTH não forem encontrados ───────────
            # (para o usuário conferir e preencher manualmente se necessário)
            if not pci_str:
                print(f"   ⚠️  PCI não encontrado para o site {site} ({cidade}) — confira manualmente.")
            if not az_str:
                print(f"   ⚠️  AZIMUTH não encontrado para o site {site} ({cidade}) — confira manualmente.")

            # ── FIX UF vazia — sequência de 3 tentativas ────────────────
            if not uf or uf in ("", "NAN", "NONE"):
                uf_base = ""

                # Tentativa 1: extrair do match já obtido acima
                if not matches.empty and "[P]UF" in matches.columns:
                    ufs = matches["[P]UF"].dropna().unique()
                    if len(ufs) > 0:
                        uf_base = str(ufs[0]).strip().upper()

                if uf_base and len(uf_base) == 2 and uf_base.isalpha():
                    uf = uf_base
                    print(f"   ℹ️  UF do site {site} detectada na Base_4G: {uf}")
                else:
                    # Tentativa 2: cache desta execução
                    if site in _uf_cache:
                        uf = _uf_cache[site]
                    else:
                        # Tentativa 3: perguntar ao usuário
                        print(f"\n⚠️  UF não encontrada para o site {site} ({cidade}).")
                        while True:
                            resp = input(f"   Informe a UF (ex: SP, MG, RJ): ").strip().upper()
                            if len(resp) == 2 and resp.isalpha():
                                uf = resp
                                _uf_cache[site] = uf
                                break
                            print("   ⚠️  UF inválida. Digite apenas 2 letras (ex: SP).")

            # ── Montar corpo do relatório ────────────────────────────────────
            lines = [
                f"{site}",
                "",
                f"{lat} {lon}",
                "",
                f"({cidade} - {uf})",
                "",
                f"PCI: {pci_str}",
                f"AZIMUTH:  {az_str}",
                "",
                "(Frequências):",
                "",
            ]

            if freq_234 and freq_234.lower() != "nan":
                for bloco in [b.strip() for b in freq_234.split("|") if b.strip()]:
                    lines.append(bloco)

            if freq_5g and freq_5g.lower() != "nan":
                lines.append(freq_5g if freq_5g.upper().startswith("5G:") else f"5G: {freq_5g}")

            lines += [
                "",
                "OBS.>",
                "",
                f"{demanda} - {integracao}",
                "",
                "LOGS armazenados no servidor.",
                f"> Finalizado:  {mes_ano}",
                "",
                "--------------------------- Nome - LOGS --------------------------------------",
                "",
            ]

            # ── Nome do log — menor banda 4G ────────────────────────────
            si  = remove_acentos(integracao).replace(" ", "_")
            ss  = remove_acentos(site).replace(" ", "_")
            sc  = remove_acentos(cidade).replace(" ", "_")
            suf = remove_acentos(uf)

            menor_4g = _menor_banda_4g(freq_234)
            if menor_4g:
                lines.append(f"_{si}_SSV_{ss}_4G_{menor_4g}_{sc}_{suf}_")
            if "3500" in freq_5g:
                lines.append(f"_{si}_SSV_{ss}_5G_3500_{sc}_{suf}_")

            all_texts.append("\n".join(lines))

        except Exception as e:
            print(f"   ⚠️  Erro na linha {idx}: {e}")

    sep = "\n" + ("-" * 89 + "\n") * 4 + "\n"
    full_text = sep.join(all_texts)

    out_file = out_dir / "RELATORIO_SSV_COMPLETO.txt"
    with open(out_file, "w", encoding="utf-8-sig") as f:
        f.write(full_text)

    print(f"   Arquivo: {out_file}")


# ============================================================
# GERAÇÃO DE MAPA
# ============================================================

def carregar_hoteis():
    """
    Lê a planilha HOTEIS diretamente do Google Sheets via ID fixo.

    Planilha: https://docs.google.com/spreadsheets/d/HOTEIS_SHEET_ID
    Aba identificada pelo gid HOTEIS_GID (965678690).

    Colunas esperadas (aceita variações de nome, case-insensitive, sem acento):
      NOME | CIDADE | LATITUDE | LONGITUDE | TELEFONE | ULTIMO_VALOR

    Fallback: se não conseguir acessar o Sheets, tenta HOTEIS.xlsx local.
    """

    def _parse_hoteis(cabecalhos_raw, linhas):
        """Converte cabeçalhos + linhas brutas em lista de dicts de hotéis."""
        cabecalhos = [
            unicodedata.normalize("NFKD", str(c).strip().upper())
            .encode("ascii", errors="ignore").decode("utf-8")
            for c in cabecalhos_raw
        ]
        col_map = {
            "nome":   next((i for i, c in enumerate(cabecalhos) if "NOME"   in c), None),
            "cidade": next((i for i, c in enumerate(cabecalhos) if "CIDADE" in c), None),
            "lat":    next((i for i, c in enumerate(cabecalhos) if c.startswith("LAT")), None),
            "lon":    next((i for i, c in enumerate(cabecalhos)
                            if c.startswith("LON") or c.startswith("LONG")), None),
            "tel":    next((i for i, c in enumerate(cabecalhos)
                            if "TEL" in c or "FONE" in c), None),
            "valor":  next((i for i, c in enumerate(cabecalhos)
                            if "VALOR" in c or "PRECO" in c), None),
        }

        def _cel(linha, chave):
            idx = col_map.get(chave)
            if idx is None or idx >= len(linha):
                return ""
            return str(linha[idx]).strip()

        hoteis = []
        for linha in linhas:
            if not any(linha):
                continue
            nome = _cel(linha, "nome")
            if not nome or nome.upper() in ("NAN", "NONE", ""):
                continue
            lat = safe_float(_cel(linha, "lat"))
            lon = safe_float(_cel(linha, "lon"))
            if lat == 0.0 and lon == 0.0:
                continue
            cidade = _cel(linha, "cidade")
            tel    = _cel(linha, "tel")
            valor  = _cel(linha, "valor")
            cidade = "" if cidade.upper() in ("NAN", "NONE") else cidade
            tel    = "" if tel.upper()    in ("NAN", "NONE") else tel
            valor  = "" if valor.upper()  in ("NAN", "NONE") else valor
            hoteis.append({
                "nome": nome, "cidade": cidade,
                "lat": lat,   "lon": lon,
                "tel": tel,   "valor": valor,
            })
        return hoteis

    # ── Tentativa principal: Google Sheets pelo ID fixo ───────────────────────
    try:
        client  = conectar_sheets()
        sh      = client.open_by_key(HOTEIS_SHEET_ID)

        # Localizar a aba pelo gid — mais confiável que buscar pelo nome
        ws_h = None
        for aba in sh.worksheets():
            if aba.id == HOTEIS_GID:
                ws_h = aba
                break

        # Fallback: primeira aba caso o gid não bata (planilha reformatada)
        if ws_h is None:
            ws_h = sh.get_worksheet(0)
            print(f"   ℹ️  Aba com gid {HOTEIS_GID} não encontrada — usando primeira aba.")

        dados = ws_h.get_all_values()
        if not dados or len(dados) < 2:
            print("   ⚠️  Planilha HOTEIS vazia no Google Sheets.")
            return []

        hoteis = _parse_hoteis(dados[0], dados[1:])
        print(f"   {len(hoteis)} hoteis carregados do Google Sheets (aba: {ws_h.title}).")
        return hoteis

    except gspread.exceptions.SpreadsheetNotFound:
        print(f"   ⚠️  Planilha HOTEIS (ID: {HOTEIS_SHEET_ID}) não encontrada.")
        print("      Verifique se a service account tem acesso à planilha.")
    except Exception as e:
        print(f"   ⚠️  Erro ao acessar HOTEIS no Google Sheets: {e}")

    # ── Fallback: HOTEIS.xlsx local ───────────────────────────────────────────
    print("   → Tentando fallback: HOTEIS.xlsx local...")
    if not HOTEIS_PATH.exists():
        print("   ⚠️  HOTEIS.xlsx também não encontrado — hoteis não serão exibidos.")
        return []

    try:
        df = pd.read_excel(HOTEIS_PATH, engine="openpyxl")
        df.columns = (
            df.columns.str.strip().str.upper()
            .str.normalize("NFKD")
            .str.encode("ascii", errors="ignore")
            .str.decode("utf-8")
        )
        # Reutiliza _parse_hoteis convertendo o DataFrame em listas
        cabecalhos = list(df.columns)
        linhas = df.astype(str).values.tolist()
        hoteis = _parse_hoteis(cabecalhos, linhas)
        print(f"   {len(hoteis)} hoteis carregados de HOTEIS.xlsx (fallback local).")
        return hoteis
    except Exception as e:
        print(f"   ❌ Erro ao ler HOTEIS.xlsx: {e}")
        return []


def gerar_mapa(df_rota, lat0, lon0, cidade0, df_fixas=None, df_aguardando=None):
    """
    Gera MAPA_ROTAS.html com HTML/JS puro (sem folium).

    Cores:
      🏠 verde  = ponto de partida (localização atual)
      🔵 azul   = atividade pendente (entra na rota)
      ✅ verde  = concluída (sem rota, só marcador)
      ❌ vermelho = improdutiva (sem rota, só marcador)
    """
    import json

    ST_CONC = "✓ Atividade concluída"
    ST_IMPR = "IMPRODUTIVO"

    # Carregar hoteis da planilha HOTEIS.xlsx
    hoteis = carregar_hoteis()

    # Montar pontos da rota (pendentes)
    pontos_rota = []
    for _, row in df_rota.iterrows():
        lat = safe_float(row.get("LAT", 0))
        lon = safe_float(row.get("LONG", 0))
        if lat == 0 and lon == 0:
            continue
        pontos_rota.append({
            "id":     str(row.get("SITE",   "")),
            "cidade": str(row.get("CIDADE", "")),
            "tec4g":  str(row.get("2G|3G|4G", "")),
            "tec5g":  str(row.get("5G",     "")),
            "hotel":  str(row.get("HOTEL",  "") or ""),
            "lat":    lat,
            "lon":    lon,
        })

    # Montar pontos fixos (concluídas / improdutivas)
    pontos_fixos = []
    if df_fixas is not None and not df_fixas.empty:
        for _, row in df_fixas.iterrows():
            lat = safe_float(row.get("LAT", 0))
            lon = safe_float(row.get("LONG", 0))
            if lat == 0 and lon == 0:
                continue
            status = str(row.get("STATUS", "")).strip()
            tipo = "concluida" if ST_CONC in status else "improdutiva"
            pontos_fixos.append({
                "id":        str(row.get("SITE",      "")),
                "cidade":    str(row.get("CIDADE",    "")),
                "tec4g":     str(row.get("2G|3G|4G",  "")),
                "tec5g":     str(row.get("5G",        "")),
                "concluido": str(row.get("CONCLUIDO", "") or ""),
                "hotel":     str(row.get("HOTEL",     "") or ""),
                "tipo":      tipo,
                "lat":       lat,
                "lon":       lon,
            })

    # Montar pontos aguardando
    pontos_aguard = []
    if df_aguardando is not None and not df_aguardando.empty:
        for _, row in df_aguardando.iterrows():
            lat = safe_float(row.get("LAT", 0))
            lon = safe_float(row.get("LONG", 0))
            if lat == 0 and lon == 0:
                continue
            pontos_aguard.append({
                "id":        str(row.get("SITE",      "")),
                "cidade":    str(row.get("CIDADE",    "")),
                "tec4g":     str(row.get("2G|3G|4G",  "")),
                "tec5g":     str(row.get("5G",        "")),
                "concluido": str(row.get("CONCLUIDO", "") or ""),
                "hotel":     str(row.get("HOTEL",     "") or ""),
                "lat":       lat,
                "lon":       lon,
            })

    partida = {"lat": lat0, "lon": lon0, "cidade": cidade0}

    j_rota   = json.dumps(pontos_rota,   ensure_ascii=False)
    j_fixos  = json.dumps(pontos_fixos,  ensure_ascii=False)
    j_aguard = json.dumps(pontos_aguard, ensure_ascii=False)
    j_hoteis = json.dumps(hoteis,        ensure_ascii=False)
    j_part   = json.dumps(partida,       ensure_ascii=False)

    html = f"""<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8"/>
<meta name="viewport" content="width=device-width,initial-scale=1.0,maximum-scale=1.0,user-scalable=no"/>
<title>DT 3.0 — Mapa de Rota</title>
<link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/leaflet@1.9.3/dist/leaflet.css"/>
<link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap@5.2.2/dist/css/bootstrap.min.css"/>
<link rel="stylesheet" href="https://netdna.bootstrapcdn.com/bootstrap/3.0.0/css/bootstrap-glyphicons.css"/>
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/Leaflet.awesome-markers/2.0.2/leaflet.awesome-markers.css"/>
<script src="https://cdn.jsdelivr.net/npm/leaflet@1.9.3/dist/leaflet.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/Leaflet.awesome-markers/2.0.2/leaflet.awesome-markers.js"></script>
<style>
  html,body,#map{{width:100%;height:100%;margin:0;padding:0;}}
  #painel{{
    position:absolute;top:10px;right:10px;z-index:1000;
    background:rgba(255,255,255,0.96);border-radius:10px;
    padding:14px 18px;box-shadow:0 4px 18px rgba(0,0,0,0.18);
    font-family:'Segoe UI',sans-serif;min-width:190px;
  }}
  #painel h4{{margin:0 0 10px;font-size:14px;font-weight:700;color:#1a237e;letter-spacing:.5px;}}
  .leg{{display:flex;align-items:center;gap:8px;margin-bottom:6px;font-size:13px;color:#333;}}
  .dot{{width:13px;height:13px;border-radius:50%;display:inline-block;flex-shrink:0;}}
  .dot-pend{{background:#2A52BE;}}
  .dot-conc{{background:#27ae60;}}
  .dot-impr{{background:#e74c3c;}}
  .dot-part{{background:#f39c12;border:2px solid #b7770d;}}
  .dot-hotel{{background:#e67e22;}}
  .dot-aguard{{background:#8e44ad;border:2px solid #6c3483;}}
  #contador{{margin-top:10px;padding-top:10px;border-top:1px solid #ddd;font-size:12px;color:#555;}}
  #contador span{{font-weight:700;color:#1a237e;}}
  .sep{{border:none;border-top:1px solid #e0e0e0;margin:10px 0;}}
  .filtro{{display:flex;align-items:center;gap:8px;margin-bottom:5px;font-size:12px;color:#444;cursor:pointer;user-select:none;}}
  .filtro input[type=checkbox]{{width:14px;height:14px;cursor:pointer;accent-color:#1a237e;flex-shrink:0;}}
  .filtro span.dot{{flex-shrink:0;}}
  #filtros-label{{font-size:11px;font-weight:700;color:#888;letter-spacing:.6px;text-transform:uppercase;margin-bottom:6px;}}
  .tip-site{{font-weight:bold;font-size:13px;color:#1a237e;}}
  .tip-hotel{{font-weight:bold;font-size:14px;color:#d35400;
    background:white;border:2px solid #e67e22;padding:4px 7px;border-radius:4px;}}
  .tip-conc{{font-size:13px;color:#27ae60;font-weight:600;}}
  .tip-impr{{font-size:13px;color:#e74c3c;font-weight:600;}}
  .popup-box{{font-family:'Segoe UI',sans-serif;font-size:13px;min-width:200px;}}
  .popup-box b{{color:#1a237e;}}
  .popup-status{{margin-top:6px;padding:4px 8px;border-radius:4px;font-weight:600;font-size:12px;}}
  .ps-conc{{background:#d4edda;color:#155724;}}
  .ps-impr{{background:#f8d7da;color:#721c24;}}
  .btn-concluir{{
    margin-top:10px;background:#27ae60;color:white;border:none;
    padding:10px;border-radius:5px;cursor:pointer;width:100%;
    font-weight:700;font-size:15px;letter-spacing:.3px;
  }}
  .btn-concluir:hover{{background:#1e8449;}}
</style>
</head>
<body>
<div id="map"></div>
<div id="painel">
  <h4>🗺 DT 3.0 — Rota</h4>
  <div class="leg"><span class="dot dot-part"></span> Ponto de partida</div>
  <div class="leg"><span class="dot dot-pend"></span> Pendente (rota ativa)</div>
  <div id="contador">
    Pendentes: <span id="cnt-pend">0</span> |
    Concluidas: <span id="cnt-conc">0</span>
  </div>
  <hr class="sep"/>
  <div id="filtros-label">Exibir camadas</div>
  <label class="filtro">
    <input type="checkbox" id="chk-conc" checked/>
    <span class="dot dot-conc"></span> Concluidas
  </label>
  <label class="filtro">
    <input type="checkbox" id="chk-impr" checked/>
    <span class="dot dot-impr"></span> Improdutivas
  </label>
  <label class="filtro">
    <input type="checkbox" id="chk-hotel" checked/>
    <span class="dot dot-hotel"></span> Hoteis
  </label>
  <label class="filtro">
    <input type="checkbox" id="chk-aguard" checked/>
    <span class="dot dot-aguard"></span> Aguardando
  </label>
</div>
<script>
var map = L.map("map").setView([{lat0},{lon0}], 7);
L.tileLayer("https://tile.openstreetmap.org/{{z}}/{{x}}/{{y}}.png",
  {{attribution:'&copy; <a href="https://www.openstreetmap.org/copyright">OpenStreetMap</a> contributors',maxZoom:19}}).addTo(map);

var PARTIDA   = {j_part};
var ROTA      = {j_rota};
var FIXOS     = {j_fixos};
var AGUARD    = {j_aguard};
var HOTEIS    = {j_hoteis};

// contador de marcações manuais feitas no mapa nesta sessão
var concCount = 0;

// ── Ícones ──────────────────────────────────────────────────────────────────
function mkIcon(color, icon){{
  return L.AwesomeMarkers.icon({{icon:icon, markerColor:color, prefix:'glyphicon'}});
}}
var icPartida  = mkIcon('orange', 'star');
var icPend     = mkIcon('blue',   'map-marker');
var icConc     = mkIcon('green',  'ok-sign');
var icImpr     = mkIcon('red',    'remove-sign');
var icHotel    = mkIcon('orange', 'tower');

// ── Hotéis ───────────────────────────────────────────────────────────────────
var mkHoteis = [];
HOTEIS.forEach(function(h){{
  var mk = L.marker([h.lat,h.lon],{{icon:icHotel}}).addTo(map);
  mk.bindTooltip(h.nome,{{sticky:true,className:'tip-hotel',direction:'top'}});
  mk.bindPopup(
    "<div class='popup-box'>"+
    "<b>"+h.nome+"</b><br>"+
    (h.cidade ? "<b>Cidade:</b> "+h.cidade+"<br>" : "")+
    (h.tel    ? "<b>Telefone:</b> "+h.tel+"<br>"  : "")+
    (h.valor  ? "<b>Ultimo valor:</b> R$ "+h.valor+"<br>" : "")+
    "<div class='popup-status' style='background:#fef3e2;color:#b7600a;margin-top:6px;"+
    "padding:4px 8px;border-radius:4px;font-weight:600;font-size:12px;'>"+
    "🏨 Hotel / Pousada</div>"+
    "</div>"
  );
  mkHoteis.push(mk);
}});

// ── Ponto de partida ─────────────────────────────────────────────────────────
var mkPartida = L.marker([PARTIDA.lat,PARTIDA.lon],{{icon:icPartida}})
  .addTo(map)
  .bindTooltip("📍 Partida: "+PARTIDA.cidade,{{sticky:true,className:'tip-site'}})
  .bindPopup("<div class='popup-box'><b>Ponto de partida</b><br>"+PARTIDA.cidade+"</div>");

// ── Rota e polilinha ──────────────────────────────────────────────────────────
var marcadores = [];   // só pendentes, na ordem da rota
var polyline;

function coordsRota(){{
  var pts = [[PARTIDA.lat,PARTIDA.lon]];
  marcadores.forEach(function(m){{
    if(map.hasLayer(m.marker)) pts.push([m.lat,m.lon]);
  }});
  return pts;
}}

function redesenharRota(){{
  if(polyline) map.removeLayer(polyline);
  var pts = coordsRota();
  if(pts.length > 1){{
    polyline = L.polyline(pts,{{color:'#2A52BE',weight:4,opacity:0.65,
      dashArray: null}}).addTo(map);
  }}
  // Atualizar contador
  var pend = marcadores.filter(function(m){{return map.hasLayer(m.marker);}}).length;
  var concFixas = FIXOS.filter(function(p){{return p.tipo === 'concluida';}}).length;
  document.getElementById('cnt-pend').textContent = pend;
  document.getElementById('cnt-conc').textContent = concFixas + concCount;
}}

ROTA.forEach(function(p, idx){{
  var marker = L.marker([p.lat,p.lon],{{icon:icPend}}).addTo(map);

  var tip = "SITE: "+p.id+"<br>"+p.cidade;
  marker.bindTooltip(tip,{{sticky:true,className:'tip-site'}});

  var pop = document.createElement('div');
  pop.className = 'popup-box';
  var freqStr = [p.tec4g, p.tec5g].filter(Boolean).join(' | ');
  pop.innerHTML =
    "<b>"+p.id+"</b><br>"+
    "<b>Cidade:</b> "+p.cidade+"<br>"+
    (freqStr ? "<b>Freq:</b> "+freqStr+"<br>" : "")+
    (p.hotel  ? "<b>Hotel:</b> "+p.hotel+"<br>" : "");

  var btn = document.createElement('button');
  btn.className = 'btn-concluir';
  btn.innerHTML = '✓ Marcar como Concluída';
  btn.onclick = function(){{
    map.removeLayer(marker);
    // Adiciona marcador verde no lugar
    var mk2 = L.marker([p.lat,p.lon],{{icon:icConc}}).addTo(map);
    mk2.className = 'marcador-conc';   // para contagem
    mk2.bindTooltip("✓ "+p.id,{{sticky:true,className:'tip-conc'}});
    mk2.bindPopup("<div class='popup-box'><b>"+p.id+"</b><br>"+p.cidade+
      "<div class='popup-status ps-conc'>✓ Concluída</div></div>");
    redesenharRota();
  }};
  pop.appendChild(btn);
  marker.bindPopup(pop);

  marcadores.push({{marker:marker, lat:p.lat, lon:p.lon}});
}});

// ── Fixos (concluídas / improdutivas do mês) ─────────────────────────────────
var mkConc = [];
var mkImpr = [];
FIXOS.forEach(function(p){{
  var ic  = p.tipo === 'concluida' ? icConc : icImpr;
  var cls = p.tipo === 'concluida' ? 'tip-conc' : 'tip-impr';
  var label = p.tipo === 'concluida' ? '✓ ' : '✗ ';
  var mk = L.marker([p.lat,p.lon],{{icon:ic}}).addTo(map);
  mk.bindTooltip(label+p.id,{{sticky:true,className:cls}});
  var statusHtml = p.tipo === 'concluida'
    ? "<div class='popup-status ps-conc'>✓ Atividade concluída</div>"
    : "<div class='popup-status ps-impr'>✗ Improdutiva</div>";
  var freqStr = [p.tec4g,p.tec5g].filter(Boolean).join(' | ');
  mk.bindPopup(
    "<div class='popup-box'>"+
    "<b>"+p.id+"</b><br>"+
    "<b>Cidade:</b> "+p.cidade+"<br>"+
    (freqStr ? "<b>Freq:</b> "+freqStr+"<br>" : "")+
    (p.concluido ? "<b>Obs:</b> "+p.concluido+"<br>" : "")+
    (p.hotel     ? "<b>Hotel:</b> "+p.hotel+"<br>" : "")+
    statusHtml+
    "</div>"
  );
  if(p.tipo === 'concluida') mkConc.push(mk);
  else mkImpr.push(mk);
}});

// ── Aguardando para deslocar ─────────────────────────────────────────────────
var icAguard = mkIcon('purple', 'time');
var mkAguard = [];
AGUARD.forEach(function(p){{
  var mk = L.marker([p.lat,p.lon],{{icon:icAguard}}).addTo(map);
  mk.bindTooltip("⏳ "+p.id+"<br>"+p.cidade,{{sticky:true,className:'tip-site'}});
  var freqStr = [p.tec4g,p.tec5g].filter(Boolean).join(' | ');
  mk.bindPopup(
    "<div class='popup-box'>"+
    "<b>"+p.id+"</b><br>"+
    "<b>Cidade:</b> "+p.cidade+"<br>"+
    (freqStr ? "<b>Freq:</b> "+freqStr+"<br>" : "")+
    (p.concluido ? "<b>Motivo:</b> "+p.concluido+"<br>" : "")+
    (p.hotel && p.hotel !== '.' ? "<b>Hotel:</b> "+p.hotel+"<br>" : "")+
    "<div class='popup-status' style='background:#e8daef;color:#6c3483;margin-top:6px;"+
    "padding:4px 8px;border-radius:4px;font-weight:600;font-size:12px;'>"+
    "⏳ Aguardando para deslocar</div>"+
    "</div>"
  );
  mkAguard.push(mk);
}});

// ── Checkboxes de visibilidade ────────────────────────────────────────────────
function toggleCamada(lista, visivel){{
  lista.forEach(function(mk){{
    if(visivel) {{ if(!map.hasLayer(mk)) mk.addTo(map); }}
    else        {{ if(map.hasLayer(mk))  map.removeLayer(mk); }}
  }});
}}

document.getElementById('chk-conc').addEventListener('change', function(){{
  toggleCamada(mkConc, this.checked);
}});
document.getElementById('chk-impr').addEventListener('change', function(){{
  toggleCamada(mkImpr, this.checked);
}});
document.getElementById('chk-hotel').addEventListener('change', function(){{
  toggleCamada(mkHoteis, this.checked);
}});
document.getElementById('chk-aguard').addEventListener('change', function(){{
  toggleCamada(mkAguard, this.checked);
}});

// ── Inicializar ───────────────────────────────────────────────────────────────
redesenharRota();

var todosMarcadores = marcadores.map(function(m){{return m.marker;}});
todosMarcadores.push(mkPartida);
if(todosMarcadores.length > 0){{
  var grp = L.featureGroup(todosMarcadores);
  map.fitBounds(grp.getBounds().pad(0.1));
}}

setTimeout(function(){{map.invalidateSize();}}, 400);
</script>
</body>
</html>"""

    path = BASE_DIR / "MAPA_ROTAS.html"
    with open(path, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"   Arquivo: {path}")


# ============================================================
# PUBLICAR MAPA NO GITHUB PAGES
# ============================================================

def publicar_mapa_github(html_path):
    """
    Faz commit do MAPA_ROTAS.html no repositório GitHub configurado
    em github_token.txt. Usa apenas a biblioteca requests — sem git.

    Formato do github_token.txt (3 linhas):
        TOKEN_AQUI
        seu-usuario
        nome-do-repositorio
    """
    if not GITHUB_TOKEN_PATH.exists():
        print("   ⚠️  github_token.txt nao encontrado — publicacao no GitHub ignorada.")
        print(f"      Crie o arquivo em: {GITHUB_TOKEN_PATH}")
        return

    try:
        linhas = GITHUB_TOKEN_PATH.read_text(encoding="utf-8").strip().splitlines()
        if len(linhas) < 3:
            print("   ❌ github_token.txt incompleto. Veja o GUIA_GITHUB.txt.")
            return
        token    = linhas[0].strip()
        usuario  = linhas[1].strip()
        repo     = linhas[2].strip()
    except Exception as e:
        print(f"   ❌ Erro ao ler github_token.txt: {e}")
        return

    import base64

    # Ler HTML gerado e converter para Base64
    try:
        conteudo_bytes = html_path.read_bytes()
        conteudo_b64   = base64.b64encode(conteudo_bytes).decode("utf-8")
    except Exception as e:
        print(f"   ❌ Erro ao ler {html_path.name}: {e}")
        return

    arquivo_remoto = "MAPA_ROTAS.html"
    api_url = f"https://api.github.com/repos/{usuario}/{repo}/contents/{arquivo_remoto}"
    headers = {
        "Authorization": f"token {token}",
        "Accept":        "application/vnd.github+json",
    }

    # Verificar se o arquivo já existe (precisamos do SHA para atualizar)
    sha_atual = None
    try:
        r = requests.get(api_url, headers=headers, timeout=10)
        if r.status_code == 200:
            sha_atual = r.json().get("sha")
    except Exception as e:
        print(f"   ❌ Erro ao consultar GitHub: {e}")
        return

    # Montar payload do commit
    agora   = datetime.now().strftime("%d/%m/%Y %H:%M")
    payload = {
        "message": f"DT 3.0 — Mapa atualizado em {agora}",
        "content": conteudo_b64,
        "branch":  "main",
    }
    if sha_atual:
        payload["sha"] = sha_atual   # obrigatório para atualizar arquivo existente

    # Enviar
    try:
        r = requests.put(api_url, headers=headers, json=payload, timeout=30)
        if r.status_code in (200, 201):
            acao = "atualizado" if sha_atual else "criado"
            url_pagina = f"https://{usuario}.github.io/{repo}/{arquivo_remoto}"
            print(f"   ✅ Mapa {acao} no GitHub!")
            print(f"   🌐 Acesse: {url_pagina}")
        else:
            print(f"   ❌ GitHub retornou status {r.status_code}: {r.json().get('message','')}")
    except Exception as e:
        print(f"   ❌ Erro ao publicar no GitHub: {e}")


# ============================================================
# MODO MAPA — atualiza apenas o mapa a partir do Google Sheets
# ============================================================

def main_mapa():
    """
    Executa apenas o passo de mapa:
    1. Conecta ao Google Sheets
    2. Lê status atuais (concluídas, improdutivas, aguardando, pendentes)
    3. Gera novo MAPA_ROTAS.html
    4. Publica no GitHub Pages
    Não otimiza rota, não gera relatório, não processa novas atividades.
    """
    print("\n" + "=" * 60)
    print("  DT 3.0 — Atualizar Mapa")
    print("=" * 60)

    if not CREDS_PATH.exists():
        print(f"\n❌ credentials.json não encontrado em:\n   {BASE_DIR}")
        input("\nPressione ENTER para sair...")
        return

    # Conectar ao Sheets
    print("\n[1/3] Conectando ao Google Sheets...")
    try:
        client = conectar_sheets()
        sheet  = client.open_by_key(SHEET_ID)
        ws     = obter_aba_vigente(sheet)
    except Exception as e:
        print(f"   ❌ Erro: {e}")
        input("\nPressione ENTER para sair...")
        return

    # Ler atividades
    print("\n[2/3] Lendo atividades da planilha...")
    df_sheets = ler_atividades_sheets(ws)
    print(f"   {len(df_sheets)} atividades encontradas.")

    # Separar por status
    df_fixas      = df_sheets[df_sheets["STATUS"].isin(STATUS_FIXOS)].copy()
    df_aguardando = df_sheets[df_sheets["STATUS"].str.strip() == ST_AGUARDANDO].copy()
    df_pendentes  = df_sheets[
        ~df_sheets["STATUS"].isin(STATUS_FIXOS) &
        (df_sheets["STATUS"].str.strip() != ST_AGUARDANDO)
    ].copy()
    df_pendentes  = df_pendentes.dropna(subset=["LAT", "LONG"])
    df_pendentes  = df_pendentes[(df_pendentes["LAT"] != 0) & (df_pendentes["LONG"] != 0)]

    print(f"   Concluidas/Improdutivas : {len(df_fixas)}")
    print(f"   Pendentes na rota       : {len(df_pendentes)}")
    print(f"   Aguardando              : {len(df_aguardando)}")

    # Determinar ponto de partida
    lat0, lon0, cidade0 = determinar_ponto_inicio(df_sheets)

    # Gerar mapa com a ordem atual da planilha (sem reotimizar)
    print("\n[3/3] Gerando mapa e publicando...")
    print("   → Mapa interativo...")
    gerar_mapa(df_pendentes, lat0, lon0, cidade0,
               df_fixas=df_fixas, df_aguardando=df_aguardando)

    print("   → Publicando no GitHub...")
    publicar_mapa_github(BASE_DIR / "MAPA_ROTAS.html")

    print("\n" + "=" * 60)
    print("  ✅ Mapa atualizado com sucesso!")
    print("=" * 60)

    try:
        if sys.stdin and sys.stdin.isatty():
            input("\nPressione ENTER para sair...")
    except Exception:
        pass


# ============================================================
# MAIN
# ============================================================

def main():
    print("\n" + "=" * 60)
    print("  DT 3.0 — Automação Drive Test")
    print("=" * 60)

    # ── Verificações iniciais ──────────────────────────────
    if not CREDS_PATH.exists():
        print(f"\n❌ credentials.json não encontrado em:\n   {BASE_DIR}")
        print("   Veja o arquivo GUIA_API_GOOGLE.txt para configurar.")
        input("\nPressione ENTER para sair...")
        return

    if not BASE_4G_PATH.exists():
        print(f"\n❌ Base_4G.xlsx não encontrado em:\n   {BASE_DIR}")
        input("\nPressione ENTER para sair...")
        return

    # ── 1. Base 4G ─────────────────────────────────────────
    print("\n[1/7] Carregando Base 4G...")
    df_base = pd.read_excel(BASE_4G_PATH, engine="openpyxl")
    print(f"   {len(df_base):,} registros.")

    # ── 2. Google Sheets ───────────────────────────────────
    print("\n[2/7] Conectando ao Google Sheets...")
    try:
        client = conectar_sheets()
        sheet  = client.open_by_key(SHEET_ID)
        ws     = obter_aba_vigente(sheet)
    except Exception as e:
        print(f"   ❌ Erro: {e}")
        input("\nPressione ENTER para sair...")
        return

    # ── 3. Ler atividades do Sheets ────────────────────────
    print("\n[3/7] Lendo atividades da planilha...")
    df_sheets = ler_atividades_sheets(ws)
    print(f"   {len(df_sheets)} atividades encontradas.")

    if df_sheets.empty:
        print("   ⚠️  Planilha vazia. Continue assim que houver dados.")

    # ── 4. Ponto de partida ────────────────────────────────
    print("\n[4/7] Determinando ponto de partida...")
    lat0, lon0, cidade0 = determinar_ponto_inicio(df_sheets)

    # ── 5. Novas atividades ────────────────────────────────
    print("\n[5/7] Processando novas atividades...")
    arquivos_novas = encontrar_arquivos_novas()
    df_novas       = pd.DataFrame()

    if arquivos_novas:
        frames_novas = []
        for arq in arquivos_novas:
            try:
                df_raw = pd.read_excel(arq, engine="openpyxl")
                df_proc = processar_arquivo(df_raw, nome_arquivo=arq.name, df_base=df_base)
                if not df_proc.empty:
                    frames_novas.append(df_proc)
                    print(f"   {arq.name}: {len(df_proc)} atividades válidas.")
                else:
                    print(f"   ⚠️  {arq.name}: nenhuma atividade válida extraída.")
            except Exception as e:
                print(f"   ❌ Erro ao processar {arq.name}: {e}")

        if frames_novas:
            df_novas = pd.concat(frames_novas, ignore_index=True)
            # Remover duplicatas de SITE caso o mesmo site apareça em dois arquivos
            df_novas = df_novas.drop_duplicates(subset=["SITE"], keep="first")
            print(f"   Total: {len(df_novas)} novas atividades únicas.")
    else:
        print("   Nenhuma nova atividade. Continuando sem novas.")

    # ── 6. Montar pool e otimizar ──────────────────────────
    print("\n[6/7] Otimizando rota...")

    # Fixas: não entram na otimização
    df_fixas = df_sheets[df_sheets["STATUS"].isin(STATUS_FIXOS)].copy()

    # Aguardando: ficam no final da planilha, fora da otimização
    df_aguardando = df_sheets[
        df_sheets["STATUS"].str.strip() == ST_AGUARDANDO
    ].copy()

    # Pool: atividades não fixas e não "Aguardando" que estão na planilha
    df_pool_sheets = df_sheets[
        ~df_sheets["STATUS"].isin(STATUS_FIXOS) &
        (df_sheets["STATUS"].str.strip() != ST_AGUARDANDO)
    ].copy()
    df_pool_sheets = df_pool_sheets.dropna(subset=["LAT", "LONG"])
    df_pool_sheets = df_pool_sheets[
        (df_pool_sheets["LAT"] != 0) & (df_pool_sheets["LONG"] != 0)
    ]

    # ── AJUSTE: complementar UF e TEC ausentes no pool do Sheets via Base_4G ──
    # Nunca toca em STATUS, CONCLUIDO, HOTEL.
    for idx, row in df_pool_sheets.iterrows():
        precisa_uf  = not str(row.get("UF",  "")).strip()
        precisa_tec = not str(row.get("TEC", "")).strip()
        if not (precisa_uf or precisa_tec):
            continue
        try:
            matches = _buscar_na_base4g(df_base, row["LAT"], row["LONG"])
            if matches.empty:
                continue
            m = matches.iloc[0]
            if precisa_uf and "[P]UF" in matches.columns:
                val = str(m.get("[P]UF", "")).strip().upper()
                if val and val not in ("", "NAN", "NONE") and len(val) == 2 and val.isalpha():
                    df_pool_sheets.at[idx, "UF"] = val
                    print(f"   ℹ️  {row['SITE']}: UF complementada do Sheets via Base_4G: {val}")
            if precisa_tec:
                for col_tec in ("TEC", "TECNOLOGIA", "TECHNOLOGY"):
                    if col_tec in matches.columns:
                        val = normalizar_tec(str(m.get(col_tec, "")).strip())
                        if val and val.upper() not in ("", "NAN", "NONE"):
                            df_pool_sheets.at[idx, "TEC"] = val
                            print(f"   ℹ️  {row['SITE']}: TEC complementada do Sheets via Base_4G: {val}")
                            break
        except Exception:
            pass

    # Sites já no pool para evitar duplicatas
    sites_originais  = set(df_sheets["SITE"].str.upper())
    sites_no_pool    = set(df_pool_sheets["SITE"].str.upper())

    # Novas que ainda não estão na planilha
    COLS = ["DEMANDA", "INTEGRAÇÃO", "SITE", "UF", "TEC",
            "LAT", "LONG", "CIDADE", "2G|3G|4G", "5G", "STATUS",
            "CONCLUIDO", "HOTEL"]

    if not df_novas.empty:
        # Novas não têm CONCLUIDO/HOTEL — preencher vazios
        df_nov = df_novas[~df_novas["SITE"].str.upper().isin(sites_no_pool)].copy()
        for col in ("CONCLUIDO", "HOTEL"):
            if col not in df_nov.columns:
                df_nov[col] = ""
        df_novas_filtradas = df_nov[COLS].copy()
    else:
        df_novas_filtradas = pd.DataFrame(columns=COLS)

    # Unir para otimização (pool_sheets já tem CONCLUIDO e HOTEL)
    frames = [f for f in [df_pool_sheets[COLS], df_novas_filtradas] if not f.empty]

    if not frames:
        print("   ⚠️  Nenhuma atividade para otimizar.")
        input("\nPressione ENTER para sair...")
        return

    df_pool = pd.concat(frames, ignore_index=True)
    print(f"   Pool: {len(df_pool)} atividades ({len(df_pool_sheets)} existentes + {len(df_novas_filtradas)} novas)")

    rota_final = otimizar_rota(df_pool, lat0, lon0)

    # ── Exportar ATIVIDADES_GERADAS.xlsx ──────────────────
    ativ_path = BASE_DIR / "ATIVIDADES_GERADAS.xlsx"
    rota_final.to_excel(ativ_path, index=False)
    print(f"   ATIVIDADES_GERADAS.xlsx salvo.")

    # ── 7. Saídas ──────────────────────────────────────────
    print("\n[7/7] Gerando saídas...")

    print("   → Relatório de texto...")
    gerar_relatorio(rota_final, df_base, PASTA_OUT)

    print("   → Mapa interativo...")
    gerar_mapa(rota_final, lat0, lon0, cidade0, df_fixas=df_fixas, df_aguardando=df_aguardando)

    print("   → Publicando mapa no GitHub...")
    publicar_mapa_github(BASE_DIR / "MAPA_ROTAS.html")

    print("   → Atualizando Google Sheets...")
    atualizar_sheets(ws, df_fixas, rota_final, sites_originais, df_aguardando)

    # ── Mover arquivos processados ────────────────────────
    if arquivos_novas:
        processados = PASTA_NOVAS / "processados"
        processados.mkdir(exist_ok=True)
        for arq in arquivos_novas:
            try:
                destino = processados / arq.name
                arq.rename(destino)
                print(f"   Movido: processados/{arq.name}")
            except Exception as e:
                print(f"   ⚠️  Nao foi possivel mover {arq.name}: {e}")

    print("\n" + "=" * 60)
    print("  ✅ DT 3.0 concluído com sucesso!")
    print("=" * 60)

    try:
        if sys.stdin and sys.stdin.isatty():
            input("\nPressione ENTER para sair...")
    except Exception:
        pass


# ============================================================

if __name__ == "__main__":
    print("\n" + "=" * 60)
    print("  DT 3.0 — Automação Drive Test")
    print("=" * 60)
    print()
    print("  Selecione o modo de execução:")
    print()
    print("  [1] Execução completa")
    print("      Processa novas atividades, otimiza rota,")
    print("      gera relatório, atualiza Sheets e mapa.")
    print()
    print("  [2] Atualizar mapa")
    print("      Lê status atuais do Sheets e gera novo")
    print("      mapa sem reprocessar nada mais.")
    print()

    while True:
        try:
            opcao = input("  Digite 1 ou 2: ").strip()
        except Exception:
            opcao = "1"

        if opcao == "1":
            try:
                main()
            except Exception as e:
                print(f"\n❌ Erro inesperado: {e}")
                import traceback
                traceback.print_exc()
                try:
                    input("\nPressione ENTER para sair...")
                except Exception:
                    pass
            break

        elif opcao == "2":
            try:
                main_mapa()
            except Exception as e:
                print(f"\n❌ Erro inesperado: {e}")
                import traceback
                traceback.print_exc()
                try:
                    input("\nPressione ENTER para sair...")
                except Exception:
                    pass
            break

        else:
            print("  ⚠️  Opção inválida. Digite 1 ou 2.")
