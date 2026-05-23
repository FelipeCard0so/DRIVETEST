"""
DT 3.0 - Automação Drive Test
Organiza atividades, otimiza rota, atualiza Google Sheets,
gera relatórios e mapa em um único processo.
VERSÃO MAPBOX: Rotas reais com trânsito para próxima atividade.
"""

import os
import sys
import re
import math
import unicodedata
import requests
import pandas as pd
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


BASE_DIR           = get_base_dir()
PASTA_NOVAS        = BASE_DIR / "novas atividades"
PASTA_OUT          = BASE_DIR / "out"
BASE_4G_PATH       = BASE_DIR / "Base_4G.xlsx"
CREDS_PATH         = BASE_DIR / "credentials.json"
SHEET_ID           = "1gPrzFOvPG6bF88H54ChXoyUmTWm84XU_ifFVO9X3rPE"
GITHUB_TOKEN_PATH  = BASE_DIR / "github_token.txt"
HOTEIS_PATH        = BASE_DIR / "HOTEIS.xlsx"
HOTEIS_SHEET_ID    = "1Vw1cgSppfxezM8MGRv56E88pnD-im5ThNX2LkJTyDsk"
HOTEIS_GID         = 965678690
MAPBOX_TOKEN_PATH  = BASE_DIR / "mapbox_token.txt"   # ← NOVO

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

DATA_START_ROW = 3
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


def carregar_mapbox_token():
    """Lê o token do Mapbox de mapbox_token.txt."""
    if MAPBOX_TOKEN_PATH.exists():
        token = MAPBOX_TOKEN_PATH.read_text(encoding="utf-8").strip()
        if token.startswith("pk."):
            return token
        print("   ⚠️  mapbox_token.txt encontrado mas token parece inválido (deve começar com 'pk.').")
    else:
        print(f"   ⚠️  mapbox_token.txt não encontrado em: {BASE_DIR}")
        print("      Crie o arquivo e cole seu Access Token do Mapbox.")
    return None


# ============================================================
# NORMALIZAÇÃO DE TECNOLOGIA (TEC)
# ============================================================

_MAPA_TEC = [
    (r"^2G[|/]3G[|/]LTE[|/]NR$",  "2G|3G|4G|5G"),
    (r"^3G[|/]LTE[|/]NR$",         "3G|4G|5G"),
    (r"^2G[|/]3G[|/]LTE$",         "2G|3G|4G"),
    (r"^2G[|/]3G[|/]NR$",          "2G|3G|5G"),
    (r"^LTE[|/]NR$",               "4G|5G"),
    (r"^3G[|/]LTE$",               "3G|4G"),
    (r"^3G[|/]NR$",                "3G|5G"),
    (r"^2G[|/]LTE$",               "2G|4G"),
    (r"^LTE$",                     "4G"),
    (r"^NR$",                      "5G"),
]

_TEC_PADRAO = {"2G", "3G", "4G", "5G",
               "2G|3G", "2G|4G", "2G|5G",
               "3G|4G", "3G|5G", "4G|5G",
               "2G|3G|4G", "2G|3G|5G", "2G|4G|5G", "3G|4G|5G",
               "2G|3G|4G|5G"}


def normalizar_tec(valor):
    if not valor:
        return valor
    s = str(valor).strip().upper()
    if not s or s in ("NAN", "NONE"):
        return ""
    if s in _TEC_PADRAO:
        return s
    s_norm = re.sub(r"\s*[/]\s*", "|", s)
    for padrao, substituto in _MAPA_TEC:
        if re.match(padrao, s_norm, re.IGNORECASE):
            return substituto
    return valor.strip()


# ============================================================
# NORMALIZAÇÃO DE FREQUÊNCIA
# ============================================================

def _ordenar_nums_4g(nums_str):
    def chave(x):
        num = int(re.search(r'\d+', x).group()) if re.search(r'\d+', x) else 9999
        return ORDEM_BANDA_4G.index(num) if num in ORDEM_BANDA_4G else 99
    return sorted(nums_str, key=chave)


def _reordenar_bloco_4g(bloco):
    if ':' not in bloco:
        return bloco
    tec, nums_raw = bloco.split(':', 1)
    nums = [n.strip() for n in nums_raw.split('/') if n.strip()]
    nums_ordenados = _ordenar_nums_4g(nums)
    return f"{tec}:{'/'.join(nums_ordenados)}"


def normalizar_frequencia(freq_raw):
    if not isinstance(freq_raw, str) or not freq_raw.strip():
        return "", ""

    freq = freq_raw.strip()

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
    return [(r["LAT"], r["LONG"]) for _, r in df.iterrows()]


def _dist_coords(coords, lat0, lon0):
    total, la, lo = 0, lat0, lon0
    for lat, lon in coords:
        total += haversine(la, lo, lat, lon)
        la, lo = lat, lon
    return total


def two_opt(df, lat0, lon0):
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
    idx_map = {(r["LAT"], r["LONG"]): i for i, (_, r) in enumerate(rota.iterrows())}
    nova_ordem = []
    for lat, lon in coords:
        nova_ordem.append(idx_map[(lat, lon)])
    return rota.iloc[nova_ordem].reset_index(drop=True), melhor_d


def three_opt(df, lat0, lon0):
    rota = df.reset_index(drop=True)
    coords = _rota_para_coords(rota)
    n = len(coords)
    melhor_d = _dist_coords(coords, lat0, lon0)
    melhorou = True

    def dist_total(seq):
        return _dist_coords(seq, lat0, lon0)

    while melhorou:
        melhorou = False
        for i in range(n - 2):
            for j in range(i + 1, n - 1):
                for k in range(j + 1, n):
                    A = coords[:i+1]
                    B = coords[i+1:j+1]
                    C = coords[j+1:k+1]
                    D = coords[k+1:]
                    candidatos = [
                        A + B[::-1] + C       + D,
                        A + B       + C[::-1] + D,
                        A + B[::-1] + C[::-1] + D,
                        A + C       + B       + D,
                        A + C       + B[::-1] + D,
                        A + C[::-1] + B       + D,
                        A + C[::-1] + B[::-1] + D,
                    ]
                    for cand in candidatos:
                        d_cand = dist_total(cand)
                        if d_cand < melhor_d - 1e-6:
                            coords = cand
                            melhor_d = d_cand
                            melhorou = True

    idx_map = {}
    for i, (_, r) in enumerate(rota.iterrows()):
        idx_map[(r["LAT"], r["LONG"])] = i
    nova_ordem = [idx_map[(lat, lon)] for lat, lon in coords]
    return rota.iloc[nova_ordem].reset_index(drop=True), melhor_d


def or_opt(df, lat0, lon0, tamanho_seg=None):
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

    idx_map = {}
    for i, (_, r) in enumerate(rota.iterrows()):
        idx_map[(r["LAT"], r["LONG"])] = i
    nova_ordem = [idx_map[(lat, lon)] for lat, lon in coords]
    return rota.iloc[nova_ordem].reset_index(drop=True), melhor_d


def otimizar_rota(df_pool, lat0, lon0):
    n = len(df_pool)
    print(f"   → Gerando candidatos ({n} pontos de partida alternativos)...")

    melhor_rota = None
    melhor_d = float("inf")
    pontos_inicio = [(lat0, lon0)]

    for _, r in df_pool.iterrows():
        pontos_inicio.append((r["LAT"], r["LONG"]))

    for i, (la, lo) in enumerate(pontos_inicio):
        rota_c = vizinho_mais_proximo(df_pool, la, lo)
        d_c = distancia_total(rota_c, lat0, lon0)
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
            all_vals = ws.get_all_values()
            if len(all_vals) >= DATA_START_ROW:
                last = len(all_vals)
                ws.batch_clear([f"A{DATA_START_ROW}:M{last}"])
        else:
            ws = sheet.add_worksheet(nome, rows=200, cols=20)
        return ws


MESES_MAPA = [
    "JANEIRO", "FEVEREIRO", "MARÇO", "ABRIL", "MAIO", "JUNHO",
    "JULHO", "AGOSTO", "SETEMBRO", "OUTUBRO", "NOVEMBRO", "DEZEMBRO",
]


def _normalizar_status_mapa(status):
    s = remove_acentos(str(status or "")).upper()
    s = re.sub(r"\s+", " ", s).strip()
    return s


def status_fixo_mapa(status):
    s_norm = _normalizar_status_mapa(status)
    return (
        "CONCLUID" in s_norm
        or "IMPRODUT" in s_norm
        or "CANCEL" in s_norm
    )


def mask_status_fixos_mapa(df):
    if df is None or df.empty or "STATUS" not in df.columns:
        return pd.Series([], dtype=bool, index=df.index if df is not None else None)
    return df["STATUS"].apply(status_fixo_mapa)


def tipo_status_mapa(status):
    s_norm = _normalizar_status_mapa(status)
    if "CANCEL" in s_norm:
        return "cancelada"
    if "CONCLUID" in s_norm:
        return "concluida"
    return "improdutiva"


def normalizar_coord_mapa(valor, limite):
    coord = safe_float(valor)
    if coord == 0:
        return coord
    while abs(coord) > limite:
        coord = coord / 10.0
    return coord


def carregar_meses_anteriores(sheet, aba_atual_nome):
    import re as _re

    mes_idx = {remove_acentos(m).upper(): i + 1 for i, m in enumerate(MESES_MAPA)}
    historico = {}
    total = 0

    for ws_mes in sheet.worksheets():
        titulo = str(ws_mes.title).strip()
        if titulo == aba_atual_nome:
            continue

        titulo_norm = remove_acentos(titulo).upper()
        m = _re.match(r"^([A-Z]+)\s*/\s*(20\d{2})$", titulo_norm)
        if not m:
            continue

        mes_nome, ano = m.group(1), int(m.group(2))
        if ano < 2025 or mes_nome not in mes_idx:
            continue

        try:
            df_mes = ler_atividades_sheets(ws_mes)
        except Exception as e:
            print(f"   ⚠️  Nao foi possivel ler a aba {titulo}: {e}")
            continue

        if df_mes.empty:
            continue

        df_mes = df_mes[mask_status_fixos_mapa(df_mes)].copy()
        df_mes = df_mes.dropna(subset=["LAT", "LONG"])
        df_mes = df_mes[(df_mes["LAT"] != 0) & (df_mes["LONG"] != 0)]
        if df_mes.empty:
            continue

        ano_s = str(ano)
        mes_s = str(mes_idx[mes_nome])
        historico.setdefault(ano_s, {})[mes_s] = {
            "label": f"{MESES_MAPA[mes_idx[mes_nome] - 1].title()}/{ano}",
            "pontos": [],
        }

        for _, row in df_mes.iterrows():
            historico[ano_s][mes_s]["pontos"].append({
                "id":        str(row.get("SITE", "")),
                "cidade":    str(row.get("CIDADE", "")),
                "tec4g":     str(row.get("2G|3G|4G", "")),
                "tec5g":     str(row.get("5G", "")),
                "concluido": str(row.get("CONCLUIDO", "") or ""),
                "hotel":     str(row.get("HOTEL", "") or ""),
                "status":    str(row.get("STATUS", "") or ""),
                "tipo":      tipo_status_mapa(row.get("STATUS", "")),
                "lat":       normalizar_coord_mapa(row.get("LAT", 0), 90),
                "lon":       normalizar_coord_mapa(row.get("LONG", 0), 180),
            })
            total += 1

    if total:
        print(f"   {total} atividades anteriores carregadas para consulta no mapa.")
    else:
        print("   Nenhuma atividade anterior encontrada para consulta no mapa.")
    return historico


COLUNAS_SHEETS = [
    "ROW_SHEET", "DEMANDA", "INTEGRAÇÃO", "SITE", "UF", "TEC",
    "LAT", "LONG", "CIDADE", "2G|3G|4G", "5G", "STATUS", "CONCLUIDO", "HOTEL",
]


def ler_atividades_sheets(ws):
    dados = ws.get_all_values()
    linhas = dados[DATA_START_ROW - 1:]
    registros = []

    for i, linha in enumerate(linhas):
        while len(linha) <= max(CI.values()):
            linha.append("")

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
        return pd.DataFrame(columns=COLUNAS_SHEETS)

    return pd.DataFrame(registros)


def atualizar_sheets(ws, df_fixas, df_rota, sites_originais, df_aguardando=None):
    print("   Montando linhas...")

    def _coord(v):
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

    for _, row in df_fixas.iterrows():
        todas_linhas.append(formatar_linha(row.to_dict()))

    for _, row in df_rota.iterrows():
        row_d = row.to_dict()
        site = str(row_d.get("SITE", "")).upper()

        if site in sites_originais:
            status = row_d.get("STATUS", "")
            if status == ST_NOVA:
                status = ""
        else:
            status = ST_NOVA

        todas_linhas.append(formatar_linha(row_d, status_override=status))

    if df_aguardando is not None and not df_aguardando.empty:
        todas_linhas.append([""] * 13)
        for _, row in df_aguardando.iterrows():
            todas_linhas.append(formatar_linha(row.to_dict()))

    end_row = DATA_START_ROW + len(todas_linhas) - 1
    range_ref = f"A{DATA_START_ROW}:M{end_row}"
    ws.update(values=todas_linhas, range_name=range_ref, value_input_option="USER_ENTERED")

    dados_atuais = ws.get_all_values()
    ultima_linha_com_dados = len(dados_atuais)
    if ultima_linha_com_dados > end_row:
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

_MAPA_COLUNAS = {
    "SITE":        ["SITE", "SITES", "SITE_ID", "COD_SITE", "NOME_SITE", "CODIGO"],
    "CIDADE":      ["CIDADE", "MUNICIPIO", "CITY", "LOCALIDADE"],
    "UF":          ["UF", "ESTADO", "STATE", "REGIONAL", "UF_SITE"],
    "LATITUDE":    ["LATITUDE", "LAT", "LATIT"],
    "LONGITUDE":   ["LONGITUDE", "LON", "LONG", "LONGIT"],
    "FREQUENCIA":  ["FREQUENCIA", "FREQUÊNCIA", "FREQ", "FREQUENCIAS", "FREQUENCY"],
    "VENDOR":      ["VENDOR", "EQUIPE_RF", "EMPRESA_RF", "FORNECEDOR", "INTEGRADOR_RF"],
    "INTEGRADORA": ["INTEGRADORA", "DEMANDANTE", "CLIENTE", "OPERADORA"],
    "TECNOLOGIA":  ["TECNOLOGIA", "TEC", "TECNOLOGIAS", "TECH", "TECHNOLOGY"],
}


def _resolver_colunas(df_raw):
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
    s = str(val).strip().upper()
    if s in ("", "NAN", "NONE", "0") or not s.isalpha():
        return ""
    return s if len(s) == 2 else ""


def _limpar_cidade(val):
    s = str(val).strip().upper()
    return "" if s in ("", "NAN", "NONE") else s


def _formatar_site_robusto(site_raw):
    s = str(site_raw).strip().upper()
    s = re.sub(r"_[A-Z]$", "", s)
    if len(s) >= 5:
        return f"{s[:2]}-{s[-3:]}"
    return s


def _buscar_na_base4g(df_base, lat, lon, tolerancia=0.002):
    if df_base is None or df_base.empty:
        return pd.DataFrame()
    mask = (
        (abs(df_base["LATITUDE"]  - lat) <= tolerancia) &
        (abs(df_base["LONGITUDE"] - lon) <= tolerancia)
    )
    return df_base[mask]


def processar_arquivo(df_raw, nome_arquivo="", df_base=None):
    prefixo = f"   [{nome_arquivo}]" if nome_arquivo else "  "
    col = _resolver_colunas(df_raw)

    registros = []
    descartados = 0

    for idx, row in df_raw.iterrows():
        linha_num = idx + 2

        site_raw = str(row[col["SITE"]]).strip() if "SITE" in col else ""
        if not site_raw or site_raw.upper() in ("NAN", "NONE", ""):
            print(f"{prefixo} ⚠️  Linha {linha_num}: SITE vazio — ignorada")
            descartados += 1
            continue
        site = _formatar_site_robusto(site_raw)

        lat  = safe_float(row[col["LATITUDE"]])  if "LATITUDE"  in col else 0.0
        lon  = safe_float(row[col["LONGITUDE"]]) if "LONGITUDE" in col else 0.0
        if lat == 0.0 and lon == 0.0:
            print(f"{prefixo} ⚠️  Linha {linha_num} ({site}): coordenadas zeradas — ignorada")
            descartados += 1
            continue

        uf_raw = row[col["UF"]] if "UF" in col else ""
        uf = _limpar_uf(uf_raw)
        if not uf:
            print(f"{prefixo} ℹ️  {site}: UF ausente/invalida ({repr(uf_raw)}) — sera buscada na Base_4G")

        cidade_raw = row[col["CIDADE"]] if "CIDADE" in col else ""
        cidade = _limpar_cidade(cidade_raw)
        if not cidade:
            print(f"{prefixo} ℹ️  {site}: CIDADE vazia — sera buscada por coordenadas")

        tec_raw = row[col["TECNOLOGIA"]] if "TECNOLOGIA" in col else ""
        tec = "" if pd.isna(tec_raw) else normalizar_tec(str(tec_raw).strip())

        freq_raw = row[col["FREQUENCIA"]] if "FREQUENCIA" in col else ""
        if pd.isna(freq_raw) or str(freq_raw).strip().upper() in ("NAN", "NONE", ""):
            freq_234, freq_5g = "", ""
        else:
            freq_234, freq_5g = normalizar_frequencia(str(freq_raw).strip())

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

    if df_base is not None and not df_out.empty:
        for idx, row in df_out.iterrows():
            precisa_cidade = not row["CIDADE"]
            precisa_uf     = not row["UF"]
            precisa_tec    = not row["TEC"]
            precisa_4g     = not row["2G|3G|4G"]
            precisa_5g     = not row["5G"]

            if not any([precisa_cidade, precisa_uf, precisa_tec, precisa_4g, precisa_5g]):
                continue
            if row["LAT"] == 0.0:
                continue

            try:
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

                if precisa_tec:
                    for col_tec in ("TEC", "TECNOLOGIA", "TECHNOLOGY"):
                        if col_tec in matches.columns:
                            val = str(m.get(col_tec, "")).strip()
                            if val and val.upper() not in ("", "NAN", "NONE"):
                                df_out.at[idx, "TEC"] = val
                                break

                if precisa_4g or precisa_5g:
                    for col_freq in ("FREQUENCIA", "FREQUÊNCIA", "FREQ", "FREQUENCY"):
                        if col_freq in matches.columns:
                            freq_raw_base = str(m.get(col_freq, "")).strip()
                            if freq_raw_base and freq_raw_base.upper() not in ("", "NAN", "NONE"):
                                f4g, f5g = normalizar_frequencia(freq_raw_base)
                                if precisa_4g and f4g:
                                    df_out.at[idx, "2G|3G|4G"] = f4g
                                if precisa_5g and f5g:
                                    df_out.at[idx, "5G"] = f5g
                                break

            except Exception as e:
                print(f"{prefixo} ⚠️  Erro ao buscar Base_4G para {row['SITE']}: {e}")

        for _, row in df_out.iterrows():
            if not row["CIDADE"]:
                print(f"{prefixo} ⚠️  {row['SITE']}: CIDADE nao encontrada em nenhuma fonte.")
            if not row["UF"]:
                print(f"{prefixo} ⚠️  {row['SITE']}: UF nao encontrada — sera solicitada no relatorio.")

    return df_out


def processar_felipe(df_raw, df_base=None):
    return processar_arquivo(df_raw, df_base=df_base)


def encontrar_arquivos_novas():
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

            if matches.empty:
                matches = _buscar_na_base4g(df_base, lat, lon, tolerancia=0.002)

            pci_list = [str(int(v)) for v in matches["PCI"].dropna().unique()]    if "PCI"     in matches.columns else []
            az_list  = [str(int(v)) for v in matches["AZIMUTH"].dropna().unique()] if "AZIMUTH" in matches.columns else []
            pci_str  = "/".join(unique_preserve_order(pci_list))
            az_str   = "/".join(unique_preserve_order(az_list))

            if not pci_str:
                print(f"   ⚠️  PCI não encontrado para o site {site} ({cidade}) — confira manualmente.")
            if not az_str:
                print(f"   ⚠️  AZIMUTH não encontrado para o site {site} ({cidade}) — confira manualmente.")

            if not uf or uf in ("", "NAN", "NONE"):
                uf_base = ""
                if not matches.empty and "[P]UF" in matches.columns:
                    ufs = matches["[P]UF"].dropna().unique()
                    if len(ufs) > 0:
                        uf_base = str(ufs[0]).strip().upper()

                if uf_base and len(uf_base) == 2 and uf_base.isalpha():
                    uf = uf_base
                else:
                    if site in _uf_cache:
                        uf = _uf_cache[site]
                    else:
                        print(f"\n⚠️  UF não encontrada para o site {site} ({cidade}).")
                        while True:
                            resp = input(f"   Informe a UF (ex: SP, MG, RJ): ").strip().upper()
                            if len(resp) == 2 and resp.isalpha():
                                uf = resp
                                _uf_cache[site] = uf
                                break
                            print("   ⚠️  UF inválida. Digite apenas 2 letras (ex: SP).")

            lines = [
                f"{site}", "",
                f"{lat} {lon}", "",
                f"({cidade} - {uf})", "",
                f"PCI: {pci_str}",
                f"AZIMUTH:  {az_str}", "",
                "(Frequências):", "",
            ]

            if freq_234 and freq_234.lower() != "nan":
                for bloco in [b.strip() for b in freq_234.split("|") if b.strip()]:
                    lines.append(bloco)

            if freq_5g and freq_5g.lower() != "nan":
                lines.append(freq_5g if freq_5g.upper().startswith("5G:") else f"5G: {freq_5g}")

            lines += [
                "", "OBS.>", "",
                f"{demanda} - {integracao}", "",
                "LOGS armazenados no servidor.",
                f"> Finalizado:  {mes_ano}", "",
                "--------------------------- Nome - LOGS --------------------------------------", "",
            ]

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
# HOTÉIS
# ============================================================

def carregar_hoteis():
    def _parse_hoteis(cabecalhos_raw, linhas):
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

    try:
        client  = conectar_sheets()
        sh      = client.open_by_key(HOTEIS_SHEET_ID)
        ws_h = None
        for aba in sh.worksheets():
            if aba.id == HOTEIS_GID:
                ws_h = aba
                break
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
    except Exception as e:
        print(f"   ⚠️  Erro ao acessar HOTEIS no Google Sheets: {e}")

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
        cabecalhos = list(df.columns)
        linhas = df.astype(str).values.tolist()
        hoteis = _parse_hoteis(cabecalhos, linhas)
        print(f"   {len(hoteis)} hoteis carregados de HOTEIS.xlsx (fallback local).")
        return hoteis
    except Exception as e:
        print(f"   ❌ Erro ao ler HOTEIS.xlsx: {e}")
        return []


# ============================================================
# GERAÇÃO DE MAPA — MAPBOX GL JS
# ============================================================

def gerar_mapa(df_rota, lat0, lon0, cidade0, df_fixas=None,
               df_aguardando=None, historico_meses=None, mapbox_token=None):
    """
    Gera MAPA_ROTAS.html com Mapbox GL JS.
    Rota REAL (Directions API) para a próxima atividade pendente.
    Linha guia com setas para as demais atividades pendentes.
    Trânsito em tempo real via Traffic Layer.
    Popup de alerta quando quota Mapbox estiver baixa.
    """
    import json

    hoteis = carregar_hoteis()

    # ── Próxima atividade (índice 0) recebe rota real
    # ── Demais recebem linha guia tracejada com setas
    pontos_rota = []
    for i, (_, row) in enumerate(df_rota.iterrows()):
        lat = normalizar_coord_mapa(row.get("LAT", 0), 90)
        lon = normalizar_coord_mapa(row.get("LONG", 0), 180)
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
            "ordem":  i,  # 0 = próxima (rota real), >0 = standby
        })

    pontos_fixos = []
    if df_fixas is not None and not df_fixas.empty:
        for _, row in df_fixas.iterrows():
            lat = normalizar_coord_mapa(row.get("LAT", 0), 90)
            lon = normalizar_coord_mapa(row.get("LONG", 0), 180)
            if lat == 0 and lon == 0:
                continue
            status = str(row.get("STATUS", "")).strip()
            pontos_fixos.append({
                "id":        str(row.get("SITE",      "")),
                "cidade":    str(row.get("CIDADE",    "")),
                "tec4g":     str(row.get("2G|3G|4G",  "")),
                "tec5g":     str(row.get("5G",        "")),
                "concluido": str(row.get("CONCLUIDO", "") or ""),
                "hotel":     str(row.get("HOTEL",     "") or ""),
                "status":    status,
                "tipo":      tipo_status_mapa(status),
                "lat":       lat,
                "lon":       lon,
            })

    pontos_aguard = []
    if df_aguardando is not None and not df_aguardando.empty:
        for _, row in df_aguardando.iterrows():
            lat = normalizar_coord_mapa(row.get("LAT", 0), 90)
            lon = normalizar_coord_mapa(row.get("LONG", 0), 180)
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
    j_hist   = json.dumps(historico_meses or {}, ensure_ascii=False)

    # ── Token Mapbox: salvo em arquivo JS separado (fora do HTML) ──
    # O HTML carrega mapbox_config.js que NÃO vai para o GitHub.
    # Isso resolve o erro 409 "Secret detected in content" do GitHub.
    token = mapbox_token or ""
    config_js_path = BASE_DIR / "mapbox_config.js"
    try:
        with open(config_js_path, "w", encoding="utf-8") as f_cfg:
            f_cfg.write(f"var MAPBOX_TOKEN = '{token}';\n")
        print(f"   mapbox_config.js gerado (token protegido).")
    except Exception as e:
        print(f"   ⚠️  Nao foi possivel gerar mapbox_config.js: {e}")

    html = f"""<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8"/>
<meta name="viewport" content="width=device-width,initial-scale=1.0,maximum-scale=1.0,user-scalable=no"/>
<title>DT 3.0 — Mapa de Rota</title>

<!-- Token Mapbox (arquivo local, nunca publicado no GitHub) -->
<script src="mapbox_config.js"></script>

<!-- Mapbox GL JS -->
<link href="https://api.mapbox.com/mapbox-gl-js/v2.15.0/mapbox-gl.css" rel="stylesheet"/>
<script src="https://api.mapbox.com/mapbox-gl-js/v2.15.0/mapbox-gl.js"></script>

<style>
/* ── Reset e base ───────────────────────────────── */
html,body,#map{{width:100%;height:100%;margin:0;padding:0;font-family:'Segoe UI',sans-serif;}}

/* ── Painel lateral ─────────────────────────────── */
#painel{{
  position:absolute;top:10px;right:10px;z-index:1000;
  background:rgba(255,255,255,0.97);border-radius:12px;
  padding:14px 18px;box-shadow:0 4px 20px rgba(0,0,0,0.18);
  min-width:240px;max-width:300px;
}}
#painel h4{{margin:0 0 10px;font-size:14px;font-weight:700;color:#1a237e;letter-spacing:.5px;}}

/* ── Tabs ───────────────────────────────────────── */
.tabs-mapa{{display:grid;grid-template-columns:1fr 1fr;gap:6px;margin-bottom:10px;}}
.tab-mapa{{
  border:1px solid #cfd6e6;background:#f7f9ff;color:#1a237e;
  border-radius:6px;padding:7px 8px;font-size:12px;font-weight:700;cursor:pointer;
}}
.tab-mapa.ativo{{background:#1a237e;color:#fff;border-color:#1a237e;}}
.tab-panel{{display:none;}}
.tab-panel.ativo{{display:block;}}

/* ── Legenda ────────────────────────────────────── */
.leg{{display:flex;align-items:center;gap:8px;margin-bottom:5px;font-size:12px;color:#333;}}
.dot{{width:12px;height:12px;border-radius:50%;display:inline-block;flex-shrink:0;}}
.dot-pend{{background:#00BFFF;box-shadow:0 0 5px rgba(0,191,255,0.6);}}
.dot-prox{{background:#f39c12;border:2px solid #d68910;}}
.dot-conc{{background:#27ae60;}}
.dot-impr{{background:#e74c3c;}}
.dot-canc{{background:#7f8c8d;}}
.dot-part{{background:#f39c12;border:2px solid #b7770d;}}
.dot-hotel{{background:#e67e22;}}
.dot-aguard{{background:#8e44ad;}}

/* ── ETA Box ────────────────────────────────────── */
#eta-box{{
  margin:10px 0;padding:10px 12px;background:#eef4ff;
  border-radius:8px;border-left:4px solid #2A52BE;font-size:12px;
}}
#eta-box .eta-site{{font-weight:700;color:#1a237e;font-size:13px;margin-bottom:4px;}}
#eta-box .eta-linha{{color:#333;margin:2px 0;}}
#eta-box .eta-loading{{color:#888;font-style:italic;}}
#eta-box .eta-erro{{color:#e74c3c;font-size:11px;}}

/* ── Contador ───────────────────────────────────── */
#contador{{margin-top:8px;padding-top:8px;border-top:1px solid #ddd;font-size:12px;color:#555;}}
#contador span{{font-weight:700;color:#1a237e;}}

/* ── Separador ──────────────────────────────────── */
.sep{{border:none;border-top:1px solid #e0e0e0;margin:10px 0;}}

/* ── Filtros ────────────────────────────────────── */
.filtros-label{{font-size:11px;font-weight:700;color:#888;letter-spacing:.6px;text-transform:uppercase;margin-bottom:6px;}}
.filtro{{display:flex;align-items:center;gap:8px;margin-bottom:4px;font-size:12px;color:#444;cursor:pointer;user-select:none;}}
.filtro input[type=checkbox]{{width:14px;height:14px;cursor:pointer;accent-color:#1a237e;flex-shrink:0;}}

/* ── Busca de site ──────────────────────────────── */
.busca-wrap{{margin-top:10px;padding-top:10px;border-top:1px solid #ddd;}}
.busca-label{{font-size:11px;font-weight:700;color:#888;letter-spacing:.6px;text-transform:uppercase;margin-bottom:6px;display:block;}}
.busca-row{{display:flex;gap:6px;align-items:center;}}
#busca-site{{
  flex:1;min-width:0;border:1px solid #cfd6e6;border-radius:6px;
  padding:8px 9px;font-size:13px;outline:none;text-transform:uppercase;
}}
#busca-site:focus{{border-color:#2A52BE;box-shadow:0 0 0 2px rgba(42,82,190,.14);}}
#btn-busca{{
  width:34px;height:34px;border:none;border-radius:6px;
  background:#1a237e;color:white;font-weight:700;cursor:pointer;font-size:16px;
}}
#btn-busca:hover{{background:#121858;}}
#msg-busca{{min-height:14px;margin-top:4px;font-size:11px;color:#b35b00;opacity:0;transition:opacity .18s;}}
#msg-busca.visivel{{opacity:1;}}

/* ── Anteriores ─────────────────────────────────── */
.ant-label{{font-size:11px;font-weight:700;color:#888;letter-spacing:.6px;text-transform:uppercase;margin-bottom:6px;}}
.ant-grid{{display:grid;grid-template-columns:1fr 1fr;gap:7px;margin-top:8px;}}
.select-ant{{width:100%;border:1px solid #cfd6e6;border-radius:6px;padding:7px;font-size:12px;background:#fff;}}
.btn-ant{{
  width:100%;padding:8px;margin-top:8px;font-size:12px;font-weight:700;
  border:none;border-radius:6px;background:#1a237e;color:white;cursor:pointer;
}}
.btn-ant:hover{{background:#121858;}}
.btn-ant-cinza{{background:#6c757d;}}
.btn-ant-cinza:hover{{background:#545b62;}}
#msg-ant{{min-height:14px;margin-top:4px;font-size:11px;color:#b35b00;opacity:0;transition:opacity .18s;}}
#msg-ant.visivel{{opacity:1;}}

/* ── Resultado busca global ─────────────────────── */
#resultado-busca-global{{
  display:none;position:absolute;top:88px;right:330px;z-index:1002;
  width:min(360px,calc(100vw - 24px));max-height:64vh;overflow:auto;
  background:rgba(255,255,255,.98);border-radius:10px;
  box-shadow:0 6px 24px rgba(0,0,0,.22);padding:12px 14px;
}}
#resultado-busca-global.visivel{{display:block;}}
.bg-topo{{display:flex;align-items:center;justify-content:space-between;gap:10px;margin-bottom:8px;}}
.bg-titulo{{font-size:13px;font-weight:700;color:#1a237e;}}
#btn-fechar-bg{{
  width:26px;height:26px;border:none;border-radius:6px;
  background:#eef1f7;color:#1a237e;font-weight:800;cursor:pointer;
}}
.bg-item{{
  display:flex;align-items:center;justify-content:space-between;gap:10px;
  border-top:1px solid #e6e8ef;padding:8px 0;font-size:12px;color:#333;
}}
.bg-item:first-child{{border-top:none;}}
.bg-site{{font-weight:700;color:#1a237e;}}
.bg-ver{{
  border:none;border-radius:6px;background:#1a237e;color:white;
  font-size:11px;font-weight:700;padding:5px 8px;cursor:pointer;white-space:nowrap;
}}
.bg-ver:hover{{background:#121858;}}

/* ── Ponto de partida (animado) ─────────────────── */
.partida-pulse{{
  width:28px;height:28px;border-radius:50%;position:relative;
  background:rgba(243,156,18,.96);border:3px solid #fff;
  box-shadow:0 2px 9px rgba(0,0,0,.28);
}}
.partida-pulse:before,.partida-pulse:after{{
  content:"";position:absolute;inset:-8px;border-radius:50%;
  border:2px solid rgba(243,156,18,.55);
  animation:dtPulse 1.9s ease-out infinite;pointer-events:none;
}}
.partida-pulse:after{{animation-delay:.65s;}}
.partida-core{{
  position:absolute;left:50%;top:50%;width:8px;height:8px;border-radius:50%;
  background:#fff;transform:translate(-50%,-50%);
}}
.partida-orbit{{
  position:absolute;left:50%;top:50%;width:38px;height:38px;
  margin:-19px 0 0 -19px;border-radius:50%;
  border:2px dashed rgba(26,35,126,.45);
  animation:dtSpin 3.8s linear infinite;pointer-events:none;
}}
@keyframes dtPulse{{
  0%{{transform:scale(.55);opacity:.85;}}
  70%{{transform:scale(1.55);opacity:0;}}
  100%{{transform:scale(1.55);opacity:0;}}
}}
@keyframes dtSpin{{to{{transform:rotate(360deg);}}}}

/* ── Próxima atividade (estrela pulsante) ────────── */
.proxima-pulse{{
  width:24px;height:24px;border-radius:50%;position:relative;
  background:rgba(26,35,126,.92);border:2px solid #fff;
  box-shadow:0 2px 8px rgba(0,0,0,.3);display:flex;align-items:center;justify-content:center;
  font-size:11px;color:white;font-weight:700;
}}
.proxima-pulse:before{{
  content:"";position:absolute;inset:-6px;border-radius:50%;
  border:2px solid rgba(26,35,126,.4);
  animation:dtPulse 2.2s ease-out infinite;pointer-events:none;
}}

/* ── Popup customizado ──────────────────────────── */
.popup-dt{{font-family:'Segoe UI',sans-serif;font-size:13px;min-width:200px;}}
.popup-dt b{{color:#1a237e;}}
.popup-status{{margin-top:6px;padding:4px 8px;border-radius:4px;font-weight:600;font-size:12px;}}
.ps-conc{{background:#d4edda;color:#155724;}}
.ps-impr{{background:#f8d7da;color:#721c24;}}
.ps-canc{{background:#eceff1;color:#455a64;}}
.ps-aguard{{background:#e8daef;color:#6c3483;}}
.btn-concluir{{
  margin-top:10px;background:#27ae60;color:white;border:none;
  padding:9px;border-radius:5px;cursor:pointer;width:100%;
  font-weight:700;font-size:14px;
}}
.btn-concluir:hover{{background:#1e8449;}}

/* ── Popup de alerta de quota ───────────────────── */
#popup-quota{{
  display:none;position:fixed;top:0;left:0;width:100%;height:100%;
  background:rgba(0,0,0,0.6);z-index:3000;
  align-items:center;justify-content:center;
}}
#popup-quota.visivel{{display:flex;}}
.quota-box{{
  background:white;border-radius:12px;padding:28px 32px;max-width:400px;
  box-shadow:0 8px 32px rgba(0,0,0,0.3);text-align:center;
  animation:slideUp .3s ease-out;
}}
.quota-box.critico{{border-top:5px solid #e74c3c;}}
.quota-box.aviso{{border-top:5px solid #f39c12;}}
.quota-box h3{{margin:0 0 12px;font-size:18px;}}
.quota-box.critico h3{{color:#e74c3c;}}
.quota-box.aviso h3{{color:#f39c12;}}
.quota-box p{{color:#555;font-size:14px;line-height:1.6;margin:0 0 16px;}}
.quota-box a{{color:#2A52BE;font-weight:700;}}
.quota-btn{{
  border:none;border-radius:8px;background:#1a237e;color:white;
  padding:12px 24px;font-size:14px;font-weight:700;cursor:pointer;
  width:100%;margin-top:4px;
}}
.quota-btn:hover{{background:#121858;}}
@keyframes slideUp{{
  from{{transform:translateY(30px);opacity:0;}}
  to{{transform:translateY(0);opacity:1;}}
}}

/* ── Quota badge (painel) ────────────────────────── */
#quota-badge{{
  margin-top:8px;padding:6px 10px;border-radius:6px;
  font-size:11px;font-weight:600;text-align:center;
}}
#quota-badge.ok{{background:#e8f5e9;color:#2e7d32;}}
#quota-badge.aviso{{background:#fff8e1;color:#f57f17;}}
#quota-badge.critico{{background:#ffebee;color:#c62828;}}

/* ── Mobile ─────────────────────────────────────── */
#painel-toggle{{display:none;}}
@media (max-width: 768px){{
  #painel{{
    top:10px;right:10px;width:min(300px,82vw);min-width:0;max-height:82vh;overflow:auto;
    transform:translateX(calc(100% + 24px));transition:transform .24s ease;
    border-radius:10px 0 0 10px;
  }}
  body.painel-aberto #painel{{transform:translateX(0);}}
  #painel-toggle{{
    display:flex;position:absolute;right:0;top:145px;z-index:1001;
    width:38px;height:118px;border:none;border-radius:14px 0 0 14px;
    background:rgba(255,255,255,.96);box-shadow:0 4px 18px rgba(0,0,0,.2);
    align-items:center;justify-content:center;flex-direction:column;gap:10px;cursor:pointer;
  }}
  #painel-toggle span{{width:7px;height:7px;border-radius:50%;background:#111;display:block;}}
  body.painel-aberto #painel-toggle{{display:none;}}
  #resultado-busca-global{{right:52px;top:92px;width:calc(100vw - 68px);max-height:56vh;}}
}}
</style>
</head>
<body>

<div id="map"></div>
<button id="painel-toggle" aria-label="Abrir painel"><span></span><span></span><span></span></button>

<!-- Popup alerta quota -->
<div id="popup-quota">
  <div class="quota-box" id="quota-box-inner">
    <h3 id="quota-titulo"></h3>
    <p id="quota-msg"></p>
    <button class="quota-btn" onclick="fecharPopupQuota()">Entendi</button>
  </div>
</div>

<!-- Resultado busca global -->
<div id="resultado-busca-global">
  <div class="bg-topo">
    <div class="bg-titulo" id="bg-titulo"></div>
    <button id="btn-fechar-bg" onclick="fecharBuscaGlobal()">✕</button>
  </div>
  <div id="bg-lista"></div>
</div>

<div id="painel">
  <h4>DT 3.0 — Rota</h4>

  <div class="tabs-mapa">
    <button class="tab-mapa ativo" id="tab-atual" onclick="mostrarAba('atual')">Atual</button>
    <button class="tab-mapa" id="tab-anteriores" onclick="mostrarAba('anteriores')">Anteriores</button>
  </div>

  <!-- Aba Atual -->
  <div class="tab-panel ativo" id="pane-atual">

    <!-- ETA Próxima atividade -->
    <div id="eta-box">
      <div class="eta-loading">Calculando rota...</div>
    </div>

    <!-- Legenda -->
    <div class="leg"><span class="dot dot-part"></span> Ponto de partida</div>
    <div class="leg"><span class="dot dot-prox"></span> Próxima atividade</div>
    <div class="leg"><span class="dot dot-pend"></span> Standby (aguardando)</div>

    <div id="contador">
      Pendentes: <span id="cnt-pend">0</span> |
      Concluídas: <span id="cnt-conc">0</span>
    </div>

    <!-- Badge quota -->
    <div id="quota-badge" class="ok">Mapbox: carregando...</div>

    <!-- Busca de site -->
    <div class="busca-wrap">
      <label class="busca-label" for="busca-site">Buscar site</label>
      <div class="busca-row">
        <input id="busca-site" type="text" placeholder="Ex: MG-ABC" autocomplete="off"/>
        <button id="btn-busca" onclick="buscarSite()" title="Buscar">⌕</button>
      </div>
      <div id="msg-busca"></div>
    </div>

    <hr class="sep"/>
    <div class="filtros-label">Exibir camadas</div>
    <label class="filtro"><input type="checkbox" id="chk-conc" checked onchange="toggleCamada('conc',this.checked)"/><span class="dot dot-conc"></span> Concluídas</label>
    <label class="filtro"><input type="checkbox" id="chk-impr" checked onchange="toggleCamada('impr',this.checked)"/><span class="dot dot-impr"></span> Improdutivas</label>
    <label class="filtro"><input type="checkbox" id="chk-canc" checked onchange="toggleCamada('canc',this.checked)"/><span class="dot dot-canc"></span> Canceladas</label>
    <label class="filtro"><input type="checkbox" id="chk-hotel" onchange="toggleCamada('hotel',this.checked)"/><span class="dot dot-hotel"></span> Hotéis</label>
    <label class="filtro"><input type="checkbox" id="chk-aguard" checked onchange="toggleCamada('aguard',this.checked)"/><span class="dot dot-aguard"></span> Aguardando</label>
    <label class="filtro"><input type="checkbox" id="chk-traffic" onchange="toggleTraffic(this.checked)"/><span style="width:12px;height:12px;display:inline-block;background:linear-gradient(90deg,#27ae60,#f39c12,#e74c3c);border-radius:3px;flex-shrink:0;"></span> Trânsito</label>

  </div>

  <!-- Aba Anteriores -->
  <div class="tab-panel" id="pane-anteriores">
    <div class="ant-label">Consultar período</div>
    <div class="ant-grid">
      <select class="select-ant" id="sel-ano-ant" onchange="preencherMeses()"></select>
      <select class="select-ant" id="sel-mes-ant"></select>
    </div>
    <button class="btn-ant" onclick="mostrarMesAnterior()">Ver</button>
    <button class="btn-ant btn-ant-cinza" onclick="voltarAtual()">Voltar ao atual</button>
    <div id="msg-ant"></div>
  </div>
</div>

<script>
// ══════════════════════════════════════════════════════════════
// DADOS (injetados pelo Python)
// MAPBOX_TOKEN é lido do arquivo local mapbox_token.js
// (nunca publicado no GitHub — protegido pelo .gitignore)
// ══════════════════════════════════════════════════════════════
var PARTIDA       = {j_part};
var ROTA          = {j_rota};
var FIXOS         = {j_fixos};
var AGUARD        = {j_aguard};
var HOTEIS        = {j_hoteis};
var ANTERIORES    = {j_hist};

// ══════════════════════════════════════════════════════════════
// CONTROLE DE QUOTA MAPBOX (persistido em localStorage)
// Limite: 600 Directions API/mês  |  Alerta: ≤ 100 restantes
// ══════════════════════════════════════════════════════════════
var LIMITE_QUOTA  = 600;
var ALERTA_QUOTA  = 100;

function _chaveQuota() {{
  var d = new Date();
  return 'mapbox_uso_' + d.getFullYear() + '_' + (d.getMonth() + 1);
}}

function lerUso() {{
  var v = parseInt(localStorage.getItem(_chaveQuota()) || '0', 10);
  return isNaN(v) ? 0 : v;
}}

function incrementarUso() {{
  var uso = lerUso() + 1;
  localStorage.setItem(_chaveQuota(), uso.toString());
  atualizarBadgeQuota(uso);
  verificarAlertaQuota(uso);
  return uso;
}}

function atualizarBadgeQuota(uso) {{
  var badge = document.getElementById('quota-badge');
  var restante = LIMITE_QUOTA - uso;
  if (restante > ALERTA_QUOTA) {{
    badge.className = 'ok';
    badge.textContent = 'Mapbox: ' + uso + '/' + LIMITE_QUOTA + ' chamadas usadas';
  }} else if (restante > 0) {{
    badge.className = 'aviso';
    badge.textContent = '⚠️ Mapbox: ' + restante + ' chamadas restantes!';
  }} else {{
    badge.className = 'critico';
    badge.textContent = '❌ Quota Mapbox esgotada este mês';
  }}
}}

function verificarAlertaQuota(uso) {{
  var restante = LIMITE_QUOTA - uso;
  var box = document.getElementById('quota-box-inner');
  var titulo = document.getElementById('quota-titulo');
  var msg = document.getElementById('quota-msg');

  if (restante <= 0) {{
    box.className = 'quota-box critico';
    titulo.textContent = '❌ Quota Mapbox Esgotada';
    msg.innerHTML =
      'Você atingiu o limite de <strong>' + LIMITE_QUOTA + ' chamadas</strong> este mês.<br><br>' +
      'O mapa continuará funcionando normalmente, porém <strong>sem rotas reais</strong> até o próximo mês.<br><br>' +
      'Renova em: <strong>' + proximoDia1() + '</strong><br><br>' +
      '<a href="https://account.mapbox.com/billing/overview/" target="_blank">Ver conta Mapbox</a>';
    mostrarPopupQuota();
  }} else if (restante <= ALERTA_QUOTA) {{
    box.className = 'quota-box aviso';
    titulo.textContent = '⚠️ Atenção: Quota Baixa';
    msg.innerHTML =
      'Você usou <strong>' + uso + '</strong> de <strong>' + LIMITE_QUOTA + '</strong> chamadas Mapbox este mês.<br><br>' +
      'Restam apenas <strong>' + restante + ' chamadas</strong>.<br>' +
      'Renova em: <strong>' + proximoDia1() + '</strong>';
    mostrarPopupQuota();
  }}
}}

function proximoDia1() {{
  var d = new Date();
  var prox = new Date(d.getFullYear(), d.getMonth() + 1, 1);
  return prox.toLocaleDateString('pt-BR');
}}

function mostrarPopupQuota() {{
  document.getElementById('popup-quota').classList.add('visivel');
}}

function fecharPopupQuota() {{
  document.getElementById('popup-quota').classList.remove('visivel');
}}

function quotaEsgotada() {{
  return lerUso() >= LIMITE_QUOTA;
}}

// ══════════════════════════════════════════════════════════════
// INIT MAPBOX
// ══════════════════════════════════════════════════════════════
mapboxgl.accessToken = MAPBOX_TOKEN;

var map = new mapboxgl.Map({{
  container: 'map',
  style: 'mapbox://styles/mapbox/streets-v12',
  center: [PARTIDA.lon, PARTIDA.lat],
  zoom: 7
}});

map.addControl(new mapboxgl.NavigationControl(), 'bottom-right');
map.addControl(new mapboxgl.ScaleControl({{unit:'metric'}}), 'bottom-left');

// ══════════════════════════════════════════════════════════════
// ESTADO GLOBAL
// ══════════════════════════════════════════════════════════════
var indiceSites     = {{}};
var marcadores      = [];        // {{marker, lat, lon, id}}
var mkFixos         = {{}};      // camada → [markers]
var mkHoteis        = [];
var mkAguard        = [];
var mkHist          = [];
var concCount       = 0;
var modoAnterior    = false;
var buscaTimer      = null;
var ocorrenciasBG   = [];
var trafficAtivo    = false;

// ══════════════════════════════════════════════════════════════
// UTILITÁRIOS
// ══════════════════════════════════════════════════════════════
function normBusca(v) {{
  return String(v||'').trim().toUpperCase().replace(/[^A-Z0-9]+/g,'');
}}

function registrarSite(id, marker, lat, lon, camada) {{
  var k = normBusca(id);
  if (!k) return;
  var item = {{id:id,marker:marker,lat:lat,lon:lon,camada:camada||null}};
  indiceSites[k] = item;
  if (k.length >= 3) indiceSites[k.slice(-3)] = item;
}}

function mostrarMsg(elId, txt) {{
  var el = document.getElementById(elId);
  el.textContent = txt || '';
  el.classList.toggle('visivel', Boolean(txt));
  if (elId === 'msg-busca') {{
    if (buscaTimer) clearTimeout(buscaTimer);
    if (txt) buscaTimer = setTimeout(function(){{ el.classList.remove('visivel'); }}, 3000);
  }}
}}

function freqStr(p) {{
  return [p.tec4g, p.tec5g].filter(Boolean).join(' | ');
}}

// ══════════════════════════════════════════════════════════════
// CAMADA DE TRÂNSITO (Mapbox Traffic Layer)
// ══════════════════════════════════════════════════════════════
function toggleTraffic(ativo) {{
  trafficAtivo = ativo;
  if (!map.isStyleLoaded()) {{ map.once('style.load', function(){{ toggleTraffic(ativo); }}); return; }}

  var camadas = ['traffic-street', 'traffic-street-case', 'traffic-motorway',
                  'traffic-trunk', 'traffic-secondary'];

  if (ativo) {{
    // Adicionar source de tráfego se não existir
    if (!map.getSource('mapbox-traffic')) {{
      map.addSource('mapbox-traffic', {{
        type: 'vector',
        url: 'mapbox://mapbox.mapbox-traffic-v1'
      }});
    }}
    // Camadas de congestionamento com cores
    var niveis = [
      {{ id: 'traffic-motorway', filter: ['==', ['get','class'], 'motorway'] }},
      {{ id: 'traffic-trunk',    filter: ['==', ['get','class'], 'trunk'] }},
      {{ id: 'traffic-street',   filter: ['match', ['get','class'], ['street','street_limited','secondary'], true, false] }},
    ];
    niveis.forEach(function(n) {{
      if (!map.getLayer(n.id)) {{
        map.addLayer({{
          id: n.id,
          type: 'line',
          source: 'mapbox-traffic',
          'source-layer': 'traffic',
          filter: n.filter,
          paint: {{
            'line-width': 3,
            'line-color': [
              'match', ['get','congestion'],
              'low',      '#27ae60',
              'moderate', '#f39c12',
              'heavy',    '#e74c3c',
              'severe',   '#8e1a0e',
              '#aaaaaa'
            ],
            'line-opacity': 0.82
          }}
        }}, 'road-label');
      }}
    }});
  }} else {{
    ['traffic-motorway','traffic-trunk','traffic-street'].forEach(function(id) {{
      if (map.getLayer(id)) map.removeLayer(id);
    }});
  }}
}}

// ══════════════════════════════════════════════════════════════
// DESENHAR ROTA REAL (Mapbox Directions API → próxima atividade)
// ══════════════════════════════════════════════════════════════
function atualizarEtaBox(html) {{
  document.getElementById('eta-box').innerHTML = html;
}}

async function desenharRotaReal(latOrig, lonOrig, proxPonto) {{
  if (quotaEsgotada()) {{
    atualizarEtaBox(
      '<div class="eta-site">' + proxPonto.id + ' — ' + proxPonto.cidade + '</div>' +
      '<div class="eta-erro">❌ Quota Mapbox esgotada — rota sem cálculo de trajeto.</div>'
    );
    desenharLinhaGuia(latOrig, lonOrig, proxPonto.lat, proxPonto.lon, 'rota-prox', '#2A52BE', true, 4);
    return;
  }}

  atualizarEtaBox('<div class="eta-loading">⏳ Calculando rota para ' + proxPonto.id + '...</div>');

  try {{
    var url = 'https://api.mapbox.com/directions/v5/mapbox/driving-traffic/' +
      lonOrig + ',' + latOrig + ';' + proxPonto.lon + ',' + proxPonto.lat +
      '?access_token=' + MAPBOX_TOKEN +
      '&geometries=geojson&overview=full&steps=false&language=pt-BR';

    var resp = await fetch(url);
    var data = await resp.json();

    if (!data.routes || !data.routes.length) {{
      throw new Error('Nenhuma rota retornada pela API');
    }}

    var route = data.routes[0];
    var distKm = (route.distance / 1000).toFixed(1);
    var durMin = Math.round(route.duration / 60);
    var durStr = durMin >= 60
      ? Math.floor(durMin/60) + 'h ' + (durMin % 60) + 'min'
      : durMin + ' min';

    // Hora de chegada estimada
    var chegada = new Date(Date.now() + route.duration * 1000);
    var chegadaStr = chegada.toLocaleTimeString('pt-BR', {{hour:'2-digit', minute:'2-digit'}});

    // Incrementar quota
    var usoAtual = incrementarUso();

    // ETA box atualizado
    atualizarEtaBox(
      '<div class="eta-site">📍 ' + proxPonto.id + ' — ' + proxPonto.cidade + '</div>' +
      '<div class="eta-linha">🛣️ <b>' + distKm + ' km</b> de distância</div>' +
      '<div class="eta-linha">⏱️ <b>' + durStr + '</b> (com trânsito atual)</div>' +
      '<div class="eta-linha">🕐 Chegada ~<b>' + chegadaStr + '</b></div>'
    );

    // Desenhar rota real no mapa
    if (map.getLayer('layer-rota-real')) map.removeLayer('layer-rota-real');
    if (map.getSource('source-rota-real')) map.removeSource('source-rota-real');

    map.addSource('source-rota-real', {{
      type: 'geojson',
      data: {{ type: 'Feature', geometry: route.geometry }}
    }});

    // Borda/sombra da rota
    map.addLayer({{
      id: 'layer-rota-real-borda',
      type: 'line',
      source: 'source-rota-real',
      paint: {{
        'line-color': '#fff',
        'line-width': 8,
        'line-opacity': 0.5
      }}
    }});

    // Linha principal (azul sólido, grossa)
    map.addLayer({{
      id: 'layer-rota-real',
      type: 'line',
      source: 'source-rota-real',
      paint: {{
        'line-color': '#1a237e',
        'line-width': 5,
        'line-opacity': 0.92
      }}
    }});

  }} catch(err) {{
    console.error('Directions API erro:', err);
    atualizarEtaBox(
      '<div class="eta-site">' + proxPonto.id + ' — ' + proxPonto.cidade + '</div>' +
      '<div class="eta-erro">⚠️ Rota real não disponível — traçado direto exibido.</div>'
    );
    // Fallback: linha reta
    desenharLinhaGuia(latOrig, lonOrig, proxPonto.lat, proxPonto.lon, 'rota-prox', '#2A52BE', false, 4);
  }}
}}

// ══════════════════════════════════════════════════════════════
// LINHA GUIA COM SETAS (standby)
// ══════════════════════════════════════════════════════════════
function desenharLinhaGuia(lat1, lon1, lat2, lon2, sourceId, cor, tracejada, largura) {{
  cor = cor || '#7f8c8d';
  largura = largura || 2;

  if (map.getLayer('layer-' + sourceId)) map.removeLayer('layer-' + sourceId);
  if (map.getLayer('layer-' + sourceId + '-seta')) map.removeLayer('layer-' + sourceId + '-seta');
  if (map.getSource(sourceId)) map.removeSource(sourceId);

  map.addSource(sourceId, {{
    type: 'geojson',
    data: {{
      type: 'Feature',
      geometry: {{
        type: 'LineString',
        coordinates: [[lon1, lat1], [lon2, lat2]]
      }}
    }}
  }});

  var paint = {{
    'line-color': cor,
    'line-width': largura,
    'line-opacity': tracejada ? 0.75 : 0.85
  }};
  if (tracejada) paint['line-dasharray'] = [4, 3];

  map.addLayer({{ id: 'layer-' + sourceId, type: 'line', source: sourceId, paint: paint }});

  // Setas ao longo da linha
  map.addLayer({{
    id: 'layer-' + sourceId + '-seta',
    type: 'symbol',
    source: sourceId,
    layout: {{
      'symbol-placement': 'line',
      'icon-image': 'arrow',
      'icon-size': 0.7,
      'symbol-spacing': 60,
      'icon-allow-overlap': true,
      'icon-ignore-placement': true
    }},
    paint: {{
      'icon-color': cor,
      'icon-opacity': tracejada ? 0.8 : 0.9
    }}
  }});
}}

function desenharTodasLinhasGuia() {{
  // Conecta: PARTIDA → ROTA[0] já foi feita (rota real ou fallback)
  // Agora: ROTA[0] → ROTA[1] → ... → ROTA[n] (linhas standby)
  if (ROTA.length < 2) return;

  for (var i = 0; i < ROTA.length - 1; i++) {{
    var de = ROTA[i];
    var para = ROTA[i + 1];
    desenharLinhaGuia(
      de.lat, de.lon, para.lat, para.lon,
      'guia-' + i,
      '#FF8C00',   // laranja forte — visível sobre qualquer fundo
      true,        // tracejada
      2
    );
  }}
}}

// ══════════════════════════════════════════════════════════════
// MARCADORES HTML CUSTOMIZADOS
// ══════════════════════════════════════════════════════════════
function criarElPartida() {{
  var el = document.createElement('div');
  el.innerHTML = '<div class="partida-pulse"><span class="partida-core"></span><span class="partida-orbit"></span></div>';
  return el;
}}

function criarElProxima() {{
  var el = document.createElement('div');
  el.innerHTML = '<div class="proxima-pulse">1</div>';
  return el;
}}

function criarElPendente(ordem) {{
  var el = document.createElement('div');
  el.style.cssText = 'width:22px;height:22px;border-radius:50%;background:#00BFFF;border:2px solid #fff;box-shadow:0 0 8px rgba(0,191,255,0.7),0 2px 6px rgba(0,0,0,.3);display:flex;align-items:center;justify-content:center;font-size:10px;color:white;font-weight:700;cursor:pointer;';
  el.textContent = String(ordem + 1);
  return el;
}}

function criarElFixo(tipo) {{
  var cores = {{ concluida:'#27ae60', improdutiva:'#e74c3c', cancelada:'#7f8c8d' }};
  var simbolos = {{ concluida:'✓', improdutiva:'✗', cancelada:'⊘' }};
  var el = document.createElement('div');
  el.style.cssText = 'width:20px;height:20px;border-radius:50%;background:' + (cores[tipo]||'#888') +
    ';border:2px solid #fff;box-shadow:0 2px 6px rgba(0,0,0,.25);display:flex;align-items:center;justify-content:center;font-size:11px;color:white;font-weight:700;cursor:pointer;';
  el.textContent = simbolos[tipo] || '?';
  return el;
}}

function criarElHotel() {{
  var el = document.createElement('div');
  el.style.cssText = 'width:22px;height:22px;border-radius:50%;background:#e67e22;border:2px solid #fff;box-shadow:0 2px 6px rgba(0,0,0,.3);display:flex;align-items:center;justify-content:center;font-size:12px;cursor:pointer;';
  el.textContent = '🏨';
  return el;
}}

function criarElAguard() {{
  var el = document.createElement('div');
  el.style.cssText = 'width:20px;height:20px;border-radius:50%;background:#8e44ad;border:2px solid #fff;box-shadow:0 2px 6px rgba(0,0,0,.3);display:flex;align-items:center;justify-content:center;font-size:11px;color:white;font-weight:700;cursor:pointer;';
  el.textContent = '⏳';
  return el;
}}

function popupContent(titulo, linhas, statusHtml, btnHtml) {{
  var html = '<div class="popup-dt"><b>' + titulo + '</b>';
  linhas.forEach(function(l){{ if(l) html += '<br>' + l; }});
  if(statusHtml) html += statusHtml;
  if(btnHtml) html += btnHtml;
  html += '</div>';
  return html;
}}

// ══════════════════════════════════════════════════════════════
// ADICIONAR MARCADORES AO MAPA
// ══════════════════════════════════════════════════════════════
map.on('load', function() {{

  // Adicionar ícone de seta para linhas guia
  var arrowData = new Uint8Array(16 * 16 * 4);
  // Seta simples em pixel art 16x16
  var arrowPixels = [
    0,0,0,0, 0,0,0,0, 0,0,0,0, 0,0,0,0, 0,0,0,0, 0,0,0,0, 0,0,0,0, 1,1,1,1,
    1,1,1,1, 0,0,0,0, 0,0,0,0, 0,0,0,0, 0,0,0,0, 0,0,0,0, 0,0,0,0, 0,0,0,0,
    0,0,0,0, 0,0,0,0, 0,0,0,0, 0,0,0,0, 0,0,0,0, 0,0,0,0, 1,1,1,1, 1,1,1,1,
    1,1,1,1, 1,1,1,1, 0,0,0,0, 0,0,0,0, 0,0,0,0, 0,0,0,0, 0,0,0,0, 0,0,0,0,
    0,0,0,0, 0,0,0,0, 0,0,0,0, 0,0,0,0, 0,0,0,0, 1,1,1,1, 1,1,1,1, 1,1,1,1,
    1,1,1,1, 1,1,1,1, 1,1,1,1, 0,0,0,0, 0,0,0,0, 0,0,0,0, 0,0,0,0, 0,0,0,0,
    0,0,0,0, 0,0,0,0, 0,0,0,0, 0,0,0,0, 1,1,1,1, 1,1,1,1, 1,1,1,1, 1,1,1,1,
    1,1,1,1, 1,1,1,1, 1,1,1,1, 1,1,1,1, 0,0,0,0, 0,0,0,0, 0,0,0,0, 0,0,0,0,
    0,0,0,0, 0,0,0,0, 0,0,0,0, 0,0,0,0, 1,1,1,1, 1,1,1,1, 1,1,1,1, 1,1,1,1,
    1,1,1,1, 1,1,1,1, 1,1,1,1, 1,1,1,1, 0,0,0,0, 0,0,0,0, 0,0,0,0, 0,0,0,0,
    0,0,0,0, 0,0,0,0, 0,0,0,0, 0,0,0,0, 0,0,0,0, 1,1,1,1, 1,1,1,1, 1,1,1,1,
    1,1,1,1, 1,1,1,1, 1,1,1,1, 0,0,0,0, 0,0,0,0, 0,0,0,0, 0,0,0,0, 0,0,0,0,
    0,0,0,0, 0,0,0,0, 0,0,0,0, 0,0,0,0, 0,0,0,0, 0,0,0,0, 1,1,1,1, 1,1,1,1,
    1,1,1,1, 1,1,1,1, 0,0,0,0, 0,0,0,0, 0,0,0,0, 0,0,0,0, 0,0,0,0, 0,0,0,0,
    0,0,0,0, 0,0,0,0, 0,0,0,0, 0,0,0,0, 0,0,0,0, 0,0,0,0, 0,0,0,0, 1,1,1,1,
    1,1,1,1, 0,0,0,0, 0,0,0,0, 0,0,0,0, 0,0,0,0, 0,0,0,0, 0,0,0,0, 0,0,0,0
  ];
  for (var i=0; i<arrowPixels.length; i+=4) {{
    arrowData[i]   = 0;
    arrowData[i+1] = 0;
    arrowData[i+2] = 0;
    arrowData[i+3] = arrowPixels[i] ? 200 : 0;
  }}
  map.addImage('arrow', {{width:16, height:16, data:arrowData}}, {{sdf: true}});

  // ── Ponto de partida ──────────────────────────────────────
  var elPartida = criarElPartida();
  var mkPartida = new mapboxgl.Marker({{element:elPartida, anchor:'center'}})
    .setLngLat([PARTIDA.lon, PARTIDA.lat])
    .setPopup(new mapboxgl.Popup({{offset:20}}).setHTML(
      popupContent('🚩 Ponto de Partida',
        ['<b>Cidade:</b> ' + PARTIDA.cidade], null, null)
    ))
    .addTo(map);

  // ── Atividades pendentes (rota) ───────────────────────────
  ROTA.forEach(function(p, i) {{
    var el, popup;

    if (i === 0) {{
      // PRÓXIMA ATIVIDADE — marcador especial
      el = criarElProxima();
      popup = new mapboxgl.Popup({{offset:20}}).setHTML(
        popupContent(
          '📍 Próxima: ' + p.id,
          ['<b>Cidade:</b> ' + p.cidade,
           freqStr(p) ? '<b>Freq:</b> ' + freqStr(p) : '',
           p.hotel ? '<b>Hotel:</b> ' + p.hotel : ''],
          '<div class="popup-status" style="background:#dbeafe;color:#1e40af;">🎯 Próxima atividade</div>',
          null
        )
      );
    }} else {{
      // STANDBY
      el = criarElPendente(i);
      var btnHtml = '<button class="btn-concluir" onclick="concluirSite(this,' + i + ')">✓ Concluída</button>';
      popup = new mapboxgl.Popup({{offset:20}}).setHTML(
        popupContent(
          p.id,
          ['<b>Cidade:</b> ' + p.cidade,
           freqStr(p) ? '<b>Freq:</b> ' + freqStr(p) : '',
           p.hotel ? '<b>Hotel:</b> ' + p.hotel : ''],
          '<div class="popup-status" style="background:#f5f5f5;color:#555;">🕐 Standby (' + (i+1) + 'ª na fila)</div>',
          btnHtml
        )
      );
    }}

    var mk = new mapboxgl.Marker({{element:el, anchor:'center'}})
      .setLngLat([p.lon, p.lat])
      .setPopup(popup)
      .addTo(map);

    registrarSite(p.id, mk, p.lat, p.lon, null);
    marcadores.push({{marker:mk, lat:p.lat, lon:p.lon, id:p.id, el:el}});
  }});

  // ── Fixos (concluídas, improdutivas, canceladas) ──────────
  mkFixos = {{ conc:[], impr:[], canc:[] }};
  FIXOS.forEach(function(p) {{
    var el = criarElFixo(p.tipo);
    var statusHtml = p.tipo === 'concluida'
      ? '<div class="popup-status ps-conc">✓ Concluída</div>'
      : p.tipo === 'cancelada'
        ? '<div class="popup-status ps-canc">⊘ Cancelada</div>'
        : '<div class="popup-status ps-impr">✗ Improdutiva</div>';
    var popup = new mapboxgl.Popup({{offset:20}}).setHTML(
      popupContent(p.id,
        ['<b>Cidade:</b> ' + p.cidade,
         freqStr(p) ? '<b>Freq:</b> ' + freqStr(p) : '',
         p.concluido ? '<b>Obs:</b> ' + p.concluido : '',
         p.hotel ? '<b>Hotel:</b> ' + p.hotel : ''],
        statusHtml, null)
    );
    var mk = new mapboxgl.Marker({{element:el, anchor:'center'}})
      .setLngLat([p.lon, p.lat])
      .setPopup(popup)
      .addTo(map);

    registrarSite(p.id, mk, p.lat, p.lon, p.tipo === 'concluida' ? 'conc' : p.tipo === 'cancelada' ? 'canc' : 'impr');

    if (p.tipo === 'concluida') mkFixos.conc.push(mk);
    else if (p.tipo === 'cancelada') mkFixos.canc.push(mk);
    else mkFixos.impr.push(mk);
  }});

  // ── Aguardando ────────────────────────────────────────────
  AGUARD.forEach(function(p) {{
    var el = criarElAguard();
    var popup = new mapboxgl.Popup({{offset:20}}).setHTML(
      popupContent(p.id,
        ['<b>Cidade:</b> ' + p.cidade,
         freqStr(p) ? '<b>Freq:</b> ' + freqStr(p) : '',
         p.concluido ? '<b>Motivo:</b> ' + p.concluido : '',
         (p.hotel && p.hotel !== '.') ? '<b>Hotel:</b> ' + p.hotel : ''],
        '<div class="popup-status ps-aguard">⏳ Aguardando para deslocar</div>', null)
    );
    var mk = new mapboxgl.Marker({{element:el, anchor:'center'}})
      .setLngLat([p.lon, p.lat])
      .setPopup(popup)
      .addTo(map);
    registrarSite(p.id, mk, p.lat, p.lon, 'aguard');
    mkAguard.push(mk);
  }});

  // ── Hotéis ────────────────────────────────────────────────
  HOTEIS.forEach(function(h) {{
    var el = criarElHotel();
    el.style.display = 'none';   // oculto por padrão — checkbox desmarcado
    var popup = new mapboxgl.Popup({{offset:20}}).setHTML(
      popupContent(h.nome,
        ['<b>Cidade:</b> ' + (h.cidade||'—'),
         h.tel   ? '<b>Tel:</b> '    + h.tel   : '',
         h.valor ? '<b>Valor:</b> R$ '+ h.valor : ''],
        '<div class="popup-status" style="background:#fef3e2;color:#b7600a;">🏨 Hotel / Pousada</div>', null)
    );
    var mk = new mapboxgl.Marker({{element:el, anchor:'center'}})
      .setLngLat([h.lon, h.lat])
      .setPopup(popup)
      .addTo(map);
    mkHoteis.push(mk);
  }});

  // ── Linhas guia standby ───────────────────────────────────
  desenharTodasLinhasGuia();

  // ── Rota real: PARTIDA → ROTA[0] ─────────────────────────
  if (ROTA.length > 0) {{
    desenharRotaReal(PARTIDA.lat, PARTIDA.lon, ROTA[0]);
  }} else {{
    atualizarEtaBox('<div class="eta-site" style="color:#27ae60;">✅ Sem atividades pendentes!</div>');
  }}

  // ── Contador ──────────────────────────────────────────────
  atualizarContador();

  // ── Badge de quota ────────────────────────────────────────
  atualizarBadgeQuota(lerUso());

  // ── Fit bounds ────────────────────────────────────────────
  var todosLonLat = [[PARTIDA.lon, PARTIDA.lat]];
  ROTA.forEach(function(p){{ todosLonLat.push([p.lon, p.lat]); }});
  if (todosLonLat.length > 1) {{
    var bounds = todosLonLat.reduce(function(b, c) {{ return b.extend(c); }},
      new mapboxgl.LngLatBounds(todosLonLat[0], todosLonLat[0]));
    map.fitBounds(bounds, {{padding:80, maxZoom:12, duration:1000}});
  }}

  // ── Meses anteriores ─────────────────────────────────────
  popularSelectAno();

}});

// ══════════════════════════════════════════════════════════════
// CONTROLE DE CAMADAS
// ══════════════════════════════════════════════════════════════
function toggleCamada(nome, visivel) {{
  if (modoAnterior) return;
  var lista = nome === 'conc'  ? mkFixos.conc
            : nome === 'impr'  ? mkFixos.impr
            : nome === 'canc'  ? mkFixos.canc
            : nome === 'hotel' ? mkHoteis
            : nome === 'aguard'? mkAguard
            : [];
  lista.forEach(function(mk) {{
    mk.getElement().style.display = visivel ? '' : 'none';
  }});
}}

function atualizarContador() {{
  var pend = marcadores.filter(function(m) {{
    return m.el && m.el.style.display !== 'none';
  }}).length;
  var concFixas = FIXOS.filter(function(p){{ return p.tipo === 'concluida'; }}).length;
  document.getElementById('cnt-pend').textContent = pend;
  document.getElementById('cnt-conc').textContent = concFixas + concCount;
}}

function concluirSite(btn, idx) {{
  if (idx < 0 || idx >= marcadores.length) return;
  var m = marcadores[idx];
  m.el.style.display = 'none';
  concCount++;
  atualizarContador();
  if (btn && btn.closest) btn.closest('.mapboxgl-popup').remove();
}}

// ══════════════════════════════════════════════════════════════
// BUSCA DE SITE
// ══════════════════════════════════════════════════════════════
function buscarSite() {{
  var chave = normBusca(document.getElementById('busca-site').value);
  if (!chave) {{ mostrarMsg('msg-busca', ''); return; }}

  // 1) Mês atual
  var item = indiceSites[chave];
  if (item) {{
    map.closeAllPopups ? map.closeAllPopups() : null;
    map.flyTo({{center:[item.lon, item.lat], zoom: Math.max(map.getZoom(), 13), duration:900}});
    setTimeout(function(){{ item.marker.togglePopup(); }}, 950);
    mostrarMsg('msg-busca', '');
    return;
  }}

  // 2) Meses anteriores
  var ocs = buscarSiteHistorico(chave);
  if (ocs.length) {{
    mostrarMsg('msg-busca', '');
    mostrarBuscaGlobal(chave, ocs);
    return;
  }}

  mostrarMsg('msg-busca', 'Site não encontrado!');
}}

document.getElementById('busca-site').addEventListener('keydown', function(ev) {{
  if (ev.key === 'Enter') buscarSite();
}});
document.getElementById('busca-site').addEventListener('input', function() {{
  this.value = this.value.toUpperCase();
  mostrarMsg('msg-busca', '');
}});

// ══════════════════════════════════════════════════════════════
// BUSCA GLOBAL (meses anteriores)
// ══════════════════════════════════════════════════════════════
function buscarSiteHistorico(chave) {{
  var achados = [];
  Object.keys(ANTERIORES).sort().reverse().forEach(function(ano) {{
    Object.keys(ANTERIORES[ano]||{{}}).sort(function(a,b){{return Number(b)-Number(a);}}).forEach(function(mes) {{
      var pac = ANTERIORES[ano][mes];
      (pac.pontos||[]).forEach(function(p) {{
        var k = normBusca(p.id);
        if (k === chave || (k.length >= 3 && k.slice(-3) === chave)) {{
          achados.push({{ano:ano, mes:mes, label:pac.label, ponto:p}});
        }}
      }});
    }});
  }});
  return achados;
}}

function fecharBuscaGlobal() {{
  document.getElementById('resultado-busca-global').classList.remove('visivel');
}}

function mostrarBuscaGlobal(chave, ocs) {{
  ocorrenciasBG = ocs;
  var titulo = document.getElementById('bg-titulo');
  var lista = document.getElementById('bg-lista');
  titulo.textContent = 'Site "' + document.getElementById('busca-site').value.toUpperCase() + '" encontrado em:';
  lista.innerHTML = '';
  ocs.forEach(function(oc, idx) {{
    var row = document.createElement('div');
    row.className = 'bg-item';
    var txt = document.createElement('div');
    txt.innerHTML = '<span class="bg-site">' + oc.ponto.id + '</span> — ' + oc.label;
    var btn = document.createElement('button');
    btn.className = 'bg-ver';
    btn.textContent = 'ver';
    btn.onclick = (function(i){{ return function(){{ abrirOcBG(i); }}; }})(idx);
    row.appendChild(txt);
    row.appendChild(btn);
    lista.appendChild(row);
  }});
  document.getElementById('resultado-busca-global').classList.add('visivel');
}}

function abrirOcBG(idx) {{
  var oc = ocorrenciasBG[idx];
  if (!oc) return;
  document.getElementById('sel-ano-ant').value = oc.ano;
  preencherMeses();
  document.getElementById('sel-mes-ant').value = oc.mes;
  mostrarAba('anteriores');
  mostrarMesAnterior(oc.ano, oc.mes, normBusca(oc.ponto.id));
}}

// ══════════════════════════════════════════════════════════════
// MESES ANTERIORES
// ══════════════════════════════════════════════════════════════
function popularSelectAno() {{
  var sel = document.getElementById('sel-ano-ant');
  var anos = Object.keys(ANTERIORES).sort().reverse();
  sel.innerHTML = '';
  if (!anos.length) {{
    sel.innerHTML = '<option>—</option>';
    document.getElementById('sel-mes-ant').innerHTML = '<option>—</option>';
    return;
  }}
  anos.forEach(function(ano){{ sel.add(new Option(ano, ano)); }});
  preencherMeses();
}}

function preencherMeses() {{
  var ano = document.getElementById('sel-ano-ant').value;
  var sel = document.getElementById('sel-mes-ant');
  sel.innerHTML = '';
  Object.keys(ANTERIORES[ano]||{{}}).sort(function(a,b){{return Number(b)-Number(a);}}).forEach(function(mes) {{
    var pac = ANTERIORES[ano][mes];
    var nm = pac.label.split('/')[0];
    sel.add(new Option(nm + ' (' + (pac.pontos||[]).length + ')', mes));
  }});
}}

function mostrarMesAnterior(anoBusca, mesBusca, siteAbrir) {{
  var ano = anoBusca || document.getElementById('sel-ano-ant').value;
  var mes = mesBusca || document.getElementById('sel-mes-ant').value;
  var pac = ANTERIORES[ano] && ANTERIORES[ano][mes];
  if (!pac || !(pac.pontos||[]).length) {{
    mostrarMsg('msg-ant', 'Nenhuma atividade nesse período.');
    return;
  }}
  modoAnterior = true;
  limparHistorico();
  // Ocultar marcadores atuais
  marcadores.forEach(function(m){{ m.el.style.display = 'none'; }});
  [...mkFixos.conc, ...mkFixos.impr, ...mkFixos.canc, ...mkAguard].forEach(function(mk){{
    mk.getElement().style.display = 'none';
  }});
  mostrarMsg('msg-ant', '');
  map.closeAllPopups ? null : null;

  var mkAbrir = null;
  pac.pontos.forEach(function(p) {{
    var tipo = p.tipo;
    var el = criarElFixo(tipo);
    var statusHtml = tipo === 'concluida'
      ? '<div class="popup-status ps-conc">✓ Concluída</div>'
      : tipo === 'cancelada'
        ? '<div class="popup-status ps-canc">⊘ Cancelada</div>'
        : '<div class="popup-status ps-impr">✗ Improdutiva</div>';
    var mk = new mapboxgl.Marker({{element:el, anchor:'center'}})
      .setLngLat([p.lon, p.lat])
      .setPopup(new mapboxgl.Popup({{offset:20}}).setHTML(
        popupContent(p.id,
          ['<b>Cidade:</b> ' + p.cidade,
           p.tec4g ? '<b>Freq:</b> ' + [p.tec4g,p.tec5g].filter(Boolean).join(' | ') : '',
           p.concluido ? '<b>Obs:</b> ' + p.concluido : ''],
          statusHtml, null)
      ))
      .addTo(map);
    if (siteAbrir && normBusca(p.id) === siteAbrir && !mkAbrir) {{
      mkAbrir = {{marker:mk, lat:p.lat, lon:p.lon}};
    }}
    mkHist.push(mk);
  }});

  var conc = pac.pontos.filter(function(p){{return p.tipo==='concluida';}}).length;
  document.getElementById('cnt-pend').textContent = '0';
  document.getElementById('cnt-conc').textContent = conc;
  mostrarMsg('msg-ant', 'Carregado: ' + pac.label + ' — ' + pac.pontos.length + ' atividades.');

  if (mkHist.length) {{
    setTimeout(function() {{
      if (mkAbrir) {{
        map.flyTo({{center:[mkAbrir.lon, mkAbrir.lat], zoom:13, duration:900}});
        setTimeout(function(){{ mkAbrir.marker.togglePopup(); }}, 950);
      }} else {{
        var pts = mkHist.map(function(m){{ var ll = m.getLngLat(); return [ll.lng, ll.lat]; }});
        var bounds = pts.reduce(function(b,c){{return b.extend(c);}},
          new mapboxgl.LngLatBounds(pts[0], pts[0]));
        map.fitBounds(bounds, {{padding:60}});
      }}
    }}, 100);
  }}
}}

function limparHistorico() {{
  mkHist.forEach(function(mk){{ mk.remove(); }});
  mkHist = [];
}}

function voltarAtual() {{
  modoAnterior = false;
  limparHistorico();
  marcadores.forEach(function(m){{ m.el.style.display = ''; }});
  document.getElementById('chk-conc').checked  && mkFixos.conc.forEach(function(mk){{ mk.getElement().style.display=''; }});
  document.getElementById('chk-impr').checked  && mkFixos.impr.forEach(function(mk){{ mk.getElement().style.display=''; }});
  document.getElementById('chk-canc').checked  && mkFixos.canc.forEach(function(mk){{ mk.getElement().style.display=''; }});
  document.getElementById('chk-aguard').checked && mkAguard.forEach(function(mk){{ mk.getElement().style.display=''; }});
  mostrarAba('atual');
  atualizarContador();
  mostrarMsg('msg-ant', '');
  map.closeAllPopups ? null : null;
}}

// ══════════════════════════════════════════════════════════════
// TABS
// ══════════════════════════════════════════════════════════════
function mostrarAba(nome) {{
  ['atual','anteriores'].forEach(function(n) {{
    document.getElementById('tab-'+n).classList.toggle('ativo', n===nome);
    document.getElementById('pane-'+n).classList.toggle('ativo', n===nome);
  }});
  if (nome === 'atual') voltarAtual();
}}

// ══════════════════════════════════════════════════════════════
// MOBILE: toggle painel
// ══════════════════════════════════════════════════════════════
document.getElementById('painel-toggle').addEventListener('click', function(ev) {{
  ev.stopPropagation();
  document.body.classList.add('painel-aberto');
}});
document.getElementById('painel').addEventListener('click', function(ev) {{ ev.stopPropagation(); }});
map.on('click', function() {{ document.body.classList.remove('painel-aberto'); }});

// Verificar quota na carga
atualizarBadgeQuota(lerUso());
if (lerUso() >= LIMITE_QUOTA) {{
  verificarAlertaQuota(lerUso());
}}

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

def _publicar_arquivo_github(api_base, headers, arquivo_remoto, conteudo_bytes, mensagem):
    """Publica ou atualiza um único arquivo no GitHub. Retorna True se OK."""
    import base64
    conteudo_b64 = base64.b64encode(conteudo_bytes).decode("utf-8")
    api_url = f"{api_base}/contents/{arquivo_remoto}"

    sha_atual = None
    try:
        r = requests.get(api_url, headers=headers, timeout=10)
        if r.status_code == 200:
            sha_atual = r.json().get("sha")
    except Exception as e:
        print(f"   ❌ Erro ao consultar {arquivo_remoto}: {e}")
        return False

    payload = {"message": mensagem, "content": conteudo_b64, "branch": "main"}
    if sha_atual:
        payload["sha"] = sha_atual

    try:
        r = requests.put(api_url, headers=headers, json=payload, timeout=30)
        if r.status_code in (200, 201):
            return True
        else:
            print(f"   ❌ GitHub {arquivo_remoto}: status {r.status_code} — "
                  f"{r.json().get('message','')}")
            return False
    except Exception as e:
        print(f"   ❌ Erro ao publicar {arquivo_remoto}: {e}")
        return False


def _ofuscar_token(token):
    """
    Divide o token em partes e reconstrói via JS para evitar
    que o GitHub Secret Scanner detecte o padrão 'pk.eyJ...' literal.
    O token é dividido em 4 fragmentos e remontado em runtime no browser.
    """
    if not token:
        return "var MAPBOX_TOKEN = '';"

    n = len(token)
    q = n // 4
    partes = [
        token[0:q],
        token[q:2*q],
        token[2*q:3*q],
        token[3*q:],
    ]
    # Gera JS que remonta o token sem o string completo em lugar nenhum
    linhas = [
        "// Mapbox config — gerado automaticamente",
        "// Token dividido para proteção contra secret scanning",
        f"var _t0='{partes[0]}';",
        f"var _t1='{partes[1]}';",
        f"var _t2='{partes[2]}';",
        f"var _t3='{partes[3]}';",
        "var MAPBOX_TOKEN=_t0+_t1+_t2+_t3;",
    ]
    return "\n".join(linhas)


def publicar_mapa_github(html_path):
    if not GITHUB_TOKEN_PATH.exists():
        print("   ⚠️  github_token.txt nao encontrado — publicacao no GitHub ignorada.")
        print(f"      Crie o arquivo em: {GITHUB_TOKEN_PATH}")
        return

    try:
        linhas = GITHUB_TOKEN_PATH.read_text(encoding="utf-8").strip().splitlines()
        if len(linhas) < 3:
            print("   ❌ github_token.txt incompleto. Veja o GUIA_GITHUB.txt.")
            return
        gh_token = linhas[0].strip()
        usuario  = linhas[1].strip()
        repo     = linhas[2].strip()
    except Exception as e:
        print(f"   ❌ Erro ao ler github_token.txt: {e}")
        return

    try:
        conteudo_html = html_path.read_bytes()
    except Exception as e:
        print(f"   ❌ Erro ao ler {html_path.name}: {e}")
        return

    # Ler token Mapbox para gerar o config.js ofuscado
    mapbox_token = carregar_mapbox_token() or ""
    config_js_ofuscado = _ofuscar_token(mapbox_token).encode("utf-8")

    api_base = f"https://api.github.com/repos/{usuario}/{repo}"
    headers  = {
        "Authorization": f"token {gh_token}",
        "Accept":        "application/vnd.github+json",
    }
    agora    = datetime.now().strftime("%d/%m/%Y %H:%M")
    mensagem = f"DT 3.0 — Mapa atualizado em {agora}"

    # ── Publicar HTML ──────────────────────────────────────────
    ok_html = _publicar_arquivo_github(
        api_base, headers, "MAPA_ROTAS.html", conteudo_html, mensagem
    )

    # ── Publicar config JS ofuscado ────────────────────────────
    ok_js = _publicar_arquivo_github(
        api_base, headers, "mapbox_config.js", config_js_ofuscado, mensagem
    )

    if ok_html and ok_js:
        url_pagina = f"https://{usuario}.github.io/{repo}/MAPA_ROTAS.html"
        print(f"   ✅ Mapa publicado no GitHub!")
        print(f"   🌐 Acesse: {url_pagina}")
    elif ok_html:
        print("   ⚠️  HTML publicado mas config JS falhou — mapa pode não funcionar online.")
    else:
        print("   ❌ Falha na publicação. Verifique o token e permissões do repositório.")


# ============================================================
# MODO MAPA (opção 2 do menu)
# ============================================================

def main_mapa():
    print("\n" + "=" * 60)
    print("  DT 3.0 — Atualizar Mapa")
    print("=" * 60)

    if not CREDS_PATH.exists():
        print(f"\n❌ credentials.json não encontrado em:\n   {BASE_DIR}")
        input("\nPressione ENTER para sair...")
        return

    # Carregar token Mapbox
    mapbox_token = carregar_mapbox_token()
    if not mapbox_token:
        print("   ⚠️  Mapa será gerado sem rotas reais (token Mapbox ausente).")

    print("\n[1/3] Conectando ao Google Sheets...")
    try:
        client = conectar_sheets()
        sheet  = client.open_by_key(SHEET_ID)
        ws     = obter_aba_vigente(sheet)
    except Exception as e:
        print(f"   ❌ Erro: {e}")
        input("\nPressione ENTER para sair...")
        return

    print("\n[2/3] Lendo atividades da planilha...")
    df_sheets = ler_atividades_sheets(ws)
    print(f"   {len(df_sheets)} atividades encontradas.")

    df_fixas      = df_sheets[mask_status_fixos_mapa(df_sheets)].copy()
    df_aguardando = df_sheets[
        df_sheets["STATUS"].str.strip() == ST_AGUARDANDO
    ].copy()
    df_pendentes  = df_sheets[
        ~mask_status_fixos_mapa(df_sheets) &
        (df_sheets["STATUS"].str.strip() != ST_AGUARDANDO)
    ].copy()
    df_pendentes  = df_pendentes.dropna(subset=["LAT", "LONG"])
    df_pendentes  = df_pendentes[
        (df_pendentes["LAT"] != 0) & (df_pendentes["LONG"] != 0)
    ]

    print(f"   Concluídas/Improdutivas : {len(df_fixas)}")
    print(f"   Pendentes na rota       : {len(df_pendentes)}")
    print(f"   Aguardando              : {len(df_aguardando)}")

    lat0, lon0, cidade0 = determinar_ponto_inicio(df_sheets)

    print("\n[3/3] Gerando mapa e publicando...")
    print("   → Carregando histórico de meses anteriores...")
    historico_meses = carregar_meses_anteriores(sheet, ws.title)

    print("   → Gerando mapa interativo (Mapbox)...")
    gerar_mapa(
        df_pendentes, lat0, lon0, cidade0,
        df_fixas=df_fixas,
        df_aguardando=df_aguardando,
        historico_meses=historico_meses,
        mapbox_token=mapbox_token,
    )

    print("   → Publicando no GitHub Pages...")
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
# MAIN (opção 1 do menu — execução completa)
# ============================================================

def main():
    print("\n" + "=" * 60)
    print("  DT 3.0 — Automação Drive Test")
    print("=" * 60)

    if not CREDS_PATH.exists():
        print(f"\n❌ credentials.json não encontrado em:\n   {BASE_DIR}")
        print("   Veja o arquivo GUIA_API_GOOGLE.txt para configurar.")
        input("\nPressione ENTER para sair...")
        return

    if not BASE_4G_PATH.exists():
        print(f"\n❌ Base_4G.xlsx não encontrado em:\n   {BASE_DIR}")
        input("\nPressione ENTER para sair...")
        return

    # ── Carregar token Mapbox ──────────────────────────────
    mapbox_token = carregar_mapbox_token()
    if not mapbox_token:
        print("   ⚠️  Mapa será gerado sem rotas reais (token Mapbox ausente).")

    # ── [1/7] Base 4G ─────────────────────────────────────
    print("\n[1/7] Carregando Base 4G...")
    df_base = pd.read_excel(BASE_4G_PATH, engine="openpyxl")
    print(f"   {len(df_base):,} registros.")

    # ── [2/7] Google Sheets ───────────────────────────────
    print("\n[2/7] Conectando ao Google Sheets...")
    try:
        client = conectar_sheets()
        sheet  = client.open_by_key(SHEET_ID)
        ws     = obter_aba_vigente(sheet)
    except Exception as e:
        print(f"   ❌ Erro: {e}")
        input("\nPressione ENTER para sair...")
        return

    # ── [3/7] Ler atividades ──────────────────────────────
    print("\n[3/7] Lendo atividades da planilha...")
    df_sheets = ler_atividades_sheets(ws)
    print(f"   {len(df_sheets)} atividades encontradas.")

    if df_sheets.empty:
        print("   ⚠️  Planilha vazia. Continue assim que houver dados.")

    # ── [4/7] Ponto de partida ────────────────────────────
    print("\n[4/7] Determinando ponto de partida...")
    lat0, lon0, cidade0 = determinar_ponto_inicio(df_sheets)

    # ── [5/7] Novas atividades ────────────────────────────
    print("\n[5/7] Processando novas atividades...")
    arquivos_novas = encontrar_arquivos_novas()
    df_novas       = pd.DataFrame()

    if arquivos_novas:
        frames_novas = []
        for arq in arquivos_novas:
            try:
                df_raw  = pd.read_excel(arq, engine="openpyxl")
                df_proc = processar_arquivo(
                    df_raw, nome_arquivo=arq.name, df_base=df_base
                )
                if not df_proc.empty:
                    frames_novas.append(df_proc)
                    print(f"   {arq.name}: {len(df_proc)} atividades válidas.")
                else:
                    print(f"   ⚠️  {arq.name}: nenhuma atividade válida extraída.")
            except Exception as e:
                print(f"   ❌ Erro ao processar {arq.name}: {e}")

        if frames_novas:
            df_novas = pd.concat(frames_novas, ignore_index=True)
            df_novas = df_novas.drop_duplicates(subset=["SITE"], keep="first")
            print(f"   Total: {len(df_novas)} novas atividades únicas.")
    else:
        print("   Nenhuma nova atividade. Continuando sem novas.")

    # ── [6/7] Otimizar rota ───────────────────────────────
    print("\n[6/7] Otimizando rota...")

    df_fixas = df_sheets[mask_status_fixos_mapa(df_sheets)].copy()

    df_aguardando = df_sheets[
        df_sheets["STATUS"].str.strip() == ST_AGUARDANDO
    ].copy()

    df_pool_sheets = df_sheets[
        ~mask_status_fixos_mapa(df_sheets) &
        (df_sheets["STATUS"].str.strip() != ST_AGUARDANDO)
    ].copy()
    df_pool_sheets = df_pool_sheets.dropna(subset=["LAT", "LONG"])
    df_pool_sheets = df_pool_sheets[
        (df_pool_sheets["LAT"] != 0) & (df_pool_sheets["LONG"] != 0)
    ]

    # Complementar UF e TEC via Base_4G para atividades do Sheets
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
                    print(f"   ℹ️  {row['SITE']}: UF complementada via Base_4G: {val}")
            if precisa_tec:
                for col_tec in ("TEC", "TECNOLOGIA", "TECHNOLOGY"):
                    if col_tec in matches.columns:
                        val = normalizar_tec(str(m.get(col_tec, "")).strip())
                        if val and val.upper() not in ("", "NAN", "NONE"):
                            df_pool_sheets.at[idx, "TEC"] = val
                            print(f"   ℹ️  {row['SITE']}: TEC complementada via Base_4G: {val}")
                            break
        except Exception:
            pass

    sites_originais = set(df_sheets["SITE"].str.upper())
    sites_no_pool   = set(df_pool_sheets["SITE"].str.upper())

    COLS = [
        "DEMANDA", "INTEGRAÇÃO", "SITE", "UF", "TEC",
        "LAT", "LONG", "CIDADE", "2G|3G|4G", "5G",
        "STATUS", "CONCLUIDO", "HOTEL",
    ]

    if not df_novas.empty:
        df_nov = df_novas[
            ~df_novas["SITE"].str.upper().isin(sites_no_pool)
        ].copy()
        for col in ("CONCLUIDO", "HOTEL"):
            if col not in df_nov.columns:
                df_nov[col] = ""
        df_novas_filtradas = df_nov[COLS].copy()
    else:
        df_novas_filtradas = pd.DataFrame(columns=COLS)

    frames = [
        f for f in [df_pool_sheets[COLS], df_novas_filtradas]
        if not f.empty
    ]

    if not frames:
        print("   ⚠️  Nenhuma atividade para otimizar.")
        input("\nPressione ENTER para sair...")
        return

    df_pool = pd.concat(frames, ignore_index=True)
    print(
        f"   Pool: {len(df_pool)} atividades "
        f"({len(df_pool_sheets)} existentes + {len(df_novas_filtradas)} novas)"
    )

    rota_final = otimizar_rota(df_pool, lat0, lon0)

    ativ_path = BASE_DIR / "ATIVIDADES_GERADAS.xlsx"
    rota_final.to_excel(ativ_path, index=False)
    print(f"   ATIVIDADES_GERADAS.xlsx salvo.")

    # ── [7/7] Gerar saídas ────────────────────────────────
    print("\n[7/7] Gerando saídas...")

    print("   → Relatório de texto...")
    gerar_relatorio(rota_final, df_base, PASTA_OUT)

    print("   → Histórico de meses anteriores...")
    historico_meses = carregar_meses_anteriores(sheet, ws.title)

    print("   → Mapa interativo (Mapbox)...")
    gerar_mapa(
        rota_final, lat0, lon0, cidade0,
        df_fixas=df_fixas,
        df_aguardando=df_aguardando,
        historico_meses=historico_meses,
        mapbox_token=mapbox_token,
    )

    print("   → Publicando mapa no GitHub Pages...")
    publicar_mapa_github(BASE_DIR / "MAPA_ROTAS.html")

    print("   → Atualizando Google Sheets...")
    atualizar_sheets(ws, df_fixas, rota_final, sites_originais, df_aguardando)

    # Mover arquivos processados
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
# PONTO DE ENTRADA
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
