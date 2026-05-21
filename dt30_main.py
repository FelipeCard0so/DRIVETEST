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

    def d(a, b):
        return haversine(a[0], a[1], b[0], b[1])

    origem = (lat0, lon0)

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

    def _get(campo, default=""):
        if campo not in col:
            return default
        v = df_raw[col[campo]]
        return v

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

            nada_precisa = not any([precisa_cidade, precisa_uf,
                                    precisa_tec, precisa_4g, precisa_5g])
            if nada_precisa or row["LAT"] == 0.0:
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
                                print(f"{prefixo} ℹ️  {row['SITE']}: TEC recuperada da Base_4G: {val}")
                                break

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
                    print(f"   ℹ️  UF do site {site} detectada na Base_4G: {uf}")
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
        print("      Verifique se a service account tem acesso à planilha.")
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


def gerar_mapa(df_rota, lat0, lon0, cidade0, df_fixas=None, df_aguardando=None, historico_meses=None):
    """
    Gera MAPA_ROTAS.html com HTML/JS puro (sem folium).
    Inclui busca global de sites em todos os meses anteriores.
    """
    import json

    hoteis = carregar_hoteis()

    pontos_rota = []
    for _, row in df_rota.iterrows():
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

    # ATENÇÃO: dentro desta f-string todo { } que for JavaScript
    # precisa ser {{ }} para o Python não confundir com interpolação.
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
    font-family:'Segoe UI',sans-serif;min-width:230px;max-width:300px;
  }}
  #painel h4{{margin:0 0 10px;font-size:14px;font-weight:700;color:#1a237e;letter-spacing:.5px;}}
  .tabs-mapa{{display:grid;grid-template-columns:1fr 1fr;gap:6px;margin-bottom:10px;}}
  .tab-mapa{{
    border:1px solid #cfd6e6;background:#f7f9ff;color:#1a237e;border-radius:6px;
    padding:7px 8px;font-size:12px;font-weight:700;cursor:pointer;
  }}
  .tab-mapa.ativo{{background:#1a237e;color:#fff;border-color:#1a237e;}}
  .tab-panel{{display:none;}}
  .tab-panel.ativo{{display:block;}}
  .leg{{display:flex;align-items:center;gap:8px;margin-bottom:6px;font-size:13px;color:#333;}}
  .dot{{width:13px;height:13px;border-radius:50%;display:inline-block;flex-shrink:0;}}
  .dot-pend{{background:#2A52BE;}}
  .dot-conc{{background:#27ae60;}}
  .dot-impr{{background:#e74c3c;}}
  .dot-canc{{background:#7f8c8d;}}
  .dot-part{{background:#f39c12;border:2px solid #b7770d;}}
  .dot-hotel{{background:#e67e22;}}
  .dot-aguard{{background:#8e44ad;border:2px solid #6c3483;}}
  #contador{{margin-top:10px;padding-top:10px;border-top:1px solid #ddd;font-size:12px;color:#555;}}
  #contador span{{font-weight:700;color:#1a237e;}}
  .sep{{border:none;border-top:1px solid #e0e0e0;margin:10px 0;}}
  .filtro{{display:flex;align-items:center;gap:8px;margin-bottom:5px;font-size:12px;color:#444;cursor:pointer;user-select:none;}}
  .filtro input[type=checkbox]{{width:14px;height:14px;cursor:pointer;accent-color:#1a237e;flex-shrink:0;}}
  .filtro span.dot{{flex-shrink:0;}}
  #filtros-label,.busca-site-label,.anteriores-label{{
    font-size:11px;font-weight:700;color:#888;letter-spacing:.6px;text-transform:uppercase;margin-bottom:6px;
  }}
  .busca-site-wrap{{margin-top:10px;padding-top:10px;border-top:1px solid #ddd;}}
  .busca-site-label{{display:block;}}
  .busca-site-row{{display:flex;gap:6px;align-items:center;}}
  #busca-site{{
    flex:1;min-width:0;border:1px solid #cfd6e6;border-radius:6px;
    padding:8px 9px;font-size:13px;outline:none;
    font-family:'Segoe UI',sans-serif;text-transform:uppercase;
  }}
  #busca-site:focus,.select-ant:focus{{border-color:#2A52BE;box-shadow:0 0 0 2px rgba(42,82,190,.14);}}
  #btn-busca-site,.btn-ant{{
    border:none;border-radius:6px;background:#1a237e;color:white;font-weight:700;cursor:pointer;
  }}
  #btn-busca-site{{width:34px;height:34px;line-height:1;}}
  #btn-busca-site:hover,.btn-ant:hover{{background:#121858;}}
  #msg-busca-site,#msg-ant{{
    min-height:16px;margin-top:5px;font-size:11px;color:#b35b00;
    opacity:0;transition:opacity .18s ease;
  }}
  #msg-busca-site.visivel,#msg-ant.visivel{{opacity:1;}}
  #resultado-busca-global{{
    display:none;position:absolute;top:88px;right:330px;z-index:1002;
    width:min(360px,calc(100vw - 24px));max-height:64vh;overflow:auto;
    background:rgba(255,255,255,.98);border-radius:10px;
    box-shadow:0 6px 24px rgba(0,0,0,.22);font-family:'Segoe UI',sans-serif;
    padding:12px 14px;
  }}
  #resultado-busca-global.visivel{{display:block;}}
  .busca-global-topo{{display:flex;align-items:center;justify-content:space-between;gap:10px;margin-bottom:8px;}}
  .busca-global-titulo{{font-size:13px;font-weight:700;color:#1a237e;line-height:1.25;}}
  #fechar-busca-global{{
    width:26px;height:26px;border:none;border-radius:6px;background:#eef1f7;
    color:#1a237e;font-weight:800;cursor:pointer;line-height:1;
  }}
  #fechar-busca-global:hover{{background:#dfe5f1;}}
  .busca-global-item{{
    display:flex;align-items:center;justify-content:space-between;gap:10px;
    border-top:1px solid #e6e8ef;padding:8px 0;font-size:12px;color:#333;
  }}
  .busca-global-item:first-child{{border-top:none;}}
  .busca-global-site{{font-weight:700;color:#1a237e;}}
  .busca-global-ver{{
    border:none;border-radius:6px;background:#1a237e;color:white;
    font-size:11px;font-weight:700;padding:6px 9px;cursor:pointer;white-space:nowrap;
  }}
  .busca-global-ver:hover{{background:#121858;}}
  @media (max-width: 768px){{
    #resultado-busca-global{{right:52px;top:92px;width:calc(100vw - 68px);max-height:56vh;}}
  }}
  .anteriores-grid{{display:grid;grid-template-columns:1fr 1fr;gap:7px;margin-top:8px;}}
  .select-ant{{width:100%;border:1px solid #cfd6e6;border-radius:6px;padding:8px 7px;font-size:12px;background:#fff;}}
  .btn-ant{{width:100%;padding:9px;margin-top:8px;font-size:12px;}}
  .partida-pulse{{
    width:30px;height:30px;border-radius:50%;position:relative;
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
    position:absolute;left:50%;top:50%;width:9px;height:9px;border-radius:50%;
    background:#fff;transform:translate(-50%,-50%);
    box-shadow:0 0 0 2px rgba(183,119,13,.55);
  }}
  .partida-orbit{{
    position:absolute;left:50%;top:50%;width:40px;height:40px;margin:-20px 0 0 -20px;
    border-radius:50%;border:2px dashed rgba(26,35,126,.55);
    animation:dtSpin 3.8s linear infinite;pointer-events:none;
  }}
  @keyframes dtPulse{{0%{{transform:scale(.55);opacity:.85;}}70%{{transform:scale(1.55);opacity:0;}}100%{{transform:scale(1.55);opacity:0;}}}}
  @keyframes dtSpin{{to{{transform:rotate(360deg);}}}}
  .tip-site{{font-weight:bold;font-size:13px;color:#1a237e;}}
  .tip-hotel{{font-weight:bold;font-size:14px;color:#d35400;background:white;border:2px solid #e67e22;padding:4px 7px;border-radius:4px;}}
  .tip-conc{{font-size:13px;color:#27ae60;font-weight:600;}}
  .tip-impr{{font-size:13px;color:#e74c3c;font-weight:600;}}
  .tip-canc{{font-size:13px;color:#7f8c8d;font-weight:600;}}
  .popup-box{{font-family:'Segoe UI',sans-serif;font-size:13px;min-width:200px;}}
  .popup-box b{{color:#1a237e;}}
  .popup-status{{margin-top:6px;padding:4px 8px;border-radius:4px;font-weight:600;font-size:12px;}}
  .ps-conc{{background:#d4edda;color:#155724;}}
  .ps-impr{{background:#f8d7da;color:#721c24;}}
  .ps-canc{{background:#eceff1;color:#455a64;}}
  .btn-concluir{{
    margin-top:10px;background:#27ae60;color:white;border:none;
    padding:10px;border-radius:5px;cursor:pointer;width:100%;
    font-weight:700;font-size:15px;letter-spacing:.3px;
  }}
  .btn-concluir:hover{{background:#1e8449;}}
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
  }}
</style>
</head>
<body>
<div id="map"></div>
<button id="painel-toggle" type="button" aria-label="Abrir painel"><span></span><span></span><span></span></button>
<div id="painel">
  <h4>DT 3.0 — Rota</h4>
  <div class="tabs-mapa">
    <button class="tab-mapa ativo" id="tab-atual" type="button">Atual</button>
    <button class="tab-mapa" id="tab-anteriores" type="button">Anteriores</button>
  </div>

  <div class="tab-panel ativo" id="pane-atual">
    <div class="leg"><span class="dot dot-part"></span> Ponto de partida</div>
    <div class="leg"><span class="dot dot-pend"></span> Pendente (rota ativa)</div>
    <div id="contador">
      Pendentes: <span id="cnt-pend">0</span> |
      Concluidas: <span id="cnt-conc">0</span>
    </div>
    <div class="busca-site-wrap">
      <label class="busca-site-label" for="busca-site">Buscar site</label>
      <div class="busca-site-row">
        <input id="busca-site" type="text" placeholder="Ex: MG-ABC" autocomplete="off"/>
        <button id="btn-busca-site" type="button" title="Encontrar site">&#8981;</button>
      </div>
      <div id="msg-busca-site" aria-live="polite"></div>
    </div>
    <hr class="sep"/>
    <div id="filtros-label">Exibir camadas</div>
    <label class="filtro"><input type="checkbox" id="chk-conc" checked/><span class="dot dot-conc"></span> Concluidas</label>
    <label class="filtro"><input type="checkbox" id="chk-impr" checked/><span class="dot dot-impr"></span> Improdutivas</label>
    <label class="filtro"><input type="checkbox" id="chk-canc" checked/><span class="dot dot-canc"></span> Canceladas</label>
    <label class="filtro"><input type="checkbox" id="chk-hotel" checked/><span class="dot dot-hotel"></span> Hoteis</label>
    <label class="filtro"><input type="checkbox" id="chk-aguard" checked/><span class="dot dot-aguard"></span> Aguardando</label>
  </div>

  <div class="tab-panel" id="pane-anteriores">
    <div class="anteriores-label">Consultar periodo</div>
    <div class="anteriores-grid">
      <select class="select-ant" id="sel-ano-ant"></select>
      <select class="select-ant" id="sel-mes-ant"></select>
    </div>
    <button class="btn-ant" id="btn-ant-ok" type="button">OK</button>
    <button class="btn-ant" id="btn-voltar-atual" type="button" style="background:#6c757d;">Voltar ao atual</button>
    <div id="msg-ant" aria-live="polite"></div>
  </div>
</div>

<!-- Painel de resultado da busca global — fica fora do #painel para não ser afetado pelo overflow -->
<div id="resultado-busca-global" aria-live="polite">
  <div class="busca-global-topo">
    <div class="busca-global-titulo" id="busca-global-titulo"></div>
    <button id="fechar-busca-global" type="button" title="Fechar">X</button>
  </div>
  <div id="busca-global-lista"></div>
</div>

<script>
var map = L.map("map").setView([{lat0},{lon0}], 7);
L.tileLayer("https://tile.openstreetmap.org/{{z}}/{{x}}/{{y}}.png",
  {{attribution:'&copy; <a href="https://www.openstreetmap.org/copyright">OpenStreetMap</a> contributors',maxZoom:19}}).addTo(map);

var PARTIDA    = {j_part};
var ROTA       = {j_rota};
var FIXOS      = {j_fixos};
var AGUARD     = {j_aguard};
var HOTEIS     = {j_hoteis};
var ANTERIORES = {j_hist};

var concCount  = 0;
var indiceSites = {{}};
var buscaTimer  = null;
var modoAnterior = false;
var mkHist = [];
var ocorrenciasBuscaGlobal = [];

/* ── Utilitários ───────────────────────────────────────────── */

function normalizarBuscaSite(valor) {{
  return String(valor || '').trim().toUpperCase().replace(/\s+/g, '').replace(/[^A-Z0-9]/g, '');
}}

function registrarSiteBusca(site, marker, lat, lon, listaCamada) {{
  var chave = normalizarBuscaSite(site);
  if (!chave) return;
  var item = {{site: site, marker: marker, lat: lat, lon: lon, listaCamada: listaCamada || null}};
  indiceSites[chave] = item;
  if (chave.length >= 3) indiceSites[chave.slice(-3)] = item;
}}

function mostrarMensagem(elId, texto) {{
  var msg = document.getElementById(elId);
  msg.textContent = texto || '';
  msg.classList.toggle('visivel', Boolean(texto));
  if (elId === 'msg-busca-site') {{
    if (buscaTimer) clearTimeout(buscaTimer);
    if (texto) buscaTimer = setTimeout(function() {{ msg.classList.remove('visivel'); }}, 2600);
  }}
}}

function mostrarMensagemBusca(texto) {{ mostrarMensagem('msg-busca-site', texto); }}

function garantirCamadaVisivel(item) {{
  if (item.listaCamada === 'conc')  document.getElementById('chk-conc').checked  = true;
  if (item.listaCamada === 'impr')  document.getElementById('chk-impr').checked  = true;
  if (item.listaCamada === 'canc')  document.getElementById('chk-canc').checked  = true;
  if (item.listaCamada === 'aguard') document.getElementById('chk-aguard').checked = true;
  if (item.listaCamada && !map.hasLayer(item.marker)) item.marker.addTo(map);
}}

/* ── Busca global em meses anteriores ─────────────────────── */

function procurarSiteHistorico(chave) {{
  var achados = [];
  Object.keys(ANTERIORES).sort().reverse().forEach(function(ano) {{
    Object.keys(ANTERIORES[ano] || {{}})
      .sort(function(a, b) {{ return Number(b) - Number(a); }})
      .forEach(function(mes) {{
        var pacote = ANTERIORES[ano][mes];
        (pacote.pontos || []).forEach(function(p) {{
          var siteChave = normalizarBuscaSite(p.id);
          if (siteChave === chave || (siteChave.length >= 3 && siteChave.slice(-3) === chave)) {{
            achados.push({{ano: ano, mes: mes, label: pacote.label, ponto: p}});
          }}
        }});
      }});
  }});
  return achados;
}}

function fecharBuscaGlobal() {{
  document.getElementById('resultado-busca-global').classList.remove('visivel');
}}

function mostrarBuscaGlobal(chave, ocorrencias) {{
  ocorrenciasBuscaGlobal = ocorrencias;
  var box    = document.getElementById('resultado-busca-global');
  var titulo = document.getElementById('busca-global-titulo');
  var lista  = document.getElementById('busca-global-lista');

  titulo.textContent = 'Site "' + document.getElementById('busca-site').value.toUpperCase() + '" encontrado em:';
  lista.innerHTML = '';

  ocorrencias.forEach(function(oc, idx) {{
    var row   = document.createElement('div');
    row.className = 'busca-global-item';

    var texto = document.createElement('div');
    texto.innerHTML = '<span class="busca-global-site">Site "' + oc.ponto.id + '"</span>&nbsp;' + oc.label;

    var btn = document.createElement('button');
    btn.type = 'button';
    btn.className = 'busca-global-ver';
    btn.textContent = 'ver';
    btn.onclick = (function(i) {{
      return function() {{ abrirOcorrenciaBuscaGlobal(i); }};
    }})(idx);

    row.appendChild(texto);
    row.appendChild(btn);
    lista.appendChild(row);
  }});

  box.classList.add('visivel');
}}

function abrirOcorrenciaBuscaGlobal(idx) {{
  var oc = ocorrenciasBuscaGlobal[idx];
  if (!oc) return;

  /* Selecionar ano */
  var selAno = document.getElementById('sel-ano-ant');
  selAno.value = oc.ano;
  var ev = document.createEvent('HTMLEvents');
  ev.initEvent('change', true, false);
  selAno.dispatchEvent(ev);

  /* Selecionar mês */
  document.getElementById('sel-mes-ant').value = oc.mes;

  /* Ir para aba Anteriores e carregar o mês */
  mostrarAba('anteriores');
  mostrarMesAnterior(oc.ano, oc.mes, normalizarBuscaSite(oc.ponto.id));
}}

/* ── Busca principal ───────────────────────────────────────── */

function buscarSiteNoMapa() {{
  var chave = normalizarBuscaSite(document.getElementById('busca-site').value);
  if (!chave) {{ mostrarMensagemBusca(''); return; }}

  /* 1) Site no mês atual */
  var item = indiceSites[chave];
  var itemPodeAbrir = item && (item.listaCamada || map.hasLayer(item.marker));
  if (itemPodeAbrir) {{
    mostrarMensagemBusca('');
    garantirCamadaVisivel(item);
    map.closePopup();
    map.flyTo([item.lat, item.lon], Math.max(map.getZoom(), 13), {{animate: true, duration: .85}});
    setTimeout(function() {{ item.marker.openPopup(); }}, 900);
    return;
  }}

  /* 2) Site em meses anteriores */
  var ocorrencias = procurarSiteHistorico(chave);
  if (ocorrencias.length) {{
    mostrarMensagemBusca('');
    mostrarBuscaGlobal(chave, ocorrencias);
    return;
  }}

  mostrarMensagemBusca('Site não encontrado!');
}}

/* ── Ícones ────────────────────────────────────────────────── */

function mkIcon(color, icon) {{
  return L.AwesomeMarkers.icon({{icon: icon, markerColor: color, prefix: 'glyphicon'}});
}}

var icPartida = L.divIcon({{
  className: 'partida-animada',
  html: '<div class="partida-pulse"><span class="partida-core"></span><span class="partida-orbit"></span></div>',
  iconSize: [30, 30], iconAnchor: [15, 15], popupAnchor: [0, -18]
}});

var icPend   = mkIcon('blue',      'map-marker');
var icConc   = mkIcon('green',     'ok-sign');
var icImpr   = mkIcon('red',       'remove-sign');
var icCanc   = mkIcon('cadetblue', 'ban-circle');
var icHotel  = mkIcon('orange',    'tower');
var icAguard = mkIcon('purple',    'time');

/* ── Hotéis ────────────────────────────────────────────────── */

var mkHoteis = [];
HOTEIS.forEach(function(h) {{
  var mk = L.marker([h.lat, h.lon], {{icon: icHotel}}).addTo(map);
  mk.bindTooltip(h.nome, {{sticky: true, className: 'tip-hotel', direction: 'top'}});
  mk.bindPopup(
    "<div class='popup-box'><b>" + h.nome + "</b><br>" +
    (h.cidade ? "<b>Cidade:</b> " + h.cidade + "<br>" : "") +
    (h.tel    ? "<b>Telefone:</b> " + h.tel  + "<br>" : "") +
    (h.valor  ? "<b>Ultimo valor:</b> R$ " + h.valor + "<br>" : "") +
    "<div class='popup-status' style='background:#fef3e2;color:#b7600a;'>Hotel / Pousada</div></div>"
  );
  mkHoteis.push(mk);
}});

/* ── Ponto de partida ──────────────────────────────────────── */

var mkPartida = L.marker([PARTIDA.lat, PARTIDA.lon], {{icon: icPartida}})
  .addTo(map)
  .bindTooltip("Partida: " + PARTIDA.cidade, {{sticky: true, className: 'tip-site'}})
  .bindPopup("<div class='popup-box'><b>Ponto de partida</b><br>" + PARTIDA.cidade + "</div>");

/* ── Estado global ─────────────────────────────────────────── */

var marcadores = [];
var polyline;
var mkConc   = [];
var mkImpr   = [];
var mkCanc   = [];
var mkAguard = [];

function coordsRota() {{
  var pts = [[PARTIDA.lat, PARTIDA.lon]];
  marcadores.forEach(function(m) {{ if (map.hasLayer(m.marker)) pts.push([m.lat, m.lon]); }});
  return pts;
}}

function redesenharRota() {{
  if (polyline) map.removeLayer(polyline);
  var pts = coordsRota();
  if (pts.length > 1)
    polyline = L.polyline(pts, {{color: '#2A52BE', weight: 4, opacity: 0.65}}).addTo(map);
  var pend      = marcadores.filter(function(m) {{ return map.hasLayer(m.marker); }}).length;
  var concFixas = FIXOS.filter(function(p) {{ return p.tipo === 'concluida'; }}).length;
  document.getElementById('cnt-pend').textContent = pend;
  document.getElementById('cnt-conc').textContent = concFixas + concCount;
}}

function popupFixo(p, statusHtml) {{
  var freqStr = [p.tec4g, p.tec5g].filter(Boolean).join(' | ');
  return "<div class='popup-box'>" +
    "<b>" + p.id + "</b><br>" +
    "<b>Cidade:</b> " + p.cidade + "<br>" +
    (freqStr     ? "<b>Freq:</b> "  + freqStr     + "<br>" : "") +
    (p.concluido ? "<b>Obs:</b> "   + p.concluido + "<br>" : "") +
    (p.hotel     ? "<b>Hotel:</b> " + p.hotel     + "<br>" : "") +
    statusHtml + "</div>";
}}

function dadosTipo(p) {{
  if (p.tipo === 'concluida')
    return {{icon: icConc, cls: 'tip-conc', label: '✓ ', layer: 'conc',
             status: "<div class='popup-status ps-conc'>✓ Atividade concluída</div>"}};
  if (p.tipo === 'cancelada')
    return {{icon: icCanc, cls: 'tip-canc', label: '⊘ ', layer: 'canc',
             status: "<div class='popup-status ps-canc'>Cancelada</div>"}};
  return {{icon: icImpr, cls: 'tip-impr', label: '✗ ', layer: 'impr',
           status: "<div class='popup-status ps-impr'>✗ Improdutiva</div>"}};
}}

/* ── Rota ativa (pendentes) ────────────────────────────────── */

ROTA.forEach(function(p) {{
  var marker = L.marker([p.lat, p.lon], {{icon: icPend}}).addTo(map);
  marker.bindTooltip("SITE: " + p.id + "<br>" + p.cidade, {{sticky: true, className: 'tip-site'}});

  var pop = document.createElement('div');
  pop.className = 'popup-box';
  var freqStr = [p.tec4g, p.tec5g].filter(Boolean).join(' | ');
  pop.innerHTML =
    "<b>" + p.id + "</b><br><b>Cidade:</b> " + p.cidade + "<br>" +
    (freqStr ? "<b>Freq:</b> " + freqStr + "<br>" : "") +
    (p.hotel ? "<b>Hotel:</b> " + p.hotel + "<br>" : "");

  var btn = document.createElement('button');
  btn.className = 'btn-concluir';
  btn.innerHTML = '&#10003; Marcar como Conclu&#237;da';
  btn.onclick = function() {{
    map.removeLayer(marker);
    concCount += 1;
    var mk2 = L.marker([p.lat, p.lon], {{icon: icConc}}).addTo(map);
    mk2.bindTooltip("✓ " + p.id, {{sticky: true, className: 'tip-conc'}});
    mk2.bindPopup("<div class='popup-box'><b>" + p.id + "</b><br>" + p.cidade +
                  "<div class='popup-status ps-conc'>✓ Concluída</div></div>");
    registrarSiteBusca(p.id, mk2, p.lat, p.lon, null);
    redesenharRota();
  }};
  pop.appendChild(btn);
  marker.bindPopup(pop);
  marcadores.push({{marker: marker, lat: p.lat, lon: p.lon}});
  registrarSiteBusca(p.id, marker, p.lat, p.lon, null);
}});

/* ── Fixos (concluídas / improdutivas / canceladas) ────────── */

FIXOS.forEach(function(p) {{
  var d  = dadosTipo(p);
  var mk = L.marker([p.lat, p.lon], {{icon: d.icon}}).addTo(map);
  mk.bindTooltip(d.label + p.id, {{sticky: true, className: d.cls}});
  mk.bindPopup(popupFixo(p, d.status));
  registrarSiteBusca(p.id, mk, p.lat, p.lon, d.layer);
  if (p.tipo === 'concluida')  mkConc.push(mk);
  else if (p.tipo === 'cancelada') mkCanc.push(mk);
  else mkImpr.push(mk);
}});

/* ── Aguardando ────────────────────────────────────────────── */

AGUARD.forEach(function(p) {{
  var mk = L.marker([p.lat, p.lon], {{icon: icAguard}}).addTo(map);
  mk.bindTooltip("Aguardando: " + p.id + "<br>" + p.cidade, {{sticky: true, className: 'tip-site'}});
  var freqStr = [p.tec4g, p.tec5g].filter(Boolean).join(' | ');
  mk.bindPopup(
    "<div class='popup-box'><b>" + p.id + "</b><br><b>Cidade:</b> " + p.cidade + "<br>" +
    (freqStr        ? "<b>Freq:</b> "   + freqStr        + "<br>" : "") +
    (p.concluido    ? "<b>Motivo:</b> " + p.concluido    + "<br>" : "") +
    (p.hotel && p.hotel !== '.' ? "<b>Hotel:</b> " + p.hotel + "<br>" : "") +
    "<div class='popup-status' style='background:#e8daef;color:#6c3483;'>Aguardando para deslocar</div></div>"
  );
  registrarSiteBusca(p.id, mk, p.lat, p.lon, 'aguard');
  mkAguard.push(mk);
}});

/* ── Controle de camadas ───────────────────────────────────── */

function toggleCamada(lista, visivel) {{
  lista.forEach(function(mk) {{
    if (visivel) {{ if (!map.hasLayer(mk)) mk.addTo(map); }}
    else         {{ if (map.hasLayer(mk))  map.removeLayer(mk); }}
  }});
}}

function setAtualVisivel(visivel) {{
  if (visivel) {{
    marcadores.forEach(function(m) {{ if (!map.hasLayer(m.marker)) m.marker.addTo(map); }});
    toggleCamada(mkConc,   document.getElementById('chk-conc').checked);
    toggleCamada(mkImpr,   document.getElementById('chk-impr').checked);
    toggleCamada(mkCanc,   document.getElementById('chk-canc').checked);
    toggleCamada(mkAguard, document.getElementById('chk-aguard').checked);
    if (!map.hasLayer(mkPartida)) mkPartida.addTo(map);
    redesenharRota();
  }} else {{
    marcadores.forEach(function(m) {{ if (map.hasLayer(m.marker)) map.removeLayer(m.marker); }});
    toggleCamada(mkConc, false); toggleCamada(mkImpr, false);
    toggleCamada(mkCanc, false); toggleCamada(mkAguard, false);
    if (polyline) map.removeLayer(polyline);
  }}
}}

function limparHistorico() {{
  mkHist.forEach(function(mk) {{ if (map.hasLayer(mk)) map.removeLayer(mk); }});
  mkHist = [];
}}

/* ── Tabs ──────────────────────────────────────────────────── */

function mostrarAba(nome) {{
  document.getElementById('tab-atual').classList.toggle('ativo',      nome === 'atual');
  document.getElementById('tab-anteriores').classList.toggle('ativo', nome === 'anteriores');
  document.getElementById('pane-atual').classList.toggle('ativo',     nome === 'atual');
  document.getElementById('pane-anteriores').classList.toggle('ativo',nome === 'anteriores');
}}

/* ── Meses anteriores ──────────────────────────────────────── */

function popularMesesAnteriores() {{
  var selAno = document.getElementById('sel-ano-ant');
  var selMes = document.getElementById('sel-mes-ant');
  var anos   = Object.keys(ANTERIORES).sort().reverse();

  selAno.innerHTML = '';
  if (!anos.length) {{
    selAno.innerHTML = '<option value="">Ano</option>';
    selMes.innerHTML = '<option value="">Mes</option>';
    mostrarMensagem('msg-ant', 'Nenhum mês anterior disponível.');
    return;
  }}

  anos.forEach(function(ano) {{ selAno.add(new Option(ano, ano)); }});

  function preencherMeses() {{
    var ano = selAno.value;
    selMes.innerHTML = '';
    Object.keys(ANTERIORES[ano] || {{}})
      .sort(function(a, b) {{ return Number(b) - Number(a); }})
      .forEach(function(mes) {{
        var pacote  = ANTERIORES[ano][mes];
        var nomeMes = pacote.label.split('/')[0];
        var qtd     = (pacote.pontos || []).length;
        selMes.add(new Option(nomeMes + ' (' + qtd + ')', mes));
      }});
  }}

  selAno.addEventListener('change', preencherMeses);
  preencherMeses();
}}

function mostrarMesAnterior(anoBusca, mesBusca, siteAbrir) {{
  var ano    = anoBusca  || document.getElementById('sel-ano-ant').value;
  var mes    = mesBusca  || document.getElementById('sel-mes-ant').value;
  var pacote = ANTERIORES[ano] && ANTERIORES[ano][mes];

  if (!pacote || !(pacote.pontos || []).length) {{
    mostrarMensagem('msg-ant', 'Nenhuma atividade encontrada nesse período.');
    return;
  }}

  modoAnterior = true;
  mostrarMensagem('msg-ant', '');
  limparHistorico();
  setAtualVisivel(false);
  map.closePopup();

  var markerParaAbrir = null;

  pacote.pontos.forEach(function(p) {{
    var d  = dadosTipo(p);
    var mk = L.marker([p.lat, p.lon], {{icon: d.icon}}).addTo(map);
    mk.bindTooltip(d.label + p.id, {{sticky: true, className: d.cls}});
    mk.bindPopup(popupFixo(p, d.status));
    registrarSiteBusca(p.id, mk, p.lat, p.lon, null);

    if (siteAbrir && normalizarBuscaSite(p.id) === siteAbrir && markerParaAbrir === null) {{
      markerParaAbrir = {{marker: mk, lat: p.lat, lon: p.lon}};
    }}
    mkHist.push(mk);
  }});

  var concluidasHist = pacote.pontos.filter(function(p) {{ return p.tipo === 'concluida'; }}).length;
  document.getElementById('cnt-pend').textContent = '0';
  document.getElementById('cnt-conc').textContent = concluidasHist;
  mostrarMensagem('msg-ant', 'Carregado: ' + pacote.label + ' — ' + pacote.pontos.length + ' atividades.');

  if (mkHist.length) {{
    setTimeout(function() {{
      map.invalidateSize();
      if (markerParaAbrir) {{
        map.flyTo([markerParaAbrir.lat, markerParaAbrir.lon],
                  Math.max(map.getZoom(), 13), {{animate: true, duration: .85}});
        setTimeout(function() {{ markerParaAbrir.marker.openPopup(); }}, 900);
      }} else {{
        var grp = L.featureGroup(mkHist);
        map.fitBounds(grp.getBounds().pad(0.16));
      }}
    }}, 80);
  }}
}}

function voltarMapaAtual() {{
  modoAnterior = false;
  limparHistorico();
  setAtualVisivel(true);
  mostrarAba('atual');
  map.closePopup();
}}

/* ── Event listeners ───────────────────────────────────────── */

document.getElementById('chk-conc').addEventListener('change',  function() {{ if (!modoAnterior) toggleCamada(mkConc,   this.checked); }});
document.getElementById('chk-impr').addEventListener('change',  function() {{ if (!modoAnterior) toggleCamada(mkImpr,   this.checked); }});
document.getElementById('chk-canc').addEventListener('change',  function() {{ if (!modoAnterior) toggleCamada(mkCanc,   this.checked); }});
document.getElementById('chk-hotel').addEventListener('change', function() {{ toggleCamada(mkHoteis, this.checked); }});
document.getElementById('chk-aguard').addEventListener('change',function() {{ if (!modoAnterior) toggleCamada(mkAguard, this.checked); }});

document.getElementById('btn-busca-site').addEventListener('click', buscarSiteNoMapa);
document.getElementById('busca-site').addEventListener('keydown', function(ev) {{ if (ev.key === 'Enter') buscarSiteNoMapa(); }});
document.getElementById('busca-site').addEventListener('input',   function()    {{ this.value = this.value.toUpperCase(); mostrarMensagemBusca(''); }});

document.getElementById('tab-atual').addEventListener('click',       voltarMapaAtual);
document.getElementById('tab-anteriores').addEventListener('click',  function() {{ mostrarAba('anteriores'); }});
document.getElementById('btn-ant-ok').addEventListener('click',      function() {{ mostrarMesAnterior(); }});
document.getElementById('btn-voltar-atual').addEventListener('click', voltarMapaAtual);
document.getElementById('fechar-busca-global').addEventListener('click', fecharBuscaGlobal);

/* Fechar painel no mobile ao clicar no mapa */
var painel       = document.getElementById('painel');
var togglePainel = document.getElementById('painel-toggle');
togglePainel.addEventListener('click', function(ev) {{ ev.stopPropagation(); document.body.classList.add('painel-aberto'); }});
painel.addEventListener('click', function(ev) {{ ev.stopPropagation(); }});
map.on('click', function() {{ document.body.classList.remove('painel-aberto'); }});

/* ── Init ──────────────────────────────────────────────────── */

redesenharRota();
popularMesesAnteriores();

var todosMarcadores = marcadores.map(function(m) {{ return m.marker; }});
todosMarcadores.push(mkPartida);
if (todosMarcadores.length > 0) {{
  var grp = L.featureGroup(todosMarcadores);
  map.fitBounds(grp.getBounds().pad(0.1));
}}

setTimeout(function() {{ map.invalidateSize(); }}, 400);
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

    sha_atual = None
    try:
        r = requests.get(api_url, headers=headers, timeout=10)
        if r.status_code == 200:
            sha_atual = r.json().get("sha")
    except Exception as e:
        print(f"   ❌ Erro ao consultar GitHub: {e}")
        return

    agora   = datetime.now().strftime("%d/%m/%Y %H:%M")
    payload = {
        "message": f"DT 3.0 — Mapa atualizado em {agora}",
        "content": conteudo_b64,
        "branch":  "main",
    }
    if sha_atual:
        payload["sha"] = sha_atual

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
# MODO MAPA
# ============================================================

def main_mapa():
    print("\n" + "=" * 60)
    print("  DT 3.0 — Atualizar Mapa")
    print("=" * 60)

    if not CREDS_PATH.exists():
        print(f"\n❌ credentials.json não encontrado em:\n   {BASE_DIR}")
        input("\nPressione ENTER para sair...")
        return

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
    df_aguardando = df_sheets[df_sheets["STATUS"].str.strip() == ST_AGUARDANDO].copy()
    df_pendentes  = df_sheets[
        ~mask_status_fixos_mapa(df_sheets) &
        (df_sheets["STATUS"].str.strip() != ST_AGUARDANDO)
    ].copy()
    df_pendentes  = df_pendentes.dropna(subset=["LAT", "LONG"])
    df_pendentes  = df_pendentes[(df_pendentes["LAT"] != 0) & (df_pendentes["LONG"] != 0)]

    print(f"   Concluidas/Improdutivas : {len(df_fixas)}")
    print(f"   Pendentes na rota       : {len(df_pendentes)}")
    print(f"   Aguardando              : {len(df_aguardando)}")

    lat0, lon0, cidade0 = determinar_ponto_inicio(df_sheets)

    print("\n[3/3] Gerando mapa e publicando...")
    print("   → Mapa interativo...")
    historico_meses = carregar_meses_anteriores(sheet, ws.title)
    gerar_mapa(df_pendentes, lat0, lon0, cidade0,
               df_fixas=df_fixas, df_aguardando=df_aguardando,
               historico_meses=historico_meses)

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

    if not CREDS_PATH.exists():
        print(f"\n❌ credentials.json não encontrado em:\n   {BASE_DIR}")
        print("   Veja o arquivo GUIA_API_GOOGLE.txt para configurar.")
        input("\nPressione ENTER para sair...")
        return

    if not BASE_4G_PATH.exists():
        print(f"\n❌ Base_4G.xlsx não encontrado em:\n   {BASE_DIR}")
        input("\nPressione ENTER para sair...")
        return

    print("\n[1/7] Carregando Base 4G...")
    df_base = pd.read_excel(BASE_4G_PATH, engine="openpyxl")
    print(f"   {len(df_base):,} registros.")

    print("\n[2/7] Conectando ao Google Sheets...")
    try:
        client = conectar_sheets()
        sheet  = client.open_by_key(SHEET_ID)
        ws     = obter_aba_vigente(sheet)
    except Exception as e:
        print(f"   ❌ Erro: {e}")
        input("\nPressione ENTER para sair...")
        return

    print("\n[3/7] Lendo atividades da planilha...")
    df_sheets = ler_atividades_sheets(ws)
    print(f"   {len(df_sheets)} atividades encontradas.")

    if df_sheets.empty:
        print("   ⚠️  Planilha vazia. Continue assim que houver dados.")

    print("\n[4/7] Determinando ponto de partida...")
    lat0, lon0, cidade0 = determinar_ponto_inicio(df_sheets)

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
            df_novas = df_novas.drop_duplicates(subset=["SITE"], keep="first")
            print(f"   Total: {len(df_novas)} novas atividades únicas.")
    else:
        print("   Nenhuma nova atividade. Continuando sem novas.")

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

    sites_originais  = set(df_sheets["SITE"].str.upper())
    sites_no_pool    = set(df_pool_sheets["SITE"].str.upper())

    COLS = ["DEMANDA", "INTEGRAÇÃO", "SITE", "UF", "TEC",
            "LAT", "LONG", "CIDADE", "2G|3G|4G", "5G", "STATUS",
            "CONCLUIDO", "HOTEL"]

    if not df_novas.empty:
        df_nov = df_novas[~df_novas["SITE"].str.upper().isin(sites_no_pool)].copy()
        for col in ("CONCLUIDO", "HOTEL"):
            if col not in df_nov.columns:
                df_nov[col] = ""
        df_novas_filtradas = df_nov[COLS].copy()
    else:
        df_novas_filtradas = pd.DataFrame(columns=COLS)

    frames = [f for f in [df_pool_sheets[COLS], df_novas_filtradas] if not f.empty]

    if not frames:
        print("   ⚠️  Nenhuma atividade para otimizar.")
        input("\nPressione ENTER para sair...")
        return

    df_pool = pd.concat(frames, ignore_index=True)
    print(f"   Pool: {len(df_pool)} atividades ({len(df_pool_sheets)} existentes + {len(df_novas_filtradas)} novas)")

    rota_final = otimizar_rota(df_pool, lat0, lon0)

    ativ_path = BASE_DIR / "ATIVIDADES_GERADAS.xlsx"
    rota_final.to_excel(ativ_path, index=False)
    print(f"   ATIVIDADES_GERADAS.xlsx salvo.")

    print("\n[7/7] Gerando saídas...")

    print("   → Relatório de texto...")
    gerar_relatorio(rota_final, df_base, PASTA_OUT)

    print("   → Mapa interativo...")
    historico_meses = carregar_meses_anteriores(sheet, ws.title)
    gerar_mapa(rota_final, lat0, lon0, cidade0,
               df_fixas=df_fixas, df_aguardando=df_aguardando,
               historico_meses=historico_meses)

    print("   → Publicando mapa no GitHub...")
    publicar_mapa_github(BASE_DIR / "MAPA_ROTAS.html")

    print("   → Atualizando Google Sheets...")
    atualizar_sheets(ws, df_fixas, rota_final, sites_originais, df_aguardando)

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