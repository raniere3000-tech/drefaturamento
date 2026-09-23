import pandas as pd
import json
import re
import os
from datetime import datetime, timedelta

# =========================================================
# CAMINHOS
# =========================================================
ARQUIVO_FATURAMENTO = "/workspaces/drefaturamento/docs/comercial.xlsx"
ARQUIVO_ARRECADACAO = "/workspaces/drefaturamento/docs/Arrecadação.xlsx"
ARQUIVO_DESCONTO = "/workspaces/drefaturamento/docs/Desconto arrec.xlsx"
ARQUIVO_SAIDA = "/workspaces/drefaturamento/docs/DRE_Final.html"
ARQUIVO_INDEX = "index.html"
SHEET_FATURAMENTO = "Export"

COL_SITUACAO_AUDITORIA = "BASEINATIVACAO.SITUACAO_FINAL"
SITUACOES_AUDITORIA = {"ATIVA FATURANDO", "CORTADA"}

MAP_FRENTE_AGUA = {
    "CORTE DE ÁGUA": "Corte de Água",
    "RELIGAÇÕES": "Religações",
    "LIGAÇÕES DE ÁGUA": "Ligações de Água",
    "SANÇÕES": "Sanções",
    "OUTROS ÁGUA": "Outros Água",
    "VENDA E ANÁLISE DE ÁGUA": "Venda e Análise de Água",
}
MAP_FRENTE_ESGOTO = {
    "LIGAÇÕES DE ESGOTO": "Ligações de Esgoto",
}

COLS_NUM_FAT = [
    "Faturamento bruto direta agua", "Faturamento bruto direta esgoto",
    "Faturamento bruto indireta agua", "Faturamento bruto indireta esgoto",
    "Faturamento bruto indiretas total", "Qtd eco fat agua", "Qtd eco fat esgoto",
    "Vol faturado agua direto M³", "Vol faturado esgoto direto M³",
    "Volume medido de agua m³", "R$ Cancelamento total", "R$ Faturamento total liquido",
]

MESES_LABEL = ["Jan", "Fev", "Mar", "Abr", "Mai", "Jun",
               "Jul", "Ago", "Set", "Out", "Nov", "Dez"]
MESES_PT = {
    "JAN": 1, "FEB": 2, "FEV": 2, "MAR": 3, "APR": 4, "ABR": 4, "MAY": 5, "MAI": 5,
    "JUN": 6, "JUL": 7, "AUG": 8, "AGO": 8, "SEP": 9, "SET": 9,
    "OCT": 10, "OUT": 10, "NOV": 11, "DEC": 12, "DEZ": 12,
}


# =========================================================
# FUNÇÕES DE APOIO — DATAS E VALORES
# =========================================================
def parse_mes_ano(valor):
    if isinstance(valor, (pd.Timestamp, datetime)):
        mes_num, ano = valor.month, valor.year
    elif isinstance(valor, (int, float)) and not pd.isna(valor) and 20000 < float(valor) < 60000:
        dt = datetime(1899, 12, 30) + timedelta(days=float(valor))
        mes_num, ano = dt.month, dt.year
    else:
        s = str(valor).strip()
        s = re.sub(r"\s*/\s*", "/", s)
        partes = [p.strip() for p in s.split("/") if p.strip() != ""]

        if len(partes) < 2:
            raise ValueError(f"Formato de data não reconhecido: '{valor}'")

        primeira = partes[0]

        if re.match(r"^[A-Za-zçÇãÃéÉ]+$", primeira):
            mes_txt = primeira[:3].upper()
            mes_num = MESES_PT.get(mes_txt)
            ano = int(re.sub(r"\D", "", partes[-1]))
        else:
            mes_num = int(re.sub(r"\D", "", primeira))
            ano = int(re.sub(r"\D", "", partes[-1]))

        if ano < 100:
            ano += 2000

    if not mes_num or mes_num < 1 or mes_num > 12:
        raise ValueError(f"Mês inválido ao converter valor: '{valor}'")
    if ano < 2000 or ano > 2100:
        raise ValueError(f"Ano fora do intervalo esperado: '{valor}'")

    label = f"{MESES_LABEL[mes_num - 1]}/{ano}"
    ordem = ano * 100 + mes_num
    return label, ordem


def limpar_valor_moeda(valor):
    if pd.isna(valor):
        return 0.0
    if isinstance(valor, (int, float)):
        return float(valor)
    s = str(valor).strip()
    negativo = "-" in s
    s = re.sub(r"[^\d.,]", "", s).replace(",", "")
    if s == "":
        return 0.0
    try:
        num = float(s)
    except ValueError:
        num = 0.0
    return -abs(num) if negativo else num


# =========================================================
# CARGA — FATURAMENTO
# =========================================================
def converter_referencia_fat(valor):
    if pd.isna(valor):
        return None
    if isinstance(valor, (pd.Timestamp, datetime)):
        return valor
    if isinstance(valor, (int, float)):
        try:
            return datetime(1899, 12, 30) + timedelta(days=float(valor))
        except Exception:
            return None
    s = str(valor).strip()
    if s == "" or s.lower() == "nan":
        return None
    try:
        f = float(s.replace(",", "."))
        if 20000 < f < 60000:
            return datetime(1899, 12, 30) + timedelta(days=f)
    except ValueError:
        pass
    dt = pd.to_datetime(s, dayfirst=True, errors="coerce")
    if pd.isna(dt):
        dt = pd.to_datetime(s, dayfirst=False, errors="coerce")
    return dt if pd.notna(dt) else None


def carregar_faturamento(caminho, sheet):
    df = pd.read_excel(caminho, sheet_name=sheet)
    df = df[~df["CIDADE"].astype(str).str.contains("Filtros aplicados", na=False)]
    df = df[df["CIDADE"].astype(str).str.upper() != "TOTAL"]

    for col in COLS_NUM_FAT:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    df["DSC_CLASSE"] = df["DSC_CLASSE"].astype(str).str.upper().str.strip()
    df["CIDADE"] = df["CIDADE"].astype(str).str.upper().str.strip()

    if COL_SITUACAO_AUDITORIA in df.columns:
        df["_SITUACAO_AUDITORIA"] = df[COL_SITUACAO_AUDITORIA].astype(str).str.upper().str.strip()
    else:
        print(f"[AVISO] Coluna '{COL_SITUACAO_AUDITORIA}' não encontrada na planilha de Faturamento. "
              f"O filtro de Visão Auditoria não terá efeito.")
        df["_SITUACAO_AUDITORIA"] = ""

    ref_cols = [c for c in df.columns if str(c).startswith("Referência")]
    melhor_serie, melhor_qtd = None, -1
    for c in ref_cols:
        serie = df[c].apply(converter_referencia_fat)
        qtd = serie.notna().sum()
        if qtd > melhor_qtd:
            melhor_qtd, melhor_serie = qtd, serie

    if melhor_serie is None:
        melhor_serie = pd.Series([None] * len(df), index=df.index)

    df["_DATA_REF"] = pd.to_datetime(melhor_serie, errors="coerce")

    antes = len(df)
    df = df[df["_DATA_REF"].notna()].copy()
    if antes != len(df):
        print(f"[AVISO] Faturamento: {antes - len(df)} linha(s) descartadas por data de referência inválida.")

    df["_MES_ANO"] = df["_DATA_REF"].apply(lambda d: f"{MESES_LABEL[d.month - 1]}/{d.year}")
    df["_MES_ORDEM"] = df["_DATA_REF"].apply(lambda d: d.year * 100 + d.month)

    resumo_2022 = (df[df["_MES_ANO"].str.endswith("2022", na=False)]
                   .groupby("_MES_ANO")["R$ Faturamento total liquido"].sum())
    print("[DEBUG] Faturamento Líquido em 2022 por mês:")
    print(resumo_2022)

    qtd_auditoria = df["_SITUACAO_AUDITORIA"].isin(SITUACOES_AUDITORIA).sum()
    print(f"[DEBUG] Linhas elegíveis para Visão Auditoria (Ativa Faturando / Cortada): {qtd_auditoria} de {len(df)}")

    return df


def calcular_dre_fat(df_sub):
    r = {}
    r["direta_agua"] = df_sub["Faturamento bruto direta agua"].sum()
    r["direta_esgoto"] = df_sub["Faturamento bruto direta esgoto"].sum()
    r["diretas_totais"] = r["direta_agua"] + r["direta_esgoto"]

    frentes_agua = {nome: df_sub.loc[df_sub["DSC_CLASSE"] == classe, "Faturamento bruto indireta agua"].sum()
                     for classe, nome in MAP_FRENTE_AGUA.items()}
    frentes_esgoto = {nome: df_sub.loc[df_sub["DSC_CLASSE"] == classe, "Faturamento bruto indireta esgoto"].sum()
                       for classe, nome in MAP_FRENTE_ESGOTO.items()}

    r["frentes_agua"] = frentes_agua
    r["frentes_esgoto"] = frentes_esgoto
    r["indireta_agua_total"] = sum(frentes_agua.values())
    r["indireta_esgoto_total"] = sum(frentes_esgoto.values())
    r["indiretas_totais"] = r["indireta_agua_total"] + r["indireta_esgoto_total"]
    r["faturamento_bruto"] = r["diretas_totais"] + r["indiretas_totais"]

    r["eco_agua"] = df_sub["Qtd eco fat agua"].sum()
    r["eco_esgoto"] = df_sub["Qtd eco fat esgoto"].sum()
    r["vol_fat_agua"] = df_sub["Vol faturado agua direto M³"].sum()
    r["vol_fat_esgoto"] = df_sub["Vol faturado esgoto direto M³"].sum()
    r["vol_medio_agua"] = df_sub["Volume medido de agua m³"].sum()

    r["tarifa_media_agua"] = (r["direta_agua"] / r["vol_fat_agua"]) if r["vol_fat_agua"] else 0
    r["tarifa_media_esgoto"] = (r["direta_esgoto"] / r["vol_fat_esgoto"]) if r["vol_fat_esgoto"] else 0

    r["cancelamento"] = df_sub["R$ Cancelamento total"].sum()
    r["faturamento_liquido"] = df_sub["R$ Faturamento total liquido"].sum()
    return r


# =========================================================
# CARGA — ARRECADAÇÃO + DESCONTO
# =========================================================
def carregar_arrecadacao(caminho):
    df = pd.read_excel(caminho, sheet_name="Query1")
    df = df[df["IsGrandTotalRowTotal"].astype(str) != "True"].copy()
    df = df[df["Cidade"].notna()]

    df["Cidade"] = df["Cidade"].astype(str).str.upper().str.strip()
    df["Subcategoria"] = df["Subcategoria"].astype(str).str.upper().str.strip()

    labels, ordens = [], []
    for idx, v in df["Mes/ano"].items():
        try:
            label, ordem = parse_mes_ano(v)
        except Exception as e:
            print(f"[AVISO] Linha {idx} com data inválida ('{v}'): {e}")
            label, ordem = None, None
        labels.append(label)
        ordens.append(ordem)

    df["_MES_ANO"] = labels
    df["_MES_ORDEM"] = ordens

    antes = len(df)
    df = df[df["_MES_ANO"].notna()]
    if antes != len(df):
        print(f"[AVISO] Arrecadação: {antes - len(df)} linha(s) descartadas por data inválida.")

    for col in ["Arrecadacao_Acumulada", "Qtd_clientes_pagantes_acumulado",
                "QTD_Contas_pagas_Arrecadação_acumulada"]:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    dup = df.duplicated(subset=["Cidade", "Subcategoria", "_MES_ANO"], keep=False)
    if dup.any():
        print(f"[AVISO] {dup.sum()} linha(s) duplicadas (Cidade+Subcategoria+Mês) na Arrecadação — verifique a base original.")

    resumo_2022 = (df[df["_MES_ANO"].str.endswith("2022", na=False)]
                   .groupby("_MES_ANO")["Arrecadacao_Acumulada"].sum())
    print("[DEBUG] Totais de Arrecadação em 2022 por mês:")
    print(resumo_2022)

    return df


def carregar_desconto(caminho):
    """
    Carrega os descontos da planilha e GARANTE que todos os valores fiquem NEGATIVOS,
    independente de como estejam preenchidos na planilha original (com ou sem sinal).
    """
    df = pd.read_excel(caminho, sheet_name="Planilha1")
    df = df[df["Mes/ano"].notna()]
    desconto_map = {}
    ajustados = 0

    for idx, row in df.iterrows():
        try:
            label, ordem = parse_mes_ano(row["Mes/ano"])
        except Exception as e:
            print(f"[AVISO] Desconto - linha {idx} com data inválida ('{row['Mes/ano']}'): {e}")
            continue

        valor_bruto = limpar_valor_moeda(row["Desconto"])

        valor_normalizado = -abs(valor_bruto)
        if valor_bruto > 0:
            ajustados += 1
            print(f"[AVISO] Desconto de {label} estava positivo ({valor_bruto:.2f}) "
                  f"e foi convertido para negativo ({valor_normalizado:.2f}).")

        desconto_map[label] = {"valor": valor_normalizado, "ordem": ordem}

    if ajustados > 0:
        print(f"[RESUMO] Total de {ajustados} valor(es) de desconto normalizados para negativo.")
    else:
        print("[RESUMO] Todos os valores de desconto já estavam negativos. Nenhum ajuste necessário.")

    return desconto_map


# =========================================================
# MONTAGEM DO DATASET UNIFICADO
# =========================================================
def montar_dataset_unificado(df_fat, df_arr, desconto_map):
    dataset = {}
    meses_map = {}
    cidades_set = set()

    df_fat_aud = df_fat[df_fat["_SITUACAO_AUDITORIA"].isin(SITUACOES_AUDITORIA)]

    for (cidade, mes), grupo in df_fat.groupby(["CIDADE", "_MES_ANO"]):
        ordem = grupo["_MES_ORDEM"].iloc[0]
        meses_map[mes] = ordem
        cidades_set.add(cidade)
        dataset.setdefault(cidade, {}).setdefault(mes, {})
        dataset[cidade][mes]["fat"] = calcular_dre_fat(grupo)

    for (cidade, mes), grupo in df_fat_aud.groupby(["CIDADE", "_MES_ANO"]):
        ordem = grupo["_MES_ORDEM"].iloc[0]
        meses_map[mes] = ordem
        cidades_set.add(cidade)
        dataset.setdefault(cidade, {}).setdefault(mes, {})
        dataset[cidade][mes]["fat_aud"] = calcular_dre_fat(grupo)

    for (cidade, mes), grupo in df_arr.groupby(["Cidade", "_MES_ANO"]):
        ordem = grupo["_MES_ORDEM"].iloc[0]
        meses_map[mes] = ordem
        cidades_set.add(cidade)
        dataset.setdefault(cidade, {}).setdefault(mes, {})

        dataset[cidade][mes]["arr"] = {
            "arrecadacao": float(grupo["Arrecadacao_Acumulada"].sum()),
            "clientes": float(grupo["Qtd_clientes_pagantes_acumulado"].sum()),
            "contas": float(grupo["QTD_Contas_pagas_Arrecadação_acumulada"].sum()),
        }

    for mes, info in desconto_map.items():
        meses_map[mes] = info["ordem"]

    meses_ordenados = [m for m, _ in sorted(meses_map.items(), key=lambda x: x[1])]
    cidades = sorted(cidades_set)
    return dataset, meses_ordenados, cidades


# =========================================================
# ESTRUTURA DAS LINHAS DO DRE
# =========================================================
LINHAS_DRE = [
    ("FATURAMENTO BRUTO", "faturamento_bruto", 0, True, False, "moeda", None),
    ("DIRETAS TOTAIS", "diretas_totais", 0, True, False, "moeda", "bruto"),
    ("Diretas Água", "direta_agua", 1, False, False, "moeda", "diretas|bruto"),
    ("Diretas Esgoto", "direta_esgoto", 1, False, False, "moeda", "diretas|bruto"),
    ("INDIRETAS TOTAIS", "indiretas_totais", 0, True, False, "moeda", "bruto"),
    ("INDIRETA ÁGUA", "indireta_agua_total", 0, True, True, "moeda", "indiretas_sub|bruto"),
    ("Corte de Água", "F_Corte de Água", 1, False, False, "moeda", "ind_agua|indiretas_sub|bruto"),
    ("Religações", "F_Religações", 1, False, False, "moeda", "ind_agua|indiretas_sub|bruto"),
    ("Ligações de Água", "F_Ligações de Água", 1, False, False, "moeda", "ind_agua|indiretas_sub|bruto"),
    ("Sanções", "F_Sanções", 1, False, False, "moeda", "ind_agua|indiretas_sub|bruto"),
    ("Outros Água", "F_Outros Água", 1, False, False, "moeda", "ind_agua_extra|indiretas_sub|bruto"),
    ("Venda e Análise de Água", "F_Venda e Análise de Água", 1, False, False, "moeda", "ind_agua_extra|indiretas_sub|bruto"),
    ("INDIRETA ESGOTO", "indireta_esgoto_total", 0, True, True, "moeda", "indiretas_sub|bruto"),
    ("Ligações de Esgoto", "E_Ligações de Esgoto", 1, False, False, "moeda", "ind_esgoto|indiretas_sub|bruto"),
    ("FAT ECONOMIAS ÁGUA (qtd)", "eco_agua", 0, True, False, "num", "extras"),
    ("FAT ECONOMIAS ESGOTO (qtd)", "eco_esgoto", 0, True, False, "num", "extras"),
    ("VOL. FAT. ÁGUA (m³)", "vol_fat_agua", 0, True, False, "num", "extras"),
    ("VOL. FAT. ESGOTO (m³)", "vol_fat_esgoto", 0, True, False, "num", "extras"),
    ("VOLUME MEDIDO DE ÁGUA (m³)", "vol_medio_agua", 0, True, False, "num", "extras"),
    ("TARIFA MÉDIA DE ÁGUA (R$/m³)", "tarifa_media_agua", 0, True, False, "moeda4", "extras"),
    ("TARIFA MÉDIA DE ESGOTO (R$/m³)", "tarifa_media_esgoto", 0, True, False, "moeda4", "extras"),
    ("CANCELAMENTO", "cancelamento", 0, True, False, "moeda", None),
    ("FATURAMENTO LÍQUIDO C/ REAJUSTE", "faturamento_liquido", 0, True, False, "moeda", None),

    # ================= ARRECADAÇÃO =================
    ("ARRECADAÇÃO REAL (Arrecadação + Desconto)", "arrecadacao_real", 0, True, False, "moeda", None),
    ("DESCONTO", "desconto", 0, True, False, "moeda", None),
    ("ARRECADAÇÃO (Sem Desconto)", "arrecadacao", 0, True, False, "moeda", None),

    ("ARRECADAÇÃO / FATURAMENTO LÍQUIDO (%)", "pct_arrec_fat", 0, True, False, "pct", "indicador"),

    ("QTD. CLIENTES PAGANTES", "clientes", 0, True, False, "num", None),
    ("QTD. CONTAS PAGAS", "contas", 0, True, False, "num", None),
    ("TICKET MÉDIO (R$/cliente)", "ticket_medio", 0, True, False, "moeda2", None),
]

GRUPOS_OCULTOS_PADRAO = ["ind_agua_extra"]

GRUPOS_TOGGLE = {
    "diretas": "DIRETAS TOTAIS",
    "ind_agua": "INDIRETA ÁGUA",
    "ind_esgoto": "INDIRETA ESGOTO",
}

GRUPO_COMBO = {
    "INDIRETAS TOTAIS": ["ind_agua", "ind_esgoto", "indiretas_sub"],
    "FATURAMENTO BRUTO": ["bruto"],
}

# =========================================================
# GRÁFICOS
# =========================================================
GRAFICOS = [
    ("DIRETAS TOTAIS", [("Diretas Água", "direta_agua"), ("Diretas Esgoto", "direta_esgoto")], "moeda"),
    ("INDIRETA ÁGUA", [
        ("Corte de Água", "F_Corte de Água"), ("Religações", "F_Religações"),
        ("Ligações de Água", "F_Ligações de Água"), ("Sanções", "F_Sanções"),
        ("Outros Água", "F_Outros Água"),
    ], "moeda"),
    ("INDIRETA ESGOTO", [("Ligações de Esgoto", "E_Ligações de Esgoto")], "moeda"),
    ("FAT. ECONOMIAS ÁGUA E ESGOTO", [("Fat. Economias Água", "eco_agua"), ("Fat. Economias Esgoto", "eco_esgoto")], "num"),
    ("VOL. FAT. ÁGUA E ESGOTO", [("Vol. Fat. Água", "vol_fat_agua"), ("Vol. Fat. Esgoto", "vol_fat_esgoto")], "num"),
    ("VOLUME MEDIDO DE ÁGUA", [("Volume Medido de Água", "vol_medio_agua")], "num"),
    ("TARIFA MÉDIA DE ÁGUA E ESGOTO", [("Tarifa Média de Água", "tarifa_media_agua"), ("Tarifa Média de Esgoto", "tarifa_media_esgoto")], "moeda4"),
    ("CANCELAMENTO", [("Cancelamento", "cancelamento")], "moeda"),
    ("FATURAMENTO LÍQUIDO C/ REAJUSTE", [("Faturamento Líquido c/ Reajuste", "faturamento_liquido")], "moeda"),
    ("ARRECADAÇÃO REAL", [("Arrecadação Real (R$)", "arrecadacao_real")], "moeda"),
    ("CLIENTES PAGANTES E CONTAS PAGAS", [
        ("Clientes Pagantes", "clientes"),
        ("Contas Pagas", "contas"),
    ], "num"),
    ("TICKET MÉDIO (R$/cliente)", [("Ticket Médio", "ticket_medio")], "moeda2"),
]

# Notas explicativas fixas por gráfico + mês
NOTAS_GRAFICOS = {
    "CANCELAMENTO": {
        "mes": "Abr/2026",
        "texto": "Tivemos um erro no lançamento da indireta de esgoto R$ 18.757.731,26."
    },
    "INDIRETA ESGOTO": {
        "mes": "Abr/2026",
        "texto": "Tivemos um erro no lançamento da indireta de esgoto R$ 18.757.731,26."
    },
}


# =========================================================
# GERAÇÃO DO HTML ÚNICO (TABELA + GRÁFICOS COM ABAS)
# =========================================================
def gerar_html_unico(dataset, meses, cidades, descontos, caminho_saida, arquivo_index):
    dataset_json = json.dumps(dataset, ensure_ascii=False)
    meses_json = json.dumps(meses, ensure_ascii=False)
    cidades_json = json.dumps(cidades, ensure_ascii=False)
    descontos_json = json.dumps(descontos, ensure_ascii=False)
    linhas_json = json.dumps(LINHAS_DRE, ensure_ascii=False)
    grupos_ocultos_json = json.dumps(GRUPOS_OCULTOS_PADRAO, ensure_ascii=False)
    grupos_toggle_json = json.dumps(GRUPOS_TOGGLE, ensure_ascii=False)
    grupo_combo_json = json.dumps(GRUPO_COMBO, ensure_ascii=False)
    graficos_json = json.dumps(GRAFICOS, ensure_ascii=False)
    notas_json = json.dumps(NOTAS_GRAFICOS, ensure_ascii=False)

    html = f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<title>DRE Final — Faturamento e Arrecadação</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4"></script>
<style>
  body {{ font-family: Arial, sans-serif; margin: 20px; background:#f5f6f8; }}
  h1 {{ font-size: 20px; color:#1F4E78; margin-bottom:6px; }}

  .auditoria-box {{ display:flex; align-items:center; gap:10px; margin-bottom:14px; }}
  .auditoria-label {{ font-size:13px; color:#1F4E78; font-weight:600; }}
  .switch {{ position: relative; display: inline-block; width: 46px; height: 24px; }}
  .switch input {{ opacity: 0; width: 0; height: 0; }}
  .slider {{ position: absolute; cursor: pointer; top: 0; left: 0; right: 0; bottom: 0;
             background-color: #ccc; transition: .3s; border-radius: 24px; }}
  .slider:before {{ position: absolute; content: ""; height: 18px; width: 18px; left: 3px; bottom: 3px;
                     background-color: white; transition: .3s; border-radius: 50%; }}
  input:checked + .slider {{ background-color: #1F4E78; }}
  input:checked + .slider:before {{ transform: translateX(22px); }}
  .auditoria-tag {{ font-size:11px; padding:2px 8px; border-radius:10px; background:#eef6ee; color:#1c6b3d;
                     font-weight:600; display:none; }}
  .auditoria-box.ativo .auditoria-tag {{ display:inline-block; }}

  #abas {{ display:flex; gap:8px; margin-bottom:14px; }}
  .aba-btn {{ padding:9px 18px; border:none; border-radius:6px 6px 0 0; background:#dfe6ec; color:#1F4E78;
              font-weight:600; font-size:13px; cursor:pointer; transition: background .25s ease, color .25s ease; }}
  .aba-btn.ativa {{ background:#1F4E78; color:#fff; }}

  #filtros {{ background:#fff; padding:12px 15px; border-radius:8px; margin-bottom:14px;
              box-shadow:0 1px 3px rgba(0,0,0,.1); display:flex; align-items:center; flex-wrap:wrap; gap:0; }}
  .btn {{ padding:8px 14px; border:none; border-radius:5px; background:#1F4E78; color:#fff;
          cursor:pointer; font-size:12.5px; text-decoration:none; display:inline-block; transition: background .2s ease; }}
  .btn.secundario {{ background:#5a7a99; }}
  .btn-nav {{ background:#fff; color:#1F4E78; font-weight:600; box-shadow:0 4px 10px rgba(0,0,0,.15);
              border:1px solid #ddd; margin-left:auto; }}
  .btn-nav:hover {{ background:#f0f4f8; }}
  .dropdown {{ position: relative; display: inline-block; margin-right: 10px; }}
  .dropdown-content {{
    display: none; position: absolute; background: #fff; min-width: 230px;
    box-shadow: 0 4px 10px rgba(0,0,0,.15); border-radius: 6px; z-index: 10;
    padding: 8px 0; max-height: 280px; overflow-y: auto; margin-top:4px;
  }}
  .dropdown-content label, .dropdown-content div.opcao-modo {{
    display: block; padding: 6px 14px; font-size: 12.5px; cursor: pointer; white-space: nowrap;
  }}
  .dropdown-content div.opcao-modo:hover {{ background:#f0f4f8; }}
  .dropdown-content hr {{ margin:6px 0; border:none; border-top:1px solid #eee; }}
  .dropdown.aberto .dropdown-content {{ display: block; }}
  .dropbtn {{ min-width: 150px; text-align: left; }}
  .intervalo-box {{ display:flex; align-items:center; gap:6px; padding:0 10px; font-size:12.5px; color:#1F4E78; }}
  .intervalo-box select {{ font-size:12.5px; padding:5px 8px; border-radius:5px; border:1px solid #ccc; }}

  table {{ border-collapse: collapse; width:100%; background:#fff; box-shadow:0 1px 3px rgba(0,0,0,.1);
           font-size:12px; }}
  th, td {{ border-bottom:1px solid #e6e6e6; padding:5px 8px; text-align:center; white-space:nowrap; }}
  th {{ background:#1F4E78; color:#fff; text-align:center; position: sticky; top:0; font-size:12px; }}
  td.item {{ text-align:left; white-space:normal; min-width:220px; }}
  tr.total td {{ font-weight:bold; background:#c9ccd1; color:#1a1a1a; }}
  tr.subtotal td {{ font-weight:bold; font-style:italic; background:#fbfcfd; color:#333; }}
  tr.nivel1 td.item {{ padding-left:26px; color:#444; font-weight:normal; }}
  td.col-total {{ font-weight:bold; }}
  .toggle-btn {{ cursor:pointer; color:#1F4E78; font-weight:bold; margin-right:6px;
                 display:inline-block; width:14px; text-align:center; }}
  #resumo {{ margin-top:8px; font-size:11.5px; color:#666; }}
  .negativo {{ color:#c0392b; }}

  #grid {{ display:grid; grid-template-columns: repeat(2, 1fr); gap:16px; }}
  .card {{ background:#fff; border-radius:8px; box-shadow:0 1px 3px rgba(0,0,0,.1); padding:12px; }}
  .card h3 {{ margin:0 0 8px 0; font-size:14px; color:#1F4E78; text-align:center; }}
  .card canvas {{ max-height:280px; }}
  .nota-grafico {{
    margin-top: 8px; font-size: 10.5px; color: #c0392b;
    background: #fdecea; border-radius: 4px; padding: 6px 8px; line-height:1.4;
  }}
  @media (max-width: 900px) {{ #grid {{ grid-template-columns: 1fr; }} }}

  /* Transição suave entre abas */
  .painel {{
    display: none;
    opacity: 0;
    transform: translateY(6px);
    transition: opacity 0.35s ease, transform 0.35s ease;
  }}
  .painel.ativo {{
    display: block;
    opacity: 1;
    transform: translateY(0);
  }}
</style>
</head>
<body>

<h1>DRE Final — Faturamento e Arrecadação</h1>

<div class="auditoria-box" id="auditoriaBox">
  <span class="auditoria-label">Visão Auditoria</span>
  <label class="switch">
    <input type="checkbox" id="chkAuditoria">
    <span class="slider"></span>
  </label>
  <span class="auditoria-tag">Filtrando: Ativa Faturando / Cortada</span>
</div>

<div id="abas">
  <button class="aba-btn ativa" id="btnAbaTabela">📋 Tabela</button>
  <button class="aba-btn" id="btnAbaGraficos">📊 Gráficos</button>
</div>

<div id="filtros">
  <div class="dropdown">
    <button class="btn dropbtn" id="btnDropCidades">Cidades ▾</button>
    <div class="dropdown-content" id="listaCidades">
      <label><input type="checkbox" id="chkTodasCidades" checked> <b>Selecionar Todas</b></label>
      <hr>
    </div>
  </div>

  <div class="dropdown">
    <button class="btn dropbtn secundario" id="btnDropModo">Modo: Mensal ▾</button>
    <div class="dropdown-content">
      <div class="opcao-modo" data-modo="mensal">Visão Mensal</div>
      <div class="opcao-modo" data-modo="anual">Visão Anual</div>
    </div>
  </div>

  <div class="intervalo-box">
    <b>Período:</b>
    De <select id="selMesInicio"></select>
    até <select id="selMesFim"></select>
  </div>

  <a class="btn btn-nav" href="{arquivo_index}">🏠 Menu Principal</a>
</div>

<!-- ===================== PAINEL TABELA ===================== -->
<div class="painel ativo" id="painelTabela">
  <table id="tabelaDRE">
    <thead><tr id="headerRow"><th style="text-align:left">Item</th></tr></thead>
    <tbody id="corpoTabela"></tbody>
  </table>
  <div id="resumo"></div>
</div>

<!-- ===================== PAINEL GRÁFICOS ===================== -->
<div class="painel" id="painelGraficos">
  <div id="grid"></div>
</div>

<script>
const dataset = {dataset_json};
const mesesTodos = {meses_json};
const cidades = {cidades_json};
const descontos = {descontos_json};
const linhas = {linhas_json};
const GRUPOS_OCULTOS_PADRAO = {grupos_ocultos_json};
const GRUPOS_TOGGLE = {grupos_toggle_json};
const GRUPO_COMBO = {grupo_combo_json};
const GRAFICOS = {graficos_json};
const NOTAS_GRAFICOS = {notas_json};

let modoVisualizacao = "mensal";
let gruposOcultos = new Set(GRUPOS_OCULTOS_PADRAO);
let idxInicio = 0;
let idxFim = mesesTodos.length - 1;
let modoAuditoria = false;
let abaAtiva = "tabela";
let chartsInstances = [];

function fmt(valor, tipo) {{
  if (valor === undefined || valor === null || isNaN(valor)) return "-";
  if (tipo === "moeda" || tipo === "moeda2") {{
    return valor.toLocaleString('pt-BR', {{minimumFractionDigits:2, maximumFractionDigits:2}});
  }}
  if (tipo === "moeda4") {{
    return valor.toLocaleString('pt-BR', {{minimumFractionDigits:2, maximumFractionDigits:4}});
  }}
  if (tipo === "pct") {{
    return valor.toLocaleString('pt-BR', {{minimumFractionDigits:2, maximumFractionDigits:2}}) + "%";
  }}
  return valor.toLocaleString('pt-BR', {{maximumFractionDigits:0}});
}}

function salvarFiltros() {{
  localStorage.setItem("dre_idxInicio", idxInicio);
  localStorage.setItem("dre_idxFim", idxFim);
  localStorage.setItem("dre_cidades", JSON.stringify(getSelecionadas()));
  localStorage.setItem("dre_auditoria", modoAuditoria ? "1" : "0");
}}

function carregarFiltros() {{
  const ini = localStorage.getItem("dre_idxInicio");
  const fim = localStorage.getItem("dre_idxFim");
  const cids = localStorage.getItem("dre_cidades");
  const aud = localStorage.getItem("dre_auditoria");
  if (ini !== null && !isNaN(parseInt(ini))) idxInicio = parseInt(ini);
  if (fim !== null && !isNaN(parseInt(fim))) idxFim = parseInt(fim);
  if (idxInicio >= mesesTodos.length) idxInicio = 0;
  if (idxFim >= mesesTodos.length) idxFim = mesesTodos.length - 1;
  if (aud === "1") modoAuditoria = true;
  if (cids) {{
    try {{
      const salvas = JSON.parse(cids);
      document.querySelectorAll(".chkCidade").forEach(c => {{ c.checked = salvas.includes(c.value); }});
    }} catch (e) {{}}
  }}
}}

function montarFiltros() {{
  const div = document.getElementById("listaCidades");
  cidades.forEach(c => {{
    const id = "chk_" + c.replace(/\\s+/g, "_");
    div.innerHTML += `<label><input type="checkbox" class="chkCidade" value="${{c}}" id="${{id}}" checked> ${{c}}</label>`;
  }});

  const selIni = document.getElementById("selMesInicio");
  const selFim = document.getElementById("selMesFim");
  mesesTodos.forEach((m, i) => {{
    selIni.innerHTML += `<option value="${{i}}">${{m}}</option>`;
    selFim.innerHTML += `<option value="${{i}}">${{m}}</option>`;
  }});

  carregarFiltros();
  selIni.value = idxInicio;
  selFim.value = idxFim;

  const chkAud = document.getElementById("chkAuditoria");
  chkAud.checked = modoAuditoria;
  document.getElementById("auditoriaBox").classList.toggle("ativo", modoAuditoria);

  chkAud.onchange = () => {{
    modoAuditoria = chkAud.checked;
    document.getElementById("auditoriaBox").classList.toggle("ativo", modoAuditoria);
    salvarFiltros(); atualizarTudo();
  }};

  selIni.onchange = () => {{
    idxInicio = parseInt(selIni.value);
    if (idxInicio > idxFim) {{ idxFim = idxInicio; selFim.value = idxFim; }}
    salvarFiltros(); atualizarTudo();
  }};
  selFim.onchange = () => {{
    idxFim = parseInt(selFim.value);
    if (idxFim < idxInicio) {{ idxInicio = idxFim; selIni.value = idxInicio; }}
    salvarFiltros(); atualizarTudo();
  }};
}}

function getSelecionadas() {{
  return Array.from(document.querySelectorAll(".chkCidade:checked")).map(el => el.value);
}}

function somarDados(cidadesSel, mes) {{
  const acc = {{}};
  let arr = {{ arrecadacao: 0, clientes: 0, contas: 0 }};

  cidadesSel.forEach(c => {{
    const d = dataset[c] && dataset[c][mes];
    if (!d) return;

    const fatUsado = modoAuditoria ? d.fat_aud : d.fat;

    if (fatUsado) {{
      for (const key in fatUsado) {{
        if (key === "frentes_agua" || key === "frentes_esgoto") {{
          for (const f in fatUsado[key]) {{
            const k = (key === "frentes_agua" ? "F_" : "E_") + f;
            acc[k] = (acc[k] || 0) + fatUsado[key][f];
          }}
        }} else {{
          acc[key] = (acc[key] || 0) + fatUsado[key];
        }}
      }}
    }}

    if (d.arr) {{
      arr.arrecadacao += d.arr.arrecadacao;
      arr.clientes += d.arr.clientes;
      arr.contas += d.arr.contas;
    }}
  }});

  acc._arr = arr;
  return acc;
}}

function anoDoMes(mes) {{ return mes.split("/")[1]; }}
function getMesesNoIntervalo() {{ return mesesTodos.slice(idxInicio, idxFim + 1); }}

function getColunas() {{
  const mesesFiltrados = getMesesNoIntervalo();
  if (modoVisualizacao === "mensal") {{
    return {{ colunas: mesesFiltrados, agrupador: (mes) => [mes] }};
  }}
  const anos = [];
  const mapaAnoMeses = {{}};
  mesesFiltrados.forEach(m => {{
    const ano = anoDoMes(m);
    if (!mapaAnoMeses[ano]) {{ mapaAnoMeses[ano] = []; anos.push(ano); }}
    mapaAnoMeses[ano].push(m);
  }});
  return {{ colunas: anos, agrupador: (ano) => mapaAnoMeses[ano] }};
}}

function somarPeriodo(cidadesSel, listaMeses) {{
  const acc = {{}};
  let arr = {{ arrecadacao: 0, clientes: 0, contas: 0, desconto: 0 }};
  listaMeses.forEach(mes => {{
    const parcial = somarDados(cidadesSel, mes);
    for (const k in parcial) {{
      if (k === "_arr") continue;
      acc[k] = (acc[k] || 0) + parcial[k];
    }}
    arr.arrecadacao += parcial._arr.arrecadacao;
    arr.clientes += parcial._arr.clientes;
    arr.contas += parcial._arr.contas;
    if (descontos[mes]) arr.desconto += descontos[mes].valor;
  }});
  arr.arrecadacao_real = arr.arrecadacao + arr.desconto;
  arr.ticket_medio = arr.clientes ? (arr.arrecadacao / arr.clientes) : 0;
  acc._arr = arr;
  return acc;
}}

function grupoVisivel(grupo) {{
  if (grupo === null) return true;
  const partes = grupo.split("|");
  return !partes.some(g => gruposOcultos.has(g));
}}

function comboOculto(gruposLista) {{ return gruposLista.every(g => gruposOcultos.has(g)); }}

function obterValor(acc, chave) {{
  if (chave === "desconto") return acc._arr.desconto;
  if (chave === "arrecadacao_real") return acc._arr.arrecadacao_real;
  if (chave === "ticket_medio") return acc._arr.ticket_medio;
  if (chave === "clientes") return acc._arr.clientes;
  if (chave === "contas") return acc._arr.contas;
  if (chave === "arrecadacao") return acc._arr.arrecadacao;
  if (chave === "pct_arrec_fat") {{
    const fatLiq = acc.faturamento_liquido || 0;
    return fatLiq !== 0 ? (acc._arr.arrecadacao_real / fatLiq) * 100 : 0;
  }}
  return acc[chave] || 0;
}}

function montarTabela() {{
  const cidadesSel = getSelecionadas();
  const {{ colunas, agrupador }} = getColunas();

  const headerRow = document.getElementById("headerRow");
  headerRow.innerHTML = '<th style="text-align:left">Item</th>';
  colunas.forEach(c => headerRow.innerHTML += `<th>${{c}}</th>`);
  headerRow.innerHTML += '<th class="col-total">Total Geral</th>';

  const corpo = document.getElementById("corpoTabela");
  corpo.innerHTML = "";

  linhas.forEach(linha => {{
    const [nome, chave, nivel, isTotal, isSub, tipo, grupo] = linha;
    if (!grupoVisivel(grupo)) return;

    const tr = document.createElement("tr");
    let classes = (isTotal ? "total " : "") + (isSub ? "subtotal " : "") + (nivel === 1 ? "nivel1 " : "");
    tr.className = classes;

    const tdItem = document.createElement("td");
    tdItem.className = "item";
    let btnToggle = null;

    if (GRUPO_COMBO[nome]) {{
      const gruposLista = GRUPO_COMBO[nome];
      btnToggle = document.createElement("span");
      btnToggle.className = "toggle-btn";
      btnToggle.textContent = comboOculto(gruposLista) ? "+" : "−";
      btnToggle.onclick = (e) => {{
        e.stopPropagation();
        if (comboOculto(gruposLista)) gruposLista.forEach(g => gruposOcultos.delete(g));
        else gruposLista.forEach(g => gruposOcultos.add(g));
        montarTabela();
      }};
    }}
    for (const g in GRUPOS_TOGGLE) {{
      if (GRUPOS_TOGGLE[g] === nome) {{
        btnToggle = document.createElement("span");
        btnToggle.className = "toggle-btn";
        const oculto = gruposOcultos.has(g);
        btnToggle.textContent = oculto ? "+" : "−";
        btnToggle.onclick = (e) => {{
          e.stopPropagation();
          if (gruposOcultos.has(g)) gruposOcultos.delete(g); else gruposOcultos.add(g);
          montarTabela();
        }};
      }}
    }}

    if (btnToggle) tdItem.appendChild(btnToggle);
    const spanNome = document.createElement("span");
    spanNome.textContent = nome;
    tdItem.appendChild(spanNome);
    tr.appendChild(tdItem);

    let totalGeral = 0;
    colunas.forEach(col => {{
      const acc = somarPeriodo(cidadesSel, agrupador(col));
      const valor = obterValor(acc, chave);
      if (chave !== "ticket_medio" && chave !== "pct_arrec_fat") {{
        totalGeral += (typeof valor === "number" ? valor : 0);
      }}
      const td = document.createElement("td");
      td.textContent = fmt(valor, tipo);
      if (valor < 0) td.classList.add("negativo");
      tr.appendChild(td);
    }});

    const tdTotal = document.createElement("td");
    tdTotal.className = "col-total";
    if (chave === "ticket_medio") {{
      const todosMeses = colunas.flatMap(c => agrupador(c));
      const accTotal = somarPeriodo(cidadesSel, todosMeses);
      tdTotal.textContent = fmt(accTotal._arr.ticket_medio, tipo);
    }} else if (chave === "pct_arrec_fat") {{
      const todosMeses = colunas.flatMap(c => agrupador(c));
      const accTotal = somarPeriodo(cidadesSel, todosMeses);
      tdTotal.textContent = fmt(obterValor(accTotal, "pct_arrec_fat"), tipo);
    }} else {{
      tdTotal.textContent = fmt(totalGeral, tipo === "moeda4" ? "moeda" : tipo);
      if (totalGeral < 0) tdTotal.classList.add("negativo");
    }}
    tr.appendChild(tdTotal);

    corpo.appendChild(tr);
  }});

  document.getElementById("resumo").textContent =
    `Cidades selecionadas: ${{cidadesSel.length}} de ${{cidades.length}} | ` +
    `Modo: ${{modoVisualizacao === "mensal" ? "Mensal" : "Anual"}} | ` +
    `Período: ${{mesesTodos[idxInicio]}} até ${{mesesTodos[idxFim]}} | ` +
    `Colunas exibidas: ${{colunas.length}} | ` +
    `Visão Auditoria: ${{modoAuditoria ? "Ativada (Ativa Faturando / Cortada)" : "Desativada"}}`;
}}

const CORES = ["#1F4E78", "#c0392b", "#2f9e5c", "#8e44ad", "#16a085"];
const COR_ESGOTO = "#e67e22";

function montarGraficos() {{
  chartsInstances.forEach(ch => ch.destroy());
  chartsInstances = [];

  const cidadesSel = getSelecionadas();
  const {{ colunas, agrupador }} = getColunas();

  const grid = document.getElementById("grid");
  grid.innerHTML = "";

  GRAFICOS.forEach((graf, idx) => {{
    const [titulo, series, tipo] = graf;

    const nota = NOTAS_GRAFICOS[titulo];
    const notaVisivel = nota && colunas.includes(nota.mes);
    const notaHtml = notaVisivel
      ? `<div class="nota-grafico">⚠️ ${{nota.texto}} (${{nota.mes}})</div>`
      : "";

    const card = document.createElement("div");
    card.className = "card";
    card.innerHTML = `<h3>${{titulo}}</h3><canvas id="chart_${{idx}}"></canvas>${{notaHtml}}`;
    grid.appendChild(card);

    const datasets = [];

    series.forEach((s, i) => {{
      const [nomeSerie, chave] = s;
      const valores = colunas.map(col => {{
        const acc = somarPeriodo(cidadesSel, agrupador(col));
        return obterValor(acc, chave);
      }});
      const cor = nomeSerie.toLowerCase().includes("esgoto") ? COR_ESGOTO : CORES[i % CORES.length];

      datasets.push({{
        label: nomeSerie,
        data: valores,
        borderColor: cor,
        backgroundColor: cor,
        tension: 0.25,
        fill: false,
        pointRadius: 2,
      }});

      const soma = valores.reduce((a, b) => a + b, 0);
      const media = valores.length ? soma / valores.length : 0;
      datasets.push({{
        label: `Média ${{nomeSerie}}`,
        data: colunas.map(() => media),
        borderColor: "#888888",
        borderDash: [6, 4],
        borderWidth: 1.5,
        pointRadius: 0,
        fill: false,
        tension: 0,
      }});
    }});

    const ctx = document.getElementById(`chart_${{idx}}`).getContext("2d");
    const chart = new Chart(ctx, {{
      type: "line",
      data: {{ labels: colunas, datasets: datasets }},
      options: {{
        responsive: true,
        interaction: {{ mode: "index", intersect: false }},
        plugins: {{ legend: {{ position: "bottom", labels: {{ font: {{ size: 10 }} }} }} }},
        scales: {{
          x: {{ ticks: {{ font: {{ size: 10 }} }}, grid: {{ display: false }} }},
          y: {{ ticks: {{ display: false }}, grid: {{ display: false }} }},
        }},
      }}
    }});
    chartsInstances.push(chart);
  }});
}}

function atualizarTudo() {{
  montarTabela();
  if (abaAtiva === "graficos") montarGraficos();
}}

function trocarAba(aba) {{
  if (aba === abaAtiva) return;
  abaAtiva = aba;

  const painelTabela = document.getElementById("painelTabela");
  const painelGraficos = document.getElementById("painelGraficos");

  if (aba === "tabela") {{
    painelGraficos.classList.remove("ativo");
    setTimeout(() => {{ painelTabela.classList.add("ativo"); }}, 60);
  }} else {{
    painelTabela.classList.remove("ativo");
    setTimeout(() => {{
      painelGraficos.classList.add("ativo");
      montarGraficos();
    }}, 60);
  }}

  document.getElementById("btnAbaTabela").classList.toggle("ativa", aba === "tabela");
  document.getElementById("btnAbaGraficos").classList.toggle("ativa", aba === "graficos");
}}

document.addEventListener("click", (e) => {{
  document.querySelectorAll(".dropdown").forEach(dd => {{ if (!dd.contains(e.target)) dd.classList.remove("aberto"); }});
}});

document.addEventListener("DOMContentLoaded", () => {{
  montarFiltros();
  montarTabela();

  document.getElementById("btnAbaTabela").onclick = () => trocarAba("tabela");
  document.getElementById("btnAbaGraficos").onclick = () => trocarAba("graficos");

  document.getElementById("btnDropCidades").onclick = (e) => {{
    e.stopPropagation();
    document.getElementById("btnDropCidades").closest(".dropdown").classList.toggle("aberto");
  }};
  document.getElementById("btnDropModo").onclick = (e) => {{
    e.stopPropagation();
    document.getElementById("btnDropModo").closest(".dropdown").classList.toggle("aberto");
  }};

  document.getElementById("chkTodasCidades").addEventListener("change", (e) => {{
    document.querySelectorAll(".chkCidade").forEach(c => c.checked = e.target.checked);
    salvarFiltros(); atualizarTudo();
  }});

  document.getElementById("listaCidades").addEventListener("change", (e) => {{
    if (e.target.classList.contains("chkCidade")) {{ salvarFiltros(); atualizarTudo(); }}
  }});

  document.querySelectorAll(".opcao-modo").forEach(op => {{
    op.onclick = () => {{
      modoVisualizacao = op.dataset.modo;
      document.getElementById("btnDropModo").textContent =
        "Modo: " + (modoVisualizacao === "mensal" ? "Mensal ▾" : "Anual ▾");
      document.getElementById("btnDropModo").closest(".dropdown").classList.remove("aberto");
      atualizarTudo();
    }};
  }});
}});
</script>
</body>
</html>
"""
    os.makedirs(os.path.dirname(caminho_saida), exist_ok=True)
    with open(caminho_saida, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"HTML único gerado: {caminho_saida}")


# =========================================================
# EXECUÇÃO
# =========================================================
if __name__ == "__main__":
    df_fat = carregar_faturamento(ARQUIVO_FATURAMENTO, SHEET_FATURAMENTO)
    df_arr = carregar_arrecadacao(ARQUIVO_ARRECADACAO)
    desconto_map = carregar_desconto(ARQUIVO_DESCONTO)

    dataset, meses, cidades = montar_dataset_unificado(df_fat, df_arr, desconto_map)

    descontos_simplificado = {m: {"valor": desconto_map[m]["valor"]} for m in desconto_map}

    gerar_html_unico(dataset, meses, cidades, descontos_simplificado, ARQUIVO_SAIDA, ARQUIVO_INDEX)
