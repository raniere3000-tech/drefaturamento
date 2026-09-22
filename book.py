import pandas as pd
import json
from datetime import datetime, timedelta

ARQUIVO_ENTRADA = "/workspaces/drefaturamento/comercial.xlsx"
ARQUIVO_SAIDA = "/workspaces/drefaturamento/docsDRE_Faturamento.html"
ARQUIVO_SAIDA_GRAFICOS = "/workspaces/drefaturamento/docsDRE_Graficos.html"
SHEET_NAME = "Export"

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

COLS_NUM = [
    "Faturamento bruto direta agua", "Faturamento bruto direta esgoto",
    "Faturamento bruto indireta agua", "Faturamento bruto indireta esgoto",
    "Faturamento bruto indiretas total", "Qtd eco fat agua", "Qtd eco fat esgoto",
    "Vol faturado agua direto M³", "Vol faturado esgoto direto M³",
    "Volume medido de agua m³", "R$ Cancelamento total", "R$ Faturamento total liquido",
]

MESES_PT = ["Jan", "Fev", "Mar", "Abr", "Mai", "Jun",
            "Jul", "Ago", "Set", "Out", "Nov", "Dez"]


def converter_referencia(valor):
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
    if pd.notna(dt):
        return dt
    return None


def carregar_dados(caminho, sheet):
    df = pd.read_excel(caminho, sheet_name=sheet)

    df = df[~df["CIDADE"].astype(str).str.contains("Filtros aplicados", na=False)]
    df = df[df["CIDADE"].astype(str).str.upper() != "TOTAL"]

    for col in COLS_NUM:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    df["DSC_CLASSE"] = df["DSC_CLASSE"].astype(str).str.upper().str.strip()
    df["CIDADE"] = df["CIDADE"].astype(str).str.upper().str.strip()

    ref_cols = [c for c in df.columns if str(c).startswith("Referência")]

    melhor_col = None
    melhor_qtd_validas = -1
    melhor_serie = None

    for c in ref_cols:
        serie_convertida = df[c].apply(converter_referencia)
        qtd_validas = serie_convertida.notna().sum()
        if qtd_validas > melhor_qtd_validas:
            melhor_qtd_validas = qtd_validas
            melhor_col = c
            melhor_serie = serie_convertida

    if melhor_serie is None:
        melhor_serie = pd.Series([None] * len(df), index=df.index)

    df["_DATA_REF"] = pd.to_datetime(melhor_serie, errors="coerce")

    total_antes = len(df)
    df = df[df["_DATA_REF"].notna()].copy()
    total_depois = len(df)
    print(f"[DEBUG] Coluna de referência usada: {melhor_col} | "
          f"Linhas removidas por 'Sem Data': {total_antes - total_depois} | "
          f"Linhas restantes: {total_depois}")

    df["_MES_ANO"] = df["_DATA_REF"].apply(
        lambda d: f"{MESES_PT[d.month - 1]}/{d.year}"
    )
    df["_ANO"] = df["_DATA_REF"].apply(lambda d: str(d.year))
    df["_MES_ORDEM"] = df["_DATA_REF"].apply(lambda d: d.year * 100 + d.month)

    return df


def calcular_dre(df_sub):
    r = {}
    r["direta_agua"] = df_sub["Faturamento bruto direta agua"].sum()
    r["direta_esgoto"] = df_sub["Faturamento bruto direta esgoto"].sum()
    r["diretas_totais"] = r["direta_agua"] + r["direta_esgoto"]

    frentes_agua = {}
    for classe, nome in MAP_FRENTE_AGUA.items():
        frentes_agua[nome] = df_sub.loc[
            df_sub["DSC_CLASSE"] == classe, "Faturamento bruto indireta agua"
        ].sum()

    frentes_esgoto = {}
    for classe, nome in MAP_FRENTE_ESGOTO.items():
        frentes_esgoto[nome] = df_sub.loc[
            df_sub["DSC_CLASSE"] == classe, "Faturamento bruto indireta esgoto"
        ].sum()

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
    r["vol_medio_esgoto"] = 0  # coluna não existe no dataset original ainda

    r["tarifa_media_agua"] = (r["direta_agua"] / r["vol_fat_agua"]) if r["vol_fat_agua"] else 0
    r["tarifa_media_esgoto"] = (r["direta_esgoto"] / r["vol_fat_esgoto"]) if r["vol_fat_esgoto"] else 0

    r["cancelamento"] = df_sub["R$ Cancelamento total"].sum()
    r["faturamento_liquido"] = df_sub["R$ Faturamento total liquido"].sum()
    return r


def montar_dataset(df):
    dataset = {}
    meses_map = {}

    for (cidade, mes), grupo in df.groupby(["CIDADE", "_MES_ANO"]):
        ordem = grupo["_MES_ORDEM"].iloc[0]
        meses_map[mes] = ordem
        dataset.setdefault(cidade, {})[mes] = calcular_dre(grupo)

    meses_ordenados = [m for m, _ in sorted(meses_map.items(), key=lambda x: x[1])]
    cidades = sorted(dataset.keys())
    return dataset, meses_ordenados, cidades


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
    ("VOLUME MÉDIO DE ÁGUA (m³)", "vol_medio_agua", 0, True, False, "num", "extras"),
    ("TARIFA MÉDIA DE ÁGUA (R$/m³)", "tarifa_media_agua", 0, True, False, "moeda4", "extras"),
    ("TARIFA MÉDIA DE ESGOTO (R$/m³)", "tarifa_media_esgoto", 0, True, False, "moeda4", "extras"),
    ("CANCELAMENTO", "cancelamento", 0, True, False, "moeda", None),
    ("FATURAMENTO LÍQUIDO C/ REAJUSTE", "faturamento_liquido", 0, True, False, "moeda", None),
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
# DEFINIÇÃO DOS GRÁFICOS (página de gráficos)
# Cada item: (Título do gráfico, [ (nome_legenda, chave_no_dre), ... ], tipo)
# =========================================================
GRAFICOS = [
    ("DIRETAS TOTAIS", [
        ("Diretas Água", "direta_agua"),
        ("Diretas Esgoto", "direta_esgoto"),
    ], "moeda"),
    ("INDIRETA ÁGUA", [
        ("Corte de Água", "F_Corte de Água"),
        ("Religações", "F_Religações"),
        ("Ligações de Água", "F_Ligações de Água"),
        ("Sanções", "F_Sanções"),
        ("Outros Água", "F_Outros Água"),
    ], "moeda"),
    ("FAT. ECONOMIAS ÁGUA E ESGOTO", [
        ("Fat. Economias Água", "eco_agua"),
        ("Fat. Economias Esgoto", "eco_esgoto"),
    ], "num"),
    ("VOL. FAT. ÁGUA E ESGOTO", [
        ("Vol. Fat. Água", "vol_fat_agua"),
        ("Vol. Fat. Esgoto", "vol_fat_esgoto"),
    ], "num"),
    ("VOLUME MÉDIO DE ÁGUA", [
        ("Volume Médio de Água", "vol_medio_agua"),
    ], "num"),
    ("TARIFA MÉDIA DE ÁGUA E ESGOTO", [
        ("Tarifa Média de Água", "tarifa_media_agua"),
        ("Tarifa Média de Esgoto", "tarifa_media_esgoto"),
    ], "moeda4"),
    ("CANCELAMENTO", [("Cancelamento", "cancelamento")], "moeda"),
    ("FATURAMENTO LÍQUIDO C/ REAJUSTE", [("Faturamento Líquido c/ Reajuste", "faturamento_liquido")], "moeda"),
]


def gerar_html(dataset, meses, cidades, caminho_saida, arquivo_graficos):
    dataset_json = json.dumps(dataset, ensure_ascii=False)
    meses_json = json.dumps(meses, ensure_ascii=False)
    cidades_json = json.dumps(cidades, ensure_ascii=False)
    linhas_json = json.dumps(LINHAS_DRE, ensure_ascii=False)
    grupos_ocultos_json = json.dumps(GRUPOS_OCULTOS_PADRAO, ensure_ascii=False)
    grupos_toggle_json = json.dumps(GRUPOS_TOGGLE, ensure_ascii=False)
    grupo_combo_json = json.dumps(GRUPO_COMBO, ensure_ascii=False)
    nome_arquivo_graficos = arquivo_graficos.split("/")[-1]

    html = f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<title>DRE de Faturamento</title>
<style>
  body {{ font-family: Arial, sans-serif; margin: 20px; background:#f5f6f8; }}
  h1 {{ font-size: 20px; color:#1F4E78; margin-bottom:12px; }}
  #filtros {{ background:#fff; padding:12px 15px; border-radius:8px; margin-bottom:14px;
              box-shadow:0 1px 3px rgba(0,0,0,.1); display:flex; align-items:center; flex-wrap:wrap; gap:0; }}
  .btn {{ padding:8px 14px; border:none; border-radius:5px; background:#1F4E78; color:#fff;
          cursor:pointer; font-size:12.5px; text-decoration:none; display:inline-block; }}
  .btn.secundario {{ background:#5a7a99; }}
  .btn.grafico {{ background:#2f9e5c; margin-left:auto; }}
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
  td.col-total {{ background:#dbe6f1; font-weight:bold; }}

  .toggle-btn {{
    cursor:pointer; color:#1F4E78; font-weight:bold; margin-right:6px;
    display:inline-block; width:14px; text-align:center;
  }}
  #resumo {{ margin-top:8px; font-size:11.5px; color:#666; }}
</style>
</head>
<body>

<h1>DRE de Faturamento — Filtro de Cidades, Anos e Meses</h1>

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

  <a class="btn grafico" href="{nome_arquivo_graficos}">📊 Ver Gráficos</a>
</div>

<table id="tabelaDRE">
  <thead><tr id="headerRow"><th style="text-align:left">Item</th></tr></thead>
  <tbody id="corpoTabela"></tbody>
</table>

<div id="resumo"></div>

<script>
const dataset = {dataset_json};
const mesesTodos = {meses_json};
const cidades = {cidades_json};
const linhas = {linhas_json};
const GRUPOS_OCULTOS_PADRAO = {grupos_ocultos_json};
const GRUPOS_TOGGLE = {grupos_toggle_json};
const GRUPO_COMBO = {grupo_combo_json};

let modoVisualizacao = "mensal";
let gruposOcultos = new Set(GRUPOS_OCULTOS_PADRAO);
let idxInicio = 0;
let idxFim = mesesTodos.length - 1;

function fmt(valor, tipo) {{
  if (valor === undefined || valor === null || isNaN(valor)) return "-";
  if (tipo === "moeda") {{
    return valor.toLocaleString('pt-BR', {{minimumFractionDigits:2, maximumFractionDigits:2}});
  }}
  if (tipo === "moeda4") {{
    return valor.toLocaleString('pt-BR', {{minimumFractionDigits:2, maximumFractionDigits:4}});
  }}
  return valor.toLocaleString('pt-BR', {{maximumFractionDigits:0}});
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
  selIni.value = 0;
  selFim.value = mesesTodos.length - 1;

  selIni.onchange = () => {{
    idxInicio = parseInt(selIni.value);
    if (idxInicio > idxFim) {{ idxFim = idxInicio; selFim.value = idxFim; }}
    montarTabela();
  }};
  selFim.onchange = () => {{
    idxFim = parseInt(selFim.value);
    if (idxFim < idxInicio) {{ idxInicio = idxFim; selIni.value = idxInicio; }}
    montarTabela();
  }};
}}

function getSelecionadas() {{
  return Array.from(document.querySelectorAll(".chkCidade:checked")).map(el => el.value);
}}

function somarDRE(cidadesSel, mes) {{
  const acc = {{}};
  cidadesSel.forEach(c => {{
    const dre = dataset[c] && dataset[c][mes];
    if (!dre) return;
    for (const key in dre) {{
      if (key === "frentes_agua" || key === "frentes_esgoto") {{
        for (const f in dre[key]) {{
          const k = (key === "frentes_agua" ? "F_" : "E_") + f;
          acc[k] = (acc[k] || 0) + dre[key][f];
        }}
      }} else {{
        acc[key] = (acc[key] || 0) + dre[key];
      }}
    }}
  }});
  return acc;
}}

function anoDoMes(mes) {{
  return mes.split("/")[1];
}}

function getMesesNoIntervalo() {{
  return mesesTodos.slice(idxInicio, idxFim + 1);
}}

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

function somarDREVarios(cidadesSel, listaMeses) {{
  const acc = {{}};
  listaMeses.forEach(mes => {{
    const parcial = somarDRE(cidadesSel, mes);
    for (const k in parcial) acc[k] = (acc[k] || 0) + parcial[k];
  }});
  return acc;
}}

function grupoVisivel(grupo) {{
  if (grupo === null) return true;
  const partes = grupo.split("|");
  return !partes.some(g => gruposOcultos.has(g));
}}

function comboOculto(gruposLista) {{
  return gruposLista.every(g => gruposOcultos.has(g));
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
    tr.className = (isTotal ? "total " : "") + (isSub ? "subtotal " : "") +
                   (nivel === 1 ? "nivel1 " : "");

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
        if (comboOculto(gruposLista)) {{
          gruposLista.forEach(g => gruposOcultos.delete(g));
        }} else {{
          gruposLista.forEach(g => gruposOcultos.add(g));
        }}
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
      const acc = somarDREVarios(cidadesSel, agrupador(col));
      const valor = acc[chave] || 0;
      totalGeral += (typeof valor === "number") ? valor : 0;
      const td = document.createElement("td");
      td.textContent = fmt(valor, tipo);
      tr.appendChild(td);
    }});

    const tdTotal = document.createElement("td");
    tdTotal.className = "col-total";
    tdTotal.textContent = fmt(totalGeral, tipo === "moeda4" ? "moeda" : tipo);
    tr.appendChild(tdTotal);

    corpo.appendChild(tr);
  }});

  document.getElementById("resumo").textContent =
    `Cidades selecionadas: ${{cidadesSel.length}} de ${{cidades.length}} | ` +
    `Modo: ${{modoVisualizacao === "mensal" ? "Mensal" : "Anual"}} | ` +
    `Período: ${{mesesTodos[idxInicio]}} até ${{mesesTodos[idxFim]}} | ` +
    `Colunas exibidas: ${{colunas.length}}`;
}}

document.addEventListener("click", (e) => {{
  document.querySelectorAll(".dropdown").forEach(dd => {{
    if (!dd.contains(e.target)) dd.classList.remove("aberto");
  }});
}});

document.addEventListener("DOMContentLoaded", () => {{
  montarFiltros();
  montarTabela();

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
    montarTabela();
  }});

  document.getElementById("listaCidades").addEventListener("change", (e) => {{
    if (e.target.classList.contains("chkCidade")) montarTabela();
  }});

  document.querySelectorAll(".opcao-modo").forEach(op => {{
    op.onclick = () => {{
      modoVisualizacao = op.dataset.modo;
      document.getElementById("btnDropModo").textContent =
        "Modo: " + (modoVisualizacao === "mensal" ? "Mensal ▾" : "Anual ▾");
      document.getElementById("btnDropModo").closest(".dropdown").classList.remove("aberto");
      montarTabela();
    }};
  }});
}});
</script>
</body>
</html>
"""
    with open(caminho_saida, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"HTML gerado: {caminho_saida}")


def gerar_html_graficos(dataset, meses, cidades, caminho_saida, arquivo_principal):
    dataset_json = json.dumps(dataset, ensure_ascii=False)
    meses_json = json.dumps(meses, ensure_ascii=False)
    cidades_json = json.dumps(cidades, ensure_ascii=False)
    graficos_json = json.dumps(GRAFICOS, ensure_ascii=False)
    nome_arquivo_principal = arquivo_principal.split("/")[-1]

    html = f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<title>Gráficos - DRE de Faturamento</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4"></script>
<style>
  body {{ font-family: Arial, sans-serif; margin: 20px; background:#f5f6f8; }}
  h1 {{ font-size: 20px; color:#1F4E78; margin-bottom:12px; }}
  #filtros {{ background:#fff; padding:12px 15px; border-radius:8px; margin-bottom:14px;
              box-shadow:0 1px 3px rgba(0,0,0,.1); display:flex; align-items:center; flex-wrap:wrap; gap:0; }}
  .btn {{ padding:8px 14px; border:none; border-radius:5px; background:#1F4E78; color:#fff;
          cursor:pointer; font-size:12.5px; text-decoration:none; display:inline-block; }}
  .btn.voltar {{ background:#5a7a99; margin-left:auto; }}
  .dropdown {{ position: relative; display: inline-block; margin-right: 10px; }}
  .dropdown-content {{
    display: none; position: absolute; background: #fff; min-width: 230px;
    box-shadow: 0 4px 10px rgba(0,0,0,.15); border-radius: 6px; z-index: 10;
    padding: 8px 0; max-height: 280px; overflow-y: auto; margin-top:4px;
  }}
  .dropdown-content label {{
    display: block; padding: 6px 14px; font-size: 12.5px; cursor: pointer; white-space: nowrap;
  }}
  .dropdown-content hr {{ margin:6px 0; border:none; border-top:1px solid #eee; }}
  .dropdown.aberto .dropdown-content {{ display: block; }}
  .dropbtn {{ min-width: 150px; text-align: left; }}
  .intervalo-box {{ display:flex; align-items:center; gap:6px; padding:0 10px; font-size:12.5px; color:#1F4E78; }}
  .intervalo-box select {{ font-size:12.5px; padding:5px 8px; border-radius:5px; border:1px solid #ccc; }}

  #grid {{ display:grid; grid-template-columns: repeat(2, 1fr); gap:16px; }}
  .card {{ background:#fff; border-radius:8px; box-shadow:0 1px 3px rgba(0,0,0,.1); padding:12px; }}
  .card h3 {{ margin:0 0 8px 0; font-size:14px; color:#1F4E78; text-align:center; }}
  .card canvas {{ max-height:280px; }}
  @media (max-width: 900px) {{ #grid {{ grid-template-columns: 1fr; }} }}
</style>
</head>
<body>

<h1>Gráficos — DRE de Faturamento</h1>

<div id="filtros">
  <div class="dropdown">
    <button class="btn dropbtn" id="btnDropCidades">Cidades ▾</button>
    <div class="dropdown-content" id="listaCidades">
      <label><input type="checkbox" id="chkTodasCidades" checked> <b>Selecionar Todas</b></label>
      <hr>
    </div>
  </div>

  <div class="intervalo-box">
    <b>Período:</b>
    De <select id="selMesInicio"></select>
    até <select id="selMesFim"></select>
  </div>

  <a class="btn voltar" href="{nome_arquivo_principal}">⬅ Voltar à Tabela</a>
</div>

<div id="grid"></div>

<script>
const dataset = {dataset_json};
const mesesTodos = {meses_json};
const cidades = {cidades_json};
const GRAFICOS = {graficos_json};

let idxInicio = 0;
let idxFim = mesesTodos.length - 1;
let chartsInstances = [];

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
  selIni.value = 0;
  selFim.value = mesesTodos.length - 1;

  selIni.onchange = () => {{
    idxInicio = parseInt(selIni.value);
    if (idxInicio > idxFim) {{ idxFim = idxInicio; selFim.value = idxFim; }}
    montarGraficos();
  }};
  selFim.onchange = () => {{
    idxFim = parseInt(selFim.value);
    if (idxFim < idxInicio) {{ idxInicio = idxFim; selIni.value = idxInicio; }}
    montarGraficos();
  }};
}}

function getSelecionadas() {{
  return Array.from(document.querySelectorAll(".chkCidade:checked")).map(el => el.value);
}}

function somarDRE(cidadesSel, mes) {{
  const acc = {{}};
  cidadesSel.forEach(c => {{
    const dre = dataset[c] && dataset[c][mes];
    if (!dre) return;
    for (const key in dre) {{
      if (key === "frentes_agua" || key === "frentes_esgoto") {{
        for (const f in dre[key]) {{
          const k = (key === "frentes_agua" ? "F_" : "E_") + f;
          acc[k] = (acc[k] || 0) + dre[key][f];
        }}
      }} else {{
        acc[key] = (acc[key] || 0) + dre[key];
      }}
    }}
  }});
  return acc;
}}

// Paleta padrão + cor fixa laranja para qualquer série de Esgoto
const CORES = ["#1F4E78", "#c0392b", "#2f9e5c", "#8e44ad", "#16a085"];
const COR_ESGOTO = "#e67e22";

function montarGraficos() {{
  chartsInstances.forEach(ch => ch.destroy());
  chartsInstances = [];

  const cidadesSel = getSelecionadas();
  const mesesFiltrados = mesesTodos.slice(idxInicio, idxFim + 1);

  const grid = document.getElementById("grid");
  grid.innerHTML = "";

  GRAFICOS.forEach((graf, idx) => {{
    const [titulo, series, tipo] = graf;

    const card = document.createElement("div");
    card.className = "card";
    card.innerHTML = `<h3>${{titulo}}</h3><canvas id="chart_${{idx}}"></canvas>`;
    grid.appendChild(card);

    const datasets = series.map((s, i) => {{
      const [nomeSerie, chave] = s;
      const valores = mesesFiltrados.map(mes => {{
        const acc = somarDRE(cidadesSel, mes);
        return acc[chave] || 0;
      }});
      // Se a série for referente a Esgoto, força a cor laranja
      const cor = nomeSerie.toLowerCase().includes("esgoto") ? COR_ESGOTO : CORES[i % CORES.length];
      return {{
        label: nomeSerie,
        data: valores,
        borderColor: cor,
        backgroundColor: cor,
        tension: 0.25,
        fill: false,
        pointRadius: 2,
      }};
    }});

    const ctx = document.getElementById(`chart_${{idx}}`).getContext("2d");
    const chart = new Chart(ctx, {{
      type: "line",
      data: {{ labels: mesesFiltrados, datasets: datasets }},
      options: {{
        responsive: true,
        interaction: {{ mode: "index", intersect: false }},
        plugins: {{ legend: {{ position: "bottom", labels: {{ font: {{ size: 10 }} }} }} }},
        scales: {{
          // Eixo X: mantém rótulos dos meses, sem linhas de grade
          x: {{
            ticks: {{ font: {{ size: 10 }} }},
            grid: {{ display: false }}
          }},
          // Eixo Y: sem grade e sem valores exibidos
          y: {{
            ticks: {{ display: false }},
            grid: {{ display: false }}
          }}
        }}
      }}
    }});
    chartsInstances.push(chart);
  }});
}}

document.addEventListener("click", (e) => {{
  document.querySelectorAll(".dropdown").forEach(dd => {{
    if (!dd.contains(e.target)) dd.classList.remove("aberto");
  }});
}});

document.addEventListener("DOMContentLoaded", () => {{
  montarFiltros();
  montarGraficos();

  document.getElementById("btnDropCidades").onclick = (e) => {{
    e.stopPropagation();
    document.getElementById("btnDropCidades").closest(".dropdown").classList.toggle("aberto");
  }};

  document.getElementById("chkTodasCidades").addEventListener("change", (e) => {{
    document.querySelectorAll(".chkCidade").forEach(c => c.checked = e.target.checked);
    montarGraficos();
  }});

  document.getElementById("listaCidades").addEventListener("change", (e) => {{
    if (e.target.classList.contains("chkCidade")) montarGraficos();
  }});
}});
</script>
</body>
</html>
"""
    with open(caminho_saida, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"HTML de gráficos gerado: {caminho_saida}")


if __name__ == "__main__":
    df = carregar_dados(ARQUIVO_ENTRADA, SHEET_NAME)
    dataset, meses, cidades = montar_dataset(df)
    gerar_html(dataset, meses, cidades, ARQUIVO_SAIDA, ARQUIVO_SAIDA_GRAFICOS)
    gerar_html_graficos(dataset, meses, cidades, ARQUIVO_SAIDA_GRAFICOS, ARQUIVO_SAIDA)
