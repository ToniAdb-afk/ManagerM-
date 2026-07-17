# -*- coding: utf-8 -*-
"""
build_disparos_alarme.py
-------------------------------------------------------------------
Le a planilha de disparos de alarme (export do Segware Sigma Cloud,
que vem com milhares de colunas extras vazias/repetidas) e gera o
dashboard 'disparos_alarme.html' com filtros e drill-down.

Uso:
    python3 build_disparos_alarme.py caminho/para/BASE_DE_DISPARO.xlsx
"""

import sys
import re
import csv
import json
import time
import html as html_lib
from datetime import datetime

# =====================================================================
# 1) EXTRACAO RAPIDA DO XLSX (lida com o export bloatado de +2000 colunas)
# =====================================================================

def extrair_xlsx_bloatado(caminho_xlsx, pasta_tmp, max_col_util=16):
    """
    Descompacta o .xlsx e le apenas as colunas uteis (A ate max_col_util)
    direto do XML, ignorando as milhares de colunas vazias/repetidas que
    esse tipo de export do Segware costuma trazer. Evita carregar o
    arquivo inteiro na memoria (pode ter varios GB de XML).
    Retorna lista de linhas (cada linha = lista de strings).
    """
    import zipfile
    import os

    os.makedirs(pasta_tmp, exist_ok=True)
    with zipfile.ZipFile(caminho_xlsx) as z:
        z.extract("xl/worksheets/sheet1.xml", pasta_tmp)
        z.extract("xl/sharedStrings.xml", pasta_tmp)

    # shared strings (pequeno, le rapido)
    from lxml import etree
    NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
    strings = []
    for _, elem in etree.iterparse(os.path.join(pasta_tmp, "xl/sharedStrings.xml"), tag=NS + "si"):
        strings.append("".join(elem.itertext()))
        elem.clear()

    cell_re = re.compile(
        r'<c r="([A-P])(\d+)"(?: [^>]*)?(?:/>|>(?:<v>([^<]*)</v>)?</c>)'
    )

    linhas = []
    buf = ""
    CHUNK = 8 * 1024 * 1024
    sheet_path = os.path.join(pasta_tmp, "xl/worksheets/sheet1.xml")
    with open(sheet_path, "r", encoding="utf-8") as f:
        while True:
            chunk = f.read(CHUNK)
            if not chunk:
                break
            buf += chunk
            while True:
                start_idx = buf.find('<row r="')
                if start_idx == -1:
                    if len(buf) > 2_000_000:
                        buf = buf[-2000:]
                    break
                end_idx = buf.find("</row>", start_idx)
                if end_idx == -1:
                    buf = buf[start_idx:]
                    break
                row_xml = buf[start_idx:end_idx]
                q_idx = row_xml.find('<c r="Q')
                prefix = row_xml if q_idx == -1 else row_xml[:q_idx]

                row_vals = [""] * max_col_util
                for m in cell_re.finditer(prefix):
                    col_letter, _row_num, val = m.group(1), m.group(2), m.group(3)
                    col_num = ord(col_letter) - ord("A")
                    if val is None:
                        continue
                    try:
                        idx = int(val)
                        text = strings[idx] if idx < len(strings) else val
                    except ValueError:
                        text = val
                    row_vals[col_num] = text
                linhas.append(row_vals)
                buf = buf[end_idx + len("</row>"):]

    return linhas


# =====================================================================
# 2) PARSING DOS CAMPOS
# =====================================================================

COLS = ["Empresa", "Conta", "Dt_Inicio", "ColD_vazia", "Dt_Atendimento",
        "Dt_Fim", "Tempo_Total", "Status", "Tipo", "Tempo_Deslocamento",
        "Descricao", "Operador_Fechamento", "Codigo_Evento", "Particao",
        "Auxiliar", "Descricao2"]


def parse_conta(conta):
    """Extrai codigo da unidade, cliente (BMB/MERCANTIL/...), UF e cidade
    a partir do texto livre da coluna 'Conta'."""
    s = (conta or "").strip()
    cliente_tag = ""
    m = re.search(r"\(([^)]+)\)\s*$", s)
    if m:
        cliente_tag = m.group(1).strip()
        s = s[:m.start()].strip()
    uf = ""
    m2 = re.search(r"-\s*([A-Za-zÀ-ÿ]{2})\s*$", s)
    if m2:
        uf = m2.group(1).upper().strip()
        s = s[:m2.start()].strip()
    m3 = re.match(r"^(\d+)", s)
    codigo = m3.group(1) if m3 else ""
    parts = [p.strip() for p in s.split(" - ") if p.strip()]
    cidade_raw = parts[-1] if parts else s
    cidade = re.sub(r"^(PAE?|PA)\s+", "", cidade_raw, flags=re.IGNORECASE).strip()
    return codigo, cliente_tag, (uf or "N/D"), (cidade or "N/D")


def parse_dt(s):
    if not s:
        return None
    try:
        return datetime.strptime(s.strip(), "%d/%m/%Y %H:%M:%S")
    except Exception:
        return None


def hhmmss_to_min(s):
    if not s:
        return None
    m = re.match(r"^(\d+):(\d{2}):(\d{2})$", s.strip())
    if not m:
        return None
    h, mi, se = int(m.group(1)), int(m.group(2)), int(m.group(3))
    return round(h * 60 + mi + se / 60, 1)


def categorizar_motivo(descricao_raw):
    """Padroniza a categoria de motivo:
    - 'Sistema' (evento administrativo generico) vira 'Disparo de Alarme'
    - Variacoes de falha de energia (QUADRO DE ENERGIA, QUEDA DE ENERGIA, etc.)
      sao consolidadas em 'FALHA DE ENERGIA ELETRICA (FAC)'
    - Variacoes de problema de rede/ethernet viram 'FALHA DE ETHERNET (FTH)'
    Isso evita que o mesmo tipo de falha fique espalhado em varias categorias
    quase-identicas e garante destaque visual pros dois tipos que o cliente
    quer acompanhar de perto (energia e ethernet).
    """
    d = (descricao_raw or "").strip()
    if not d:
        return "N/D"
    d_upper = d.upper()
    if d_upper == "SISTEMA":
        return "Disparo de Alarme"
    if "ETHERNET" in d_upper:
        return "FALHA DE ETHERNET (FTH)"
    if "ENERGIA" in d_upper:
        return "FALHA DE ENERGIA ELÉTRICA (FAC)"
    return d


def processar_linhas(linhas):
    registros = []
    for row in linhas[1:]:  # pula cabecalho
        row = row + [""] * (16 - len(row))
        d = dict(zip(COLS, row))
        conta = d["Conta"]
        if not conta or not conta.strip():
            continue
        codigo, cliente_tag, uf, cidade = parse_conta(conta)
        dt_inicio = parse_dt(d["Dt_Inicio"])
        dt_atend = parse_dt(d["Dt_Atendimento"])
        dt_fim = parse_dt(d["Dt_Fim"])
        tempo_total_min = hhmmss_to_min(d["Tempo_Total"])
        tempo_atend_min = None
        if dt_inicio and dt_atend:
            tempo_atend_min = round((dt_atend - dt_inicio).total_seconds() / 60, 1)

        registros.append({
            "dt_inicio": dt_inicio,
            "cliente": cliente_tag or "N/D",
            "uf": uf,
            "cidade": cidade,
            "unidade_codigo": codigo,
            "unidade_raw": conta.strip(),
            "motivo": categorizar_motivo(d["Descricao"]),
            "status": (d["Status"] or "N/D").strip() or "N/D",
            "tempo_total_min": tempo_total_min,
            "tempo_atend_min": tempo_atend_min,
            "operador": (d["Operador_Fechamento"] or "").strip(),
        })
    return registros


# =====================================================================
# 3) AGREGACOES E DATASET COMPACTO (indices em vez de strings repetidas)
# =====================================================================

def montar_dataset(registros):
    cidades = sorted({r["cidade"] for r in registros})
    motivos_contagem = {}
    for r in registros:
        motivos_contagem[r["motivo"]] = motivos_contagem.get(r["motivo"], 0) + 1
    # so guardamos como categoria os motivos mais frequentes; os raros viram "Outros"
    motivos_top = [m for m, _ in sorted(motivos_contagem.items(), key=lambda x: -x[1])[:60]]
    motivos_set = set(motivos_top)
    clientes = sorted({r["cliente"] for r in registros})
    ufs = sorted({r["uf"] for r in registros})
    status_list = sorted({r["status"] for r in registros})
    unidades = sorted({r["unidade_raw"] for r in registros})

    cidade_idx = {c: i for i, c in enumerate(cidades)}
    motivo_idx = {m: i for i, m in enumerate(motivos_top + ["Outros"])}
    cliente_idx = {c: i for i, c in enumerate(clientes)}
    uf_idx = {u: i for i, u in enumerate(ufs)}
    status_idx = {s: i for i, s in enumerate(status_list)}
    unidade_idx = {u: i for i, u in enumerate(unidades)}

    linhas_compactas = []
    for r in registros:
        motivo_final = r["motivo"] if r["motivo"] in motivos_set else "Outros"
        linhas_compactas.append([
            r["dt_inicio"].strftime("%Y-%m-%dT%H:%M:%S") if r["dt_inicio"] else None,
            cliente_idx[r["cliente"]],
            uf_idx[r["uf"]],
            cidade_idx[r["cidade"]],
            unidade_idx[r["unidade_raw"]],
            motivo_idx[motivo_final],
            status_idx[r["status"]],
            r["tempo_total_min"],
            r["tempo_atend_min"],
        ])

    dataset = {
        "cidades": cidades,
        "motivos": motivos_top + ["Outros"],
        "clientes": clientes,
        "ufs": ufs,
        "status": status_list,
        "unidades": unidades,
        "linhas": linhas_compactas,
    }
    return dataset


def montar_kpis_e_agregados(registros):
    total = len(registros)
    unidades_distintas = len({r["unidade_raw"] for r in registros})
    cidades_distintas = len({r["cidade"] for r in registros})
    fechados = sum(1 for r in registros if r["status"] == "Fechado")
    tempos_atend = [r["tempo_atend_min"] for r in registros if r["tempo_atend_min"] is not None and r["tempo_atend_min"] >= 0]
    tempos_total = [r["tempo_total_min"] for r in registros if r["tempo_total_min"] is not None and r["tempo_total_min"] >= 0]

    def mediana(lst):
        if not lst:
            return None
        s = sorted(lst)
        n = len(s)
        return s[n // 2] if n % 2 else (s[n // 2 - 1] + s[n // 2]) / 2

    kpis = {
        "total": total,
        "unidades_distintas": unidades_distintas,
        "cidades_distintas": cidades_distintas,
        "pct_fechados": round(100 * fechados / total, 1) if total else 0,
        "tempo_atend_mediano": round(mediana(tempos_atend), 1) if tempos_atend else None,
        "tempo_total_mediano": round(mediana(tempos_total), 1) if tempos_total else None,
    }

    # volume mensal
    volume_mes = {}
    for r in registros:
        if r["dt_inicio"]:
            chave = r["dt_inicio"].strftime("%Y-%m")
            volume_mes[chave] = volume_mes.get(chave, 0) + 1
    volume_mes_ordenado = [{"mes": k, "total": v} for k, v in sorted(volume_mes.items())]

    # top cidades
    cont_cidade = {}
    for r in registros:
        cont_cidade[r["cidade"]] = cont_cidade.get(r["cidade"], 0) + 1
    top_cidades = [{"cidade": c, "total": t} for c, t in sorted(cont_cidade.items(), key=lambda x: -x[1])[:10]]

    # top motivos - top 15 por contagem, mas garantindo que Falha de Energia (FAC)
    # e Falha de Ethernet (FTH) sempre apareçam, mesmo que nao entrem no top 15
    # natural, ja que o cliente quer visibilidade garantida pra esses dois tipos
    cont_motivo = {}
    for r in registros:
        cont_motivo[r["motivo"]] = cont_motivo.get(r["motivo"], 0) + 1
    ranking_completo = sorted(cont_motivo.items(), key=lambda x: -x[1])
    top_motivos_lista = ranking_completo[:15]
    labels_no_top = {m for m, _ in top_motivos_lista}
    for label_fixo in ("FALHA DE ENERGIA ELÉTRICA (FAC)", "FALHA DE ETHERNET (FTH)"):
        if label_fixo not in labels_no_top and label_fixo in cont_motivo:
            top_motivos_lista.append((label_fixo, cont_motivo[label_fixo]))
    top_motivos = [{"motivo": m, "total": t} for m, t in top_motivos_lista]

    # status
    cont_status = {}
    for r in registros:
        cont_status[r["status"]] = cont_status.get(r["status"], 0) + 1
    status_dist = [{"status": s, "total": t} for s, t in sorted(cont_status.items(), key=lambda x: -x[1])]

    # disparos por horario do dia (0-23h) - mostra se ha concentracao na
    # abertura/fechamento das unidades
    cont_hora = {h: 0 for h in range(24)}
    for r in registros:
        if r["dt_inicio"]:
            cont_hora[r["dt_inicio"].hour] += 1
    por_hora = [{"hora": f"{h:02d}h", "total": cont_hora[h]} for h in range(24)]

    agregados = {
        "volume_mes": volume_mes_ordenado,
        "top_cidades": top_cidades,
        "top_motivos": top_motivos,
        "status_dist": status_dist,
        "por_hora": por_hora,
    }
    return kpis, agregados


# =====================================================================
# 4) GERACAO DO HTML
# =====================================================================

TEMPLATE = r"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Disparos de Alarme · Manager+</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.min.js"></script>
<style>
  * { box-sizing:border-box; margin:0; padding:0; }
  body { font-family:'Inter',-apple-system,sans-serif; background:#080b10; color:#e8edf5; min-height:100vh; }
  a.back { color:#7a8ba8; text-decoration:none; font-size:12px; }
  a.back:hover { color:#16c79a; }
  header { background:#0e1420; border-bottom:1px solid #1e2d4a; padding:18px 28px; display:flex; align-items:center; justify-content:space-between; flex-wrap:wrap; gap:10px; }
  header h1 { font-size:18px; font-weight:800; }
  header h1 span { color:#16c79a; }
  main { padding:24px 28px 60px; max-width:1500px; margin:0 auto; }

  .kpis { display:grid; grid-template-columns:repeat(auto-fit,minmax(160px,1fr)); gap:14px; margin-bottom:22px; }
  .kpi { background:#0e1420; border:1px solid #1e2d4a; border-radius:14px; padding:16px 18px; }
  .kpi .label { font-size:11px; color:#7a8ba8; text-transform:uppercase; letter-spacing:.5px; margin-bottom:6px; }
  .kpi .value { font-size:24px; font-weight:800; color:#16c79a; }
  .kpi .value.small { font-size:19px; }

  .filters { background:#0e1420; border:1px solid #1e2d4a; border-radius:14px; padding:16px 18px; margin-bottom:22px; display:flex; gap:12px; flex-wrap:wrap; align-items:end; }
  .filters .fg { display:flex; flex-direction:column; gap:5px; }
  .filters label { font-size:10px; color:#7a8ba8; text-transform:uppercase; letter-spacing:.5px; }
  .filters select, .filters input { background:#080b10; border:1px solid #1e2d4a; color:#e8edf5; border-radius:8px; padding:8px 10px; font-size:12px; min-width:130px; }
  .filters input[type=text] { min-width:200px; }
  .filters button { background:#16c79a; border:none; color:#06251f; font-weight:700; padding:9px 16px; border-radius:8px; font-size:12px; cursor:pointer; }
  .filters button.secondary { background:transparent; border:1px solid #1e2d4a; color:#e8edf5; }
  .active-filter { display:inline-flex; align-items:center; gap:6px; background:rgba(22,199,154,.12); border:1px solid rgba(22,199,154,.35); color:#16c79a; font-size:11px; padding:4px 10px; border-radius:20px; margin:2px 4px 2px 0; cursor:pointer; }

  .charts { display:grid; grid-template-columns:1.3fr 1fr; gap:16px; margin-bottom:16px; }
  .charts-row2 { display:grid; grid-template-columns:1fr 1fr; gap:16px; margin-bottom:22px; }
  .chart-card { background:#0e1420; border:1px solid #1e2d4a; border-radius:14px; padding:18px; }
  .chart-card h3 { font-size:13px; font-weight:700; margin-bottom:12px; color:#e8edf5; }
  .chart-card .hint { font-size:10px; color:#4a5f7a; margin-top:8px; }
  .chart-wrap { position:relative; height:260px; }

  table { width:100%; border-collapse:collapse; font-size:12px; }
  thead th { text-align:left; padding:10px 12px; color:#7a8ba8; font-size:10px; text-transform:uppercase; letter-spacing:.5px; border-bottom:1px solid #1e2d4a; cursor:pointer; user-select:none; white-space:nowrap; }
  thead th:hover { color:#16c79a; }
  tbody td { padding:9px 12px; border-bottom:1px solid #131b28; }
  tbody tr:hover { background:#0e1420; }
  .table-card { background:#0e1420; border:1px solid #1e2d4a; border-radius:14px; padding:18px; overflow-x:auto; }
  .table-top { display:flex; justify-content:space-between; align-items:center; margin-bottom:14px; flex-wrap:wrap; gap:10px; }
  .table-top .count { font-size:12px; color:#7a8ba8; }
  .pill { display:inline-block; padding:2px 8px; border-radius:10px; font-size:10px; font-weight:700; }
  .pill.fechado { background:rgba(39,174,96,.15); color:#27ae60; }
  .pill.aberto { background:rgba(231,76,60,.15); color:#e74c3c; }
  .pill.obs { background:rgba(230,126,34,.15); color:#e67e22; }
  .pagination { display:flex; gap:6px; justify-content:center; margin-top:16px; flex-wrap:wrap; }
  .pagination button { background:#080b10; border:1px solid #1e2d4a; color:#e8edf5; padding:6px 12px; border-radius:8px; font-size:12px; cursor:pointer; }
  .pagination button.active { background:#16c79a; color:#06251f; border-color:#16c79a; font-weight:700; }
  .pagination button:disabled { opacity:.3; cursor:default; }
</style>
</head>
<body>

<header>
  <div>
    <a class="back" href="manager-plus-index.html">&larr; Voltar ao Hub</a>
    <h1 style="margin-top:4px;">🚨 Disparos de <span>Alarme</span></h1>
  </div>
  <div style="font-size:11px;color:#4a5f7a;">Gerado em __DATA_GERACAO__ · __TOTAL__ registros</div>
</header>

<main>
  <div class="kpis">
    <div class="kpi"><div class="label">Total de Disparos</div><div class="value" id="kpi-total">-</div></div>
    <div class="kpi"><div class="label">Unidades Distintas</div><div class="value" id="kpi-unidades">-</div></div>
    <div class="kpi"><div class="label">Cidades Distintas</div><div class="value" id="kpi-cidades">-</div></div>
    <div class="kpi"><div class="label">% Fechados</div><div class="value" id="kpi-fechados">-</div></div>
    <div class="kpi"><div class="label">Tempo Mediano p/ Atender</div><div class="value small" id="kpi-tempo-atend">-</div></div>
    <div class="kpi"><div class="label">Tempo Mediano Total</div><div class="value small" id="kpi-tempo-total">-</div></div>
  </div>

  <div class="filters">
    <div class="fg"><label>Cliente</label><select id="f-cliente"><option value="">Todos</option></select></div>
    <div class="fg"><label>UF</label><select id="f-uf"><option value="">Todos</option></select></div>
    <div class="fg"><label>Cidade</label><select id="f-cidade"><option value="">Todas</option></select></div>
    <div class="fg"><label>Status</label><select id="f-status"><option value="">Todos</option></select></div>
    <div class="fg"><label>Data de</label><input type="date" id="f-data-de"></div>
    <div class="fg"><label>Data até</label><input type="date" id="f-data-ate"></div>
    <div class="fg"><label>Buscar (unidade / motivo)</label><input type="text" id="f-busca" placeholder="ex: PANICO, CAMPINAS, 0103..."></div>
    <div class="fg"><button onclick="limparFiltros()" class="secondary">Limpar filtros</button></div>
    <div class="fg"><button onclick="exportarCSV()">Exportar CSV</button></div>
  </div>
  <div id="filtros-ativos" style="margin:-10px 0 18px;"></div>

  <div class="charts">
    <div class="chart-card">
      <h3>Volume de Disparos por Mês</h3>
      <div class="chart-wrap"><canvas id="chart-mes"></canvas></div>
    </div>
    <div class="chart-card">
      <h3>Status</h3>
      <div class="chart-wrap"><canvas id="chart-status"></canvas></div>
      <div class="hint">Clique numa fatia para filtrar</div>
    </div>
  </div>
  <div class="charts-row2">
    <div class="chart-card">
      <h3>Top 10 Cidades</h3>
      <div class="chart-wrap"><canvas id="chart-cidades"></canvas></div>
      <div class="hint">Clique numa barra para filtrar por cidade</div>
    </div>
    <div class="chart-card">
      <h3>Top 15 Motivos</h3>
      <div class="chart-wrap"><canvas id="chart-motivos"></canvas></div>
      <div class="hint">Clique numa barra para filtrar por motivo · laranja = Falha de Energia (FAC) · ciano = Falha de Ethernet (FTH)</div>
    </div>
  </div>
  <div class="charts" style="grid-template-columns:1fr;">
    <div class="chart-card">
      <h3>Disparos por Horário do Dia</h3>
      <div class="chart-wrap"><canvas id="chart-hora"></canvas></div>
      <div class="hint">Clique numa barra para filtrar por horário · picos aqui costumam coincidir com abertura/fechamento das unidades</div>
    </div>
  </div>

  <div class="table-card">
    <div class="table-top">
      <div class="count"><span id="count-mostrando">0</span> de <span id="count-total">0</span> registros</div>
    </div>
    <div style="overflow-x:auto;">
    <table>
      <thead>
        <tr>
          <th data-col="dt">Data/Hora ▾</th>
          <th data-col="unidade">Unidade</th>
          <th data-col="cidade">Cidade/UF</th>
          <th data-col="cliente">Cliente</th>
          <th data-col="motivo">Motivo</th>
          <th data-col="status">Status</th>
          <th data-col="tempo_total">Tempo Total</th>
          <th data-col="operador">Operador Fechamento</th>
        </tr>
      </thead>
      <tbody id="tbody"></tbody>
    </table>
    </div>
    <div class="pagination" id="pagination"></div>
  </div>
</main>

<script>
const DATASET = __DATASET_JSON__;
const KPIS = __KPIS_JSON__;
const AGREGADOS = __AGREGADOS_JSON__;

// Reconstroi registros "completos" a partir do formato compacto (indices -> strings)
const REGISTROS = DATASET.linhas.map(l => ({
  dt: l[0] ? new Date(l[0]) : null,
  cliente: DATASET.clientes[l[1]],
  uf: DATASET.ufs[l[2]],
  cidade: DATASET.cidades[l[3]],
  unidade: DATASET.unidades[l[4]],
  motivo: DATASET.motivos[l[5]],
  status: DATASET.status[l[6]],
  tempo_total: l[7],
  tempo_atend: l[8],
}));

// ---- popular selects ----
function popularSelect(id, valores) {
  const sel = document.getElementById(id);
  valores.forEach(v => {
    const opt = document.createElement('option');
    opt.value = v; opt.textContent = v;
    sel.appendChild(opt);
  });
}
popularSelect('f-cliente', DATASET.clientes);
popularSelect('f-uf', DATASET.ufs);
popularSelect('f-cidade', DATASET.cidades);
popularSelect('f-status', DATASET.status);

// ---- KPIs ----
document.getElementById('kpi-total').textContent = KPIS.total.toLocaleString('pt-BR');
document.getElementById('kpi-unidades').textContent = KPIS.unidades_distintas.toLocaleString('pt-BR');
document.getElementById('kpi-cidades').textContent = KPIS.cidades_distintas.toLocaleString('pt-BR');
document.getElementById('kpi-fechados').textContent = KPIS.pct_fechados + '%';
document.getElementById('kpi-tempo-atend').textContent = KPIS.tempo_atend_mediano != null ? KPIS.tempo_atend_mediano + ' min' : '-';
document.getElementById('kpi-tempo-total').textContent = KPIS.tempo_total_mediano != null ? KPIS.tempo_total_mediano + ' min' : '-';

// ---- estado de filtros ----
let filtros = { cliente:'', uf:'', cidade:'', status:'', dataDe:'', dataAte:'', busca:'', motivoDrill:'', horaDrill:null };
let paginaAtual = 1;
const PAGE_SIZE = 30;
let ordenacao = { col:'dt', dir:'desc' };

function aplicarFiltros() {
  return REGISTROS.filter(r => {
    if (filtros.cliente && r.cliente !== filtros.cliente) return false;
    if (filtros.uf && r.uf !== filtros.uf) return false;
    if (filtros.cidade && r.cidade !== filtros.cidade) return false;
    if (filtros.status && r.status !== filtros.status) return false;
    if (filtros.motivoDrill && r.motivo !== filtros.motivoDrill) return false;
    if (filtros.horaDrill !== null && (!r.dt || r.dt.getHours() !== filtros.horaDrill)) return false;
    if (filtros.dataDe && (!r.dt || r.dt < new Date(filtros.dataDe))) return false;
    if (filtros.dataAte && (!r.dt || r.dt > new Date(filtros.dataAte + 'T23:59:59'))) return false;
    if (filtros.busca) {
      const b = filtros.busca.toLowerCase();
      if (!r.unidade.toLowerCase().includes(b) && !r.motivo.toLowerCase().includes(b) && !r.cidade.toLowerCase().includes(b)) return false;
    }
    return true;
  });
}

function ordenar(lista) {
  const { col, dir } = ordenacao;
  const mult = dir === 'asc' ? 1 : -1;
  return lista.slice().sort((a,b) => {
    let va = col === 'dt' ? (a.dt ? a.dt.getTime() : -Infinity) : a[col];
    let vb = col === 'dt' ? (b.dt ? b.dt.getTime() : -Infinity) : b[col];
    if (col === 'tempo_total') { va = va ?? -Infinity; vb = vb ?? -Infinity; }
    if (typeof va === 'string') return va.localeCompare(vb) * mult;
    return (va - vb) * mult;
  });
}

function fmtData(d) {
  if (!d) return '-';
  return d.toLocaleString('pt-BR', {day:'2-digit',month:'2-digit',year:'numeric',hour:'2-digit',minute:'2-digit'});
}
function fmtMin(m) {
  if (m == null) return '-';
  const h = Math.floor(m/60), mi = Math.round(m%60);
  return h > 0 ? `${h}h${String(mi).padStart(2,'0')}` : `${mi}min`;
}
function pillClass(status) {
  if (status === 'Fechado') return 'fechado';
  if (status === 'Observação') return 'obs';
  return 'aberto';
}

function renderizar() {
  const filtrados = ordenar(aplicarFiltros());
  document.getElementById('count-mostrando').textContent = filtrados.length.toLocaleString('pt-BR');
  document.getElementById('count-total').textContent = REGISTROS.length.toLocaleString('pt-BR');

  const totalPaginas = Math.max(1, Math.ceil(filtrados.length / PAGE_SIZE));
  if (paginaAtual > totalPaginas) paginaAtual = totalPaginas;
  const inicio = (paginaAtual - 1) * PAGE_SIZE;
  const pagina = filtrados.slice(inicio, inicio + PAGE_SIZE);

  const tbody = document.getElementById('tbody');
  tbody.innerHTML = pagina.map(r => `
    <tr>
      <td>${fmtData(r.dt)}</td>
      <td>${r.unidade.length > 55 ? r.unidade.slice(0,55)+'…' : r.unidade}</td>
      <td>${r.cidade} / ${r.uf}</td>
      <td>${r.cliente}</td>
      <td>${r.motivo}</td>
      <td><span class="pill ${pillClass(r.status)}">${r.status}</span></td>
      <td>${fmtMin(r.tempo_total)}</td>
      <td>${r.operador || '-'}</td>
    </tr>
  `).join('');

  // paginacao
  const pagDiv = document.getElementById('pagination');
  let btns = '';
  btns += `<button ${paginaAtual===1?'disabled':''} onclick="irPagina(${paginaAtual-1})">‹</button>`;
  const janela = 3;
  for (let p = 1; p <= totalPaginas; p++) {
    if (p === 1 || p === totalPaginas || Math.abs(p - paginaAtual) <= janela) {
      btns += `<button class="${p===paginaAtual?'active':''}" onclick="irPagina(${p})">${p}</button>`;
    } else if (Math.abs(p - paginaAtual) === janela + 1) {
      btns += `<span style="color:#4a5f7a;">…</span>`;
    }
  }
  btns += `<button ${paginaAtual===totalPaginas?'disabled':''} onclick="irPagina(${paginaAtual+1})">›</button>`;
  pagDiv.innerHTML = btns;

  renderizarFiltrosAtivos();
}
function irPagina(p) { paginaAtual = p; renderizar(); window.scrollTo({top: document.querySelector('.table-card').offsetTop - 20, behavior:'smooth'}); }

function renderizarFiltrosAtivos() {
  const div = document.getElementById('filtros-ativos');
  const chips = [];
  if (filtros.cliente) chips.push(['Cliente: '+filtros.cliente, () => { filtros.cliente=''; document.getElementById('f-cliente').value=''; }]);
  if (filtros.uf) chips.push(['UF: '+filtros.uf, () => { filtros.uf=''; document.getElementById('f-uf').value=''; }]);
  if (filtros.cidade) chips.push(['Cidade: '+filtros.cidade, () => { filtros.cidade=''; document.getElementById('f-cidade').value=''; }]);
  if (filtros.status) chips.push(['Status: '+filtros.status, () => { filtros.status=''; document.getElementById('f-status').value=''; }]);
  if (filtros.motivoDrill) chips.push(['Motivo: '+filtros.motivoDrill, () => { filtros.motivoDrill=''; }]);
  if (filtros.horaDrill !== null) chips.push(['Horário: '+String(filtros.horaDrill).padStart(2,'0')+'h', () => { filtros.horaDrill=null; }]);
  if (filtros.dataDe) chips.push(['De: '+filtros.dataDe, () => { filtros.dataDe=''; document.getElementById('f-data-de').value=''; }]);
  if (filtros.dataAte) chips.push(['Até: '+filtros.dataAte, () => { filtros.dataAte=''; document.getElementById('f-data-ate').value=''; }]);
  if (filtros.busca) chips.push(['Busca: '+filtros.busca, () => { filtros.busca=''; document.getElementById('f-busca').value=''; }]);
  div.innerHTML = chips.map((c,i) => `<span class="active-filter" onclick="removerFiltro(${i})">${c[0]} ✕</span>`).join('');
  window.__chipHandlers = chips.map(c => c[1]);
}
function removerFiltro(i) { window.__chipHandlers[i](); paginaAtual = 1; renderizar(); }

function limparFiltros() {
  filtros = { cliente:'', uf:'', cidade:'', status:'', dataDe:'', dataAte:'', busca:'', motivoDrill:'', horaDrill:null };
  ['f-cliente','f-uf','f-cidade','f-status','f-data-de','f-data-ate','f-busca'].forEach(id => document.getElementById(id).value = '');
  paginaAtual = 1;
  renderizar();
}

['f-cliente','f-uf','f-cidade','f-status'].forEach(id => {
  document.getElementById(id).addEventListener('change', e => {
    const chave = id.replace('f-','');
    filtros[chave] = e.target.value;
    paginaAtual = 1;
    renderizar();
  });
});
document.getElementById('f-data-de').addEventListener('change', e => { filtros.dataDe = e.target.value; paginaAtual=1; renderizar(); });
document.getElementById('f-data-ate').addEventListener('change', e => { filtros.dataAte = e.target.value; paginaAtual=1; renderizar(); });
document.getElementById('f-busca').addEventListener('input', e => { filtros.busca = e.target.value; paginaAtual=1; renderizar(); });

document.querySelectorAll('thead th').forEach(th => {
  th.addEventListener('click', () => {
    const col = th.dataset.col;
    if (ordenacao.col === col) ordenacao.dir = ordenacao.dir === 'asc' ? 'desc' : 'asc';
    else { ordenacao.col = col; ordenacao.dir = 'asc'; }
    renderizar();
  });
});

function exportarCSV() {
  const filtrados = ordenar(aplicarFiltros());
  const cabecalho = ['Data/Hora','Unidade','Cidade','UF','Cliente','Motivo','Status','Tempo Total (min)','Operador'];
  const linhas = filtrados.map(r => [
    r.dt ? r.dt.toISOString() : '', r.unidade, r.cidade, r.uf, r.cliente, r.motivo, r.status, r.tempo_total ?? '', r.operador
  ]);
  const csv = [cabecalho, ...linhas].map(l => l.map(v => `"${String(v).replace(/"/g,'""')}"`).join(',')).join('\n');
  const blob = new Blob(['\ufeff' + csv], {type:'text/csv;charset=utf-8;'});
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url; a.download = `disparos_alarme_${new Date().toISOString().slice(0,10)}.csv`;
  a.click();
  URL.revokeObjectURL(url);
}

// ---- Graficos ----
Chart.defaults.color = '#7a8ba8';
Chart.defaults.font.family = 'Inter, sans-serif';

const chartMes = new Chart(document.getElementById('chart-mes'), {
  type: 'bar',
  data: {
    labels: AGREGADOS.volume_mes.map(v => v.mes),
    datasets: [{ data: AGREGADOS.volume_mes.map(v => v.total), backgroundColor: '#16c79a' }]
  },
  options: {
    responsive:true, maintainAspectRatio:false,
    plugins: { legend: { display:false } },
    scales: { x: { grid:{display:false} }, y: { grid:{color:'#1e2d4a'} } }
  }
});

const chartStatus = new Chart(document.getElementById('chart-status'), {
  type: 'doughnut',
  data: {
    labels: AGREGADOS.status_dist.map(s => s.status),
    datasets: [{ data: AGREGADOS.status_dist.map(s => s.total), backgroundColor: ['#27ae60','#e74c3c','#e67e22','#3498db'] }]
  },
  options: {
    responsive:true, maintainAspectRatio:false,
    plugins: { legend: { position:'bottom', labels:{boxWidth:10,font:{size:10}} } },
    onClick: (evt, elems) => {
      if (!elems.length) return;
      const idx = elems[0].index;
      const status = AGREGADOS.status_dist[idx].status;
      filtros.status = status; document.getElementById('f-status').value = status;
      paginaAtual = 1; renderizar();
    }
  }
});

const chartCidades = new Chart(document.getElementById('chart-cidades'), {
  type: 'bar',
  data: {
    labels: AGREGADOS.top_cidades.map(c => c.cidade),
    datasets: [{ data: AGREGADOS.top_cidades.map(c => c.total), backgroundColor: '#3498db' }]
  },
  options: {
    indexAxis: 'y', responsive:true, maintainAspectRatio:false,
    plugins: { legend: { display:false } },
    scales: { x: { grid:{color:'#1e2d4a'} }, y: { grid:{display:false} } },
    onClick: (evt, elems) => {
      if (!elems.length) return;
      const idx = elems[0].index;
      const cidade = AGREGADOS.top_cidades[idx].cidade;
      filtros.cidade = cidade; document.getElementById('f-cidade').value = cidade;
      paginaAtual = 1; renderizar();
    }
  }
});

const chartMotivos = new Chart(document.getElementById('chart-motivos'), {
  type: 'bar',
  data: {
    labels: AGREGADOS.top_motivos.map(m => m.motivo.length > 28 ? m.motivo.slice(0,28)+'…' : m.motivo),
    datasets: [{
      data: AGREGADOS.top_motivos.map(m => m.total),
      backgroundColor: AGREGADOS.top_motivos.map(m => {
        if (m.motivo.includes('FAC')) return '#e67e22';   // falha de energia - laranja
        if (m.motivo.includes('FTH')) return '#00bcd4';   // falha de ethernet - ciano
        return '#9b59b6';
      })
    }]
  },
  options: {
    indexAxis: 'y', responsive:true, maintainAspectRatio:false,
    plugins: { legend: { display:false } },
    scales: { x: { grid:{color:'#1e2d4a'} }, y: { grid:{display:false}, ticks:{font:{size:10}} } },
    onClick: (evt, elems) => {
      if (!elems.length) return;
      const idx = elems[0].index;
      const motivo = AGREGADOS.top_motivos[idx].motivo;
      filtros.motivoDrill = motivo;
      paginaAtual = 1; renderizar();
    }
  }
});

const chartHora = new Chart(document.getElementById('chart-hora'), {
  type: 'bar',
  data: {
    labels: AGREGADOS.por_hora.map(h => h.hora),
    datasets: [{ data: AGREGADOS.por_hora.map(h => h.total), backgroundColor: '#16c79a', borderRadius: 4 }]
  },
  options: {
    responsive:true, maintainAspectRatio:false,
    plugins: { legend: { display:false } },
    scales: { x: { grid:{display:false} }, y: { grid:{color:'#1e2d4a'} } },
    onClick: (evt, elems) => {
      if (!elems.length) return;
      const idx = elems[0].index;
      filtros.horaDrill = idx; // 0-23
      paginaAtual = 1; renderizar();
    }
  }
});

renderizar();
</script>
</body>
</html>
"""


def gerar_html(dataset, kpis, agregados, caminho_saida):
    html = TEMPLATE
    html = html.replace("__DATASET_JSON__", json.dumps(dataset, ensure_ascii=False))
    html = html.replace("__KPIS_JSON__", json.dumps(kpis, ensure_ascii=False))
    html = html.replace("__AGREGADOS_JSON__", json.dumps(agregados, ensure_ascii=False))
    html = html.replace("__DATA_GERACAO__", datetime.now().strftime("%d/%m/%Y %H:%M"))
    html = html.replace("__TOTAL__", str(kpis["total"]))
    with open(caminho_saida, "w", encoding="utf-8") as f:
        f.write(html)


# =====================================================================
# 5) MAIN
# =====================================================================

def main():
    if len(sys.argv) < 2:
        print("Uso: python3 build_disparos_alarme.py caminho/para/BASE_DE_DISPARO.xlsx")
        sys.exit(1)

    caminho_xlsx = sys.argv[1]
    t0 = time.time()
    print("Extraindo dados do xlsx (isso pode levar ~1 min em arquivos grandes)...")
    linhas = extrair_xlsx_bloatado(caminho_xlsx, "tmp_extract")
    print(f"  {len(linhas)} linhas lidas em {round(time.time()-t0,1)}s")

    registros = processar_linhas(linhas)
    print(f"  {len(registros)} registros validos processados")

    dataset = montar_dataset(registros)
    kpis, agregados = montar_kpis_e_agregados(registros)

    gerar_html(dataset, kpis, agregados, "disparos_alarme.html")
    print("Dashboard gerado: disparos_alarme.html")


if __name__ == "__main__":
    main()
