# -*- coding: utf-8 -*-
"""
build_ocorrencias.py
-------------------------------------------------------------------
Le a planilha de Ocorrências (aba 'Ocorrências') e gera o dashboard
'ocorrencias.html' com filtros e drill-down.

Uso:
    python3 build_ocorrencias.py caminho/para/Ocorrencias.xlsx
"""

import sys
import json
from datetime import datetime
import pandas as pd

MESES_ORDEM = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
               "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]


def carregar(caminho_xlsx):
    df = pd.read_excel(caminho_xlsx, sheet_name="Ocorrências")
    df.columns = [c.strip() for c in df.columns]
    return df


def montar_registros(df):
    registros = []
    for _, row in df.iterrows():
        data_raw = str(row["Data"]).strip()
        data_iso = None
        if data_raw and data_raw.lower() != "não informada" and data_raw.lower() != "nan":
            try:
                data_iso = datetime.strptime(data_raw, "%d/%m/%Y").strftime("%Y-%m-%d")
            except Exception:
                data_iso = None
        registros.append({
            "n": int(row["Nº"]),
            "data": data_iso,
            "data_raw": data_raw if data_raw.lower() != "nan" else "Não informada",
            "mes": str(row["Mês"]).strip(),
            "pa": str(row["PA"]).strip(),
            "local": str(row["Local"]).strip(),
            "ocorrencia": str(row["Ocorrência"]).strip(),
            "tipo": str(row["Tipo"]).strip(),
            "acao": str(row["Observação / Ação"]).strip(),
        })
    return registros


def montar_kpis_e_agregados(registros):
    total = len(registros)
    unidades_distintas = len({r["pa"] for r in registros})
    com_data = sum(1 for r in registros if r["data"])

    cont_tipo = {}
    for r in registros:
        cont_tipo[r["tipo"]] = cont_tipo.get(r["tipo"], 0) + 1
    tipo_mais_frequente = max(cont_tipo.items(), key=lambda x: x[1])[0] if cont_tipo else "-"

    meses_presentes = sorted({r["mes"] for r in registros},
                              key=lambda m: MESES_ORDEM.index(m) if m in MESES_ORDEM else 99)

    kpis = {
        "total": total,
        "unidades_distintas": unidades_distintas,
        "meses_cobertos": len(meses_presentes),
        "tipo_mais_frequente": tipo_mais_frequente,
        "pct_com_data": round(100 * com_data / total, 1) if total else 0,
    }

    volume_mes = [{"mes": m, "total": cont} for m, cont in
                  sorted(({m: sum(1 for r in registros if r["mes"] == m) for m in meses_presentes}).items(),
                         key=lambda x: MESES_ORDEM.index(x[0]) if x[0] in MESES_ORDEM else 99)]

    top_tipos = [{"tipo": t, "total": c} for t, c in sorted(cont_tipo.items(), key=lambda x: -x[1])]

    cont_pa = {}
    for r in registros:
        chave = f'{r["pa"]} - {r["local"]}'
        cont_pa[chave] = cont_pa.get(chave, 0) + 1
    top_pas = [{"pa": p, "total": c} for p, c in sorted(cont_pa.items(), key=lambda x: -x[1]) if c > 1]

    agregados = {
        "volume_mes": volume_mes,
        "top_tipos": top_tipos,
        "top_pas": top_pas,
    }
    return kpis, agregados


TEMPLATE = r"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Ocorrências · Manager+</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.min.js"></script>
<style>
  * { box-sizing:border-box; margin:0; padding:0; }
  body { font-family:'Inter',-apple-system,sans-serif; background:#080b10; color:#e8edf5; min-height:100vh; }
  a.back { color:#7a8ba8; text-decoration:none; font-size:12px; }
  a.back:hover { color:#3498db; }
  header { background:#0e1420; border-bottom:1px solid #1e2d4a; padding:18px 28px; display:flex; align-items:center; justify-content:space-between; flex-wrap:wrap; gap:10px; }
  header h1 { font-size:18px; font-weight:800; }
  header h1 span { color:#3498db; }
  main { padding:24px 28px 60px; max-width:1500px; margin:0 auto; }

  .kpis { display:grid; grid-template-columns:repeat(auto-fit,minmax(160px,1fr)); gap:14px; margin-bottom:22px; }
  .kpi { background:#0e1420; border:1px solid #1e2d4a; border-radius:14px; padding:16px 18px; }
  .kpi .label { font-size:11px; color:#7a8ba8; text-transform:uppercase; letter-spacing:.5px; margin-bottom:6px; }
  .kpi .value { font-size:24px; font-weight:800; color:#3498db; }
  .kpi .value.small { font-size:16px; }

  .filters { background:#0e1420; border:1px solid #1e2d4a; border-radius:14px; padding:16px 18px; margin-bottom:22px; display:flex; gap:12px; flex-wrap:wrap; align-items:end; }
  .filters .fg { display:flex; flex-direction:column; gap:5px; }
  .filters label { font-size:10px; color:#7a8ba8; text-transform:uppercase; letter-spacing:.5px; }
  .filters select, .filters input { background:#080b10; border:1px solid #1e2d4a; color:#e8edf5; border-radius:8px; padding:8px 10px; font-size:12px; min-width:130px; }
  .filters input[type=text] { min-width:220px; }
  .filters button { background:#3498db; border:none; color:#06182b; font-weight:700; padding:9px 16px; border-radius:8px; font-size:12px; cursor:pointer; }
  .filters button.secondary { background:transparent; border:1px solid #1e2d4a; color:#e8edf5; }
  .active-filter { display:inline-flex; align-items:center; gap:6px; background:rgba(52,152,219,.12); border:1px solid rgba(52,152,219,.35); color:#3498db; font-size:11px; padding:4px 10px; border-radius:20px; margin:2px 4px 2px 0; cursor:pointer; }

  .charts { display:grid; grid-template-columns:1.2fr 1fr; gap:16px; margin-bottom:22px; }
  .chart-card { background:#0e1420; border:1px solid #1e2d4a; border-radius:14px; padding:18px; }
  .chart-card h3 { font-size:13px; font-weight:700; margin-bottom:12px; color:#e8edf5; }
  .chart-card .hint { font-size:10px; color:#4a5f7a; margin-top:8px; }
  .chart-wrap { position:relative; height:260px; }

  table { width:100%; border-collapse:collapse; font-size:12px; }
  thead th { text-align:left; padding:10px 12px; color:#7a8ba8; font-size:10px; text-transform:uppercase; letter-spacing:.5px; border-bottom:1px solid #1e2d4a; cursor:pointer; user-select:none; white-space:nowrap; }
  thead th:hover { color:#3498db; }
  tbody td { padding:9px 12px; border-bottom:1px solid #131b28; vertical-align:top; }
  tbody tr:hover { background:#0e1420; }
  .table-card { background:#0e1420; border:1px solid #1e2d4a; border-radius:14px; padding:18px; overflow-x:auto; }
  .table-top { display:flex; justify-content:space-between; align-items:center; margin-bottom:14px; flex-wrap:wrap; gap:10px; }
  .table-top .count { font-size:12px; color:#7a8ba8; }
  .tipo-pill { display:inline-block; padding:2px 8px; border-radius:10px; font-size:10px; font-weight:700; background:rgba(52,152,219,.15); color:#3498db; white-space:nowrap; }
</style>
</head>
<body>

<header>
  <div>
    <a class="back" href="manager-plus-index.html">&larr; Voltar ao Hub</a>
    <h1 style="margin-top:4px;">📝 <span>Ocorrências</span></h1>
  </div>
  <div style="font-size:11px;color:#4a5f7a;">Gerado em __DATA_GERACAO__ · __TOTAL__ registros</div>
</header>

<main>
  <div class="kpis">
    <div class="kpi"><div class="label">Total de Ocorrências</div><div class="value" id="kpi-total">-</div></div>
    <div class="kpi"><div class="label">Unidades Distintas (PA)</div><div class="value" id="kpi-unidades">-</div></div>
    <div class="kpi"><div class="label">Meses Cobertos</div><div class="value" id="kpi-meses">-</div></div>
    <div class="kpi"><div class="label">Tipo Mais Frequente</div><div class="value small" id="kpi-tipo">-</div></div>
    <div class="kpi"><div class="label">% Com Data Informada</div><div class="value" id="kpi-pctdata">-</div></div>
  </div>

  <div class="filters">
    <div class="fg"><label>Mês</label><select id="f-mes"><option value="">Todos</option></select></div>
    <div class="fg"><label>Tipo</label><select id="f-tipo"><option value="">Todos</option></select></div>
    <div class="fg"><label>Buscar (PA / local / ocorrência)</label><input type="text" id="f-busca" placeholder="ex: cliente retido, Maceió, 413..."></div>
    <div class="fg"><button onclick="limparFiltros()" class="secondary">Limpar filtros</button></div>
    <div class="fg"><button onclick="exportarCSV()">Exportar CSV</button></div>
  </div>
  <div id="filtros-ativos" style="margin:-10px 0 18px;"></div>

  <div class="charts">
    <div class="chart-card">
      <h3>Ocorrências por Mês</h3>
      <div class="chart-wrap"><canvas id="chart-mes"></canvas></div>
      <div class="hint">Clique numa barra para filtrar por mês</div>
    </div>
    <div class="chart-card">
      <h3>Por Tipo</h3>
      <div class="chart-wrap"><canvas id="chart-tipo"></canvas></div>
      <div class="hint">Clique numa barra para filtrar por tipo</div>
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
          <th data-col="n">Nº</th>
          <th data-col="data">Data</th>
          <th data-col="mes">Mês</th>
          <th data-col="pa">PA</th>
          <th data-col="local">Local</th>
          <th data-col="ocorrencia">Ocorrência</th>
          <th data-col="tipo">Tipo</th>
          <th data-col="acao">Observação / Ação</th>
        </tr>
      </thead>
      <tbody id="tbody"></tbody>
    </table>
    </div>
  </div>
</main>

<script>
const REGISTROS = __REGISTROS_JSON__;
const KPIS = __KPIS_JSON__;
const AGREGADOS = __AGREGADOS_JSON__;

function popularSelect(id, valores) {
  const sel = document.getElementById(id);
  valores.forEach(v => {
    const opt = document.createElement('option');
    opt.value = v; opt.textContent = v;
    sel.appendChild(opt);
  });
}
const mesesOrdem = ["Janeiro","Fevereiro","Março","Abril","Maio","Junho","Julho","Agosto","Setembro","Outubro","Novembro","Dezembro"];
const mesesPresentes = [...new Set(REGISTROS.map(r=>r.mes))].sort((a,b)=>mesesOrdem.indexOf(a)-mesesOrdem.indexOf(b));
const tiposPresentes = [...new Set(REGISTROS.map(r=>r.tipo))].sort();
popularSelect('f-mes', mesesPresentes);
popularSelect('f-tipo', tiposPresentes);

document.getElementById('kpi-total').textContent = KPIS.total;
document.getElementById('kpi-unidades').textContent = KPIS.unidades_distintas;
document.getElementById('kpi-meses').textContent = KPIS.meses_cobertos;
document.getElementById('kpi-tipo').textContent = KPIS.tipo_mais_frequente;
document.getElementById('kpi-pctdata').textContent = KPIS.pct_com_data + '%';

let filtros = { mes:'', tipo:'', busca:'' };
let ordenacao = { col:'n', dir:'asc' };

function aplicarFiltros() {
  return REGISTROS.filter(r => {
    if (filtros.mes && r.mes !== filtros.mes) return false;
    if (filtros.tipo && r.tipo !== filtros.tipo) return false;
    if (filtros.busca) {
      const b = filtros.busca.toLowerCase();
      if (!r.pa.toLowerCase().includes(b) && !r.local.toLowerCase().includes(b) &&
          !r.ocorrencia.toLowerCase().includes(b) && !r.tipo.toLowerCase().includes(b)) return false;
    }
    return true;
  });
}

function ordenar(lista) {
  const { col, dir } = ordenacao;
  const mult = dir === 'asc' ? 1 : -1;
  return lista.slice().sort((a,b) => {
    let va = a[col], vb = b[col];
    if (col === 'data') { va = a.data || '0000'; vb = b.data || '0000'; }
    if (typeof va === 'string') return va.localeCompare(vb) * mult;
    return (va - vb) * mult;
  });
}

function renderizar() {
  const filtrados = ordenar(aplicarFiltros());
  document.getElementById('count-mostrando').textContent = filtrados.length;
  document.getElementById('count-total').textContent = REGISTROS.length;

  const tbody = document.getElementById('tbody');
  tbody.innerHTML = filtrados.map(r => `
    <tr>
      <td>${r.n}</td>
      <td>${r.data_raw}</td>
      <td>${r.mes}</td>
      <td>${r.pa}</td>
      <td>${r.local}</td>
      <td>${r.ocorrencia}</td>
      <td><span class="tipo-pill">${r.tipo}</span></td>
      <td>${r.acao}</td>
    </tr>
  `).join('');

  renderizarFiltrosAtivos();
}

function renderizarFiltrosAtivos() {
  const div = document.getElementById('filtros-ativos');
  const chips = [];
  if (filtros.mes) chips.push(['Mês: '+filtros.mes, () => { filtros.mes=''; document.getElementById('f-mes').value=''; }]);
  if (filtros.tipo) chips.push(['Tipo: '+filtros.tipo, () => { filtros.tipo=''; document.getElementById('f-tipo').value=''; }]);
  if (filtros.busca) chips.push(['Busca: '+filtros.busca, () => { filtros.busca=''; document.getElementById('f-busca').value=''; }]);
  div.innerHTML = chips.map((c,i) => `<span class="active-filter" onclick="removerFiltro(${i})">${c[0]} ✕</span>`).join('');
  window.__chipHandlers = chips.map(c => c[1]);
}
function removerFiltro(i) { window.__chipHandlers[i](); renderizar(); }

function limparFiltros() {
  filtros = { mes:'', tipo:'', busca:'' };
  ['f-mes','f-tipo','f-busca'].forEach(id => document.getElementById(id).value = '');
  renderizar();
}

['f-mes','f-tipo'].forEach(id => {
  document.getElementById(id).addEventListener('change', e => {
    const chave = id.replace('f-','');
    filtros[chave] = e.target.value;
    renderizar();
  });
});
document.getElementById('f-busca').addEventListener('input', e => { filtros.busca = e.target.value; renderizar(); });

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
  const cabecalho = ['Nº','Data','Mês','PA','Local','Ocorrência','Tipo','Observação / Ação'];
  const linhas = filtrados.map(r => [r.n, r.data_raw, r.mes, r.pa, r.local, r.ocorrencia, r.tipo, r.acao]);
  const csv = [cabecalho, ...linhas].map(l => l.map(v => `"${String(v).replace(/"/g,'""')}"`).join(',')).join('\n');
  const blob = new Blob(['\ufeff' + csv], {type:'text/csv;charset=utf-8;'});
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url; a.download = `ocorrencias_${new Date().toISOString().slice(0,10)}.csv`;
  a.click();
  URL.revokeObjectURL(url);
}

Chart.defaults.color = '#7a8ba8';
Chart.defaults.font.family = 'Inter, sans-serif';

const chartMes = new Chart(document.getElementById('chart-mes'), {
  type: 'bar',
  data: {
    labels: AGREGADOS.volume_mes.map(v => v.mes),
    datasets: [{ data: AGREGADOS.volume_mes.map(v => v.total), backgroundColor: '#3498db' }]
  },
  options: {
    responsive:true, maintainAspectRatio:false,
    plugins: { legend: { display:false } },
    scales: { x: { grid:{display:false} }, y: { grid:{color:'#1e2d4a'}, ticks:{stepSize:1} } },
    onClick: (evt, elems) => {
      if (!elems.length) return;
      const idx = elems[0].index;
      const mes = AGREGADOS.volume_mes[idx].mes;
      filtros.mes = mes; document.getElementById('f-mes').value = mes;
      renderizar();
    }
  }
});

const chartTipo = new Chart(document.getElementById('chart-tipo'), {
  type: 'bar',
  data: {
    labels: AGREGADOS.top_tipos.map(t => t.tipo.length > 26 ? t.tipo.slice(0,26)+'…' : t.tipo),
    datasets: [{ data: AGREGADOS.top_tipos.map(t => t.total), backgroundColor: '#16c79a' }]
  },
  options: {
    indexAxis: 'y', responsive:true, maintainAspectRatio:false,
    plugins: { legend: { display:false } },
    scales: { x: { grid:{color:'#1e2d4a'}, ticks:{stepSize:1} }, y: { grid:{display:false}, ticks:{font:{size:10}} } },
    onClick: (evt, elems) => {
      if (!elems.length) return;
      const idx = elems[0].index;
      const tipo = AGREGADOS.top_tipos[idx].tipo;
      filtros.tipo = tipo; document.getElementById('f-tipo').value = tipo;
      renderizar();
    }
  }
});

renderizar();
</script>
</body>
</html>
"""


def gerar_html(registros, kpis, agregados, caminho_saida):
    html = TEMPLATE
    html = html.replace("__REGISTROS_JSON__", json.dumps(registros, ensure_ascii=False))
    html = html.replace("__KPIS_JSON__", json.dumps(kpis, ensure_ascii=False))
    html = html.replace("__AGREGADOS_JSON__", json.dumps(agregados, ensure_ascii=False))
    html = html.replace("__DATA_GERACAO__", datetime.now().strftime("%d/%m/%Y %H:%M"))
    html = html.replace("__TOTAL__", str(kpis["total"]))
    with open(caminho_saida, "w", encoding="utf-8") as f:
        f.write(html)


def main():
    if len(sys.argv) < 2:
        print("Uso: python3 build_ocorrencias.py caminho/para/Ocorrencias.xlsx")
        sys.exit(1)

    df = carregar(sys.argv[1])
    registros = montar_registros(df)
    kpis, agregados = montar_kpis_e_agregados(registros)
    gerar_html(registros, kpis, agregados, "ocorrencias.html")
    print(f"Dashboard gerado: ocorrencias.html ({kpis['total']} registros)")


if __name__ == "__main__":
    main()
