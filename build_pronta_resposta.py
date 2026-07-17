# -*- coding: utf-8 -*-
"""
build_pronta_resposta.py
-------------------------------------------------------------------
Le a planilha 'Pronta Resposta' (aba unica) e gera o dashboard
'pronta_resposta.html', recalculando tudo (KPIs, graficos, tabela).

Uso:
    python3 build_pronta_resposta.py caminho/para/Pronta_Resposta.xlsx
"""

import sys
import json
import re
from datetime import datetime, time as dtime
import pandas as pd

MESES_ABREV = {1: "Jan", 2: "Fev", 3: "Mar", 4: "Abr", 5: "Mai", 6: "Jun",
               7: "Jul", 8: "Ago", 9: "Set", 10: "Out", 11: "Nov", 12: "Dez"}


def carregar(caminho):
    xls = pd.ExcelFile(caminho)
    # Procura a aba que realmente tem os dados (contem a coluna 'AP'),
    # em vez de assumir que e sempre a primeira - planilhas podem ganhar
    # abas extras (resumos, rascunhos) na frente ao longo do tempo.
    aba_certa = None
    for nome in xls.sheet_names:
        cabecalho = pd.read_excel(xls, sheet_name=nome, nrows=0).columns
        cabecalho = [str(c).strip() for c in cabecalho]
        if "AP" in cabecalho and "Motivo" in cabecalho:
            aba_certa = nome
            break
    if aba_certa is None:
        raise RuntimeError(
            f"Nao encontrei uma aba com as colunas esperadas (AP, Motivo). "
            f"Abas disponiveis: {xls.sheet_names}"
        )
    df = pd.read_excel(xls, sheet_name=aba_certa)
    df.columns = [c.strip() for c in df.columns]
    print(f"  (lendo aba '{aba_certa}' de {len(xls.sheet_names)} aba(s): {xls.sheet_names})")
    return df


def normalizar_texto(v):
    if v is None:
        return ""
    s = str(v).strip()
    if s.lower() in ("nan", "-", ""):
        return ""
    return s


def para_time(v):
    """Aceita datetime.time, datetime.datetime ou string 'HH:MM:SS' e retorna datetime.time."""
    if v is None:
        return None
    if isinstance(v, dtime):
        return v
    if isinstance(v, datetime):
        return v.time()
    s = str(v).strip()
    if not s or s == "-" or s.lower() == "nan":
        return None
    for fmt in ("%H:%M:%S", "%H:%M"):
        try:
            return datetime.strptime(s, fmt).time()
        except Exception:
            continue
    return None


def minutos_entre(inicio, fim):
    """Diferenca em minutos entre dois horarios (time), assumindo que 'fim' pode
    cair no dia seguinte se for menor que 'inicio' (cruzou a meia-noite)."""
    if inicio is None or fim is None:
        return None
    ini_seg = inicio.hour * 3600 + inicio.minute * 60 + inicio.second
    fim_seg = fim.hour * 3600 + fim.minute * 60 + fim.second
    delta = fim_seg - ini_seg
    if delta < 0:
        delta += 24 * 3600
    return round(delta / 60, 1)


def montar_registros(df):
    registros = []
    for _, row in df.iterrows():
        ap = str(row.get("AP", "")).strip()
        try:
            ap = str(int(float(ap)))
        except Exception:
            pass

        cidade = normalizar_texto(row.get("CIDADE"))
        uf = normalizar_texto(row.get("UF"))
        tipo = normalizar_texto(row.get("Tipo Serviço"))
        empresa = normalizar_texto(row.get("Empresa")) or "N/D"
        descricao = normalizar_texto(row.get("Descrição"))
        motivo = normalizar_texto(row.get("Motivo"))
        solicitante = normalizar_texto(row.get("SOLICITANTE"))
        loja = normalizar_texto(row.get("Loja"))
        cnpj = normalizar_texto(row.get("CNPJ"))

        dt_acionamento = row.get("Data Acionamento")
        if isinstance(dt_acionamento, datetime):
            data_iso = dt_acionamento.strftime("%Y-%m-%d")
            data_br = dt_acionamento.strftime("%d/%m/%Y")
        else:
            data_iso, data_br = None, normalizar_texto(dt_acionamento) or "Não informada"

        hora_acionamento = para_time(row.get("Acionamento"))
        hora_chegada = para_time(row.get("Hora Chegada (Preservação)"))
        tempo_min = minutos_entre(hora_acionamento, hora_chegada)

        sla = normalizar_texto(row.get("SLA")) or "-"

        registros.append({
            "ap": ap,
            "loja": loja,
            "cnpj": cnpj,
            "cidade": cidade,
            "uf": uf,
            "tipo": tipo,
            "descricao": descricao,
            "motivo": motivo,
            "solicitante": solicitante,
            "empresa": empresa,
            "data": data_br,
            "data_iso": data_iso,
            "hora": hora_acionamento.strftime("%H:%M") if hora_acionamento else "-",
            "tempo_min": tempo_min,
            "sla": sla,
        })
    return registros


def mediana(lst):
    if not lst:
        return None
    s = sorted(lst)
    n = len(s)
    return s[n // 2] if n % 2 else (s[n // 2 - 1] + s[n // 2]) / 2


def montar_kpis_e_agregados(registros):
    total = len(registros)
    postos = len({r["ap"] for r in registros if r["ap"]})
    cidades = len({r["cidade"] for r in registros if r["cidade"]})

    com_sla = [r for r in registros if r["sla"] in ("DENTRO DO SLA", "FORA DO SLA")]
    dentro = sum(1 for r in com_sla if r["sla"] == "DENTRO DO SLA")
    fora = sum(1 for r in com_sla if r["sla"] == "FORA DO SLA")
    sla_pct = round(100 * dentro / len(com_sla), 1) if com_sla else None

    tempos = [r["tempo_min"] for r in registros if r["tempo_min"] is not None]
    tempo_mediano = mediana(tempos)
    tempo_medio = round(sum(tempos) / len(tempos), 1) if tempos else None

    datas_validas = sorted({r["data_iso"] for r in registros if r["data_iso"]})
    data_min = datas_validas[0] if datas_validas else None
    data_max = datas_validas[-1] if datas_validas else None

    # volume mensal
    cont_mes = {}
    for r in registros:
        if r["data_iso"]:
            chave = r["data_iso"][:7]  # YYYY-MM
            cont_mes[chave] = cont_mes.get(chave, 0) + 1
    meses_ordenados = sorted(cont_mes.keys())
    labels_mes = []
    for m in meses_ordenados:
        ano, mes = m.split("-")
        labels_mes.append(f"{MESES_ABREV[int(mes)]}/{ano[2:]}")
    valores_mes = [cont_mes[m] for m in meses_ordenados]

    queda_pct = None
    if len(valores_mes) >= 2 and valores_mes[0] > 0:
        queda_pct = round(100 * (valores_mes[0] - valores_mes[-1]) / valores_mes[0], 1)

    # top cidades
    cont_cidade = {}
    for r in registros:
        if r["cidade"]:
            cont_cidade[r["cidade"]] = cont_cidade.get(r["cidade"], 0) + 1
    top_cidades = sorted(cont_cidade.items(), key=lambda x: -x[1])[:10]

    # tipo de servico
    cont_tipo = {}
    for r in registros:
        t = r["tipo"] or "N/D"
        cont_tipo[t] = cont_tipo.get(t, 0) + 1
    tipo_ordenado = sorted(cont_tipo.items(), key=lambda x: -x[1])

    # motivo (na verdade, categorizado pela coluna Descrição - e o agrupamento
    # que reflete corretamente o tipo de ocorrencia, ex.: CFTV INOPERANTE,
    # QUEDA DE ENERGIA, etc. A coluna 'Motivo' e um detalhamento mais fino
    # e fica disponivel na tabela/busca, mas o grafico usa Descrição.)
    cont_motivo = {}
    for r in registros:
        m = r["descricao"] or "N/D"
        cont_motivo[m] = cont_motivo.get(m, 0) + 1
    top_motivo = sorted(cont_motivo.items(), key=lambda x: -x[1])[:8]

    # empresa
    cont_empresa = {}
    for r in registros:
        e = r["empresa"] or "N/D"
        cont_empresa[e] = cont_empresa.get(e, 0) + 1
    top_empresa = sorted(cont_empresa.items(), key=lambda x: -x[1])

    # solicitante
    cont_solicitante = {}
    for r in registros:
        s = r["solicitante"] or "N/D"
        cont_solicitante[s] = cont_solicitante.get(s, 0) + 1
    top_solicitante = sorted(cont_solicitante.items(), key=lambda x: -x[1])[:3]

    kpi = {
        "total": total,
        "postos": postos,
        "cidades": cidades,
        "sla_pct": sla_pct,
        "dentro": dentro,
        "fora": fora,
        "tempo_mediano": tempo_mediano,
        "tempo_medio": tempo_medio,
        "queda_pct": queda_pct,
        "data_min": data_min,
        "data_max": data_max,
    }

    dados = {
        "kpi": kpi,
        "volume_mensal": {"labels": labels_mes, "valores": valores_mes},
        "top_cidades": {"labels": [c for c, _ in top_cidades], "valores": [v for _, v in top_cidades]},
        "tipo_servico": {"labels": [t for t, _ in tipo_ordenado], "valores": [v for _, v in tipo_ordenado]},
        "motivo": {"labels": [m for m, _ in top_motivo], "valores": [v for _, v in top_motivo]},
        "empresa": {"labels": [e for e, _ in top_empresa], "valores": [v for _, v in top_empresa]},
        "solicitante": {"labels": [s for s, _ in top_solicitante], "valores": [v for _, v in top_solicitante]},
        "registros": registros,
    }
    return dados


def gerar_html(dados, template_path, saida_path):
    with open(template_path, "r", encoding="utf-8") as f:
        html = f.read()

    dados_json = json.dumps(dados, ensure_ascii=False)

    # Substitui o conteudo do <script id="dashboard-data" type="application/json">...</script>
    padrao = re.compile(
        r'(<script id="dashboard-data" type="application/json">).*?(</script>)',
        re.S
    )
    html_novo, n = padrao.subn(lambda m: m.group(1) + dados_json + m.group(2), html)
    if n != 1:
        raise RuntimeError(f"Esperava substituir 1 bloco de dados, substitui {n}. Verifique o template.")

    with open(saida_path, "w", encoding="utf-8") as f:
        f.write(html_novo)


def main():
    if len(sys.argv) < 2:
        print("Uso: python3 build_pronta_resposta.py caminho/para/Pronta_Resposta.xlsx [template.html]")
        sys.exit(1)

    caminho_xlsx = sys.argv[1]
    template_path = sys.argv[2] if len(sys.argv) > 2 else "pronta_resposta_template.html"

    df = carregar(caminho_xlsx)
    registros = montar_registros(df)
    dados = montar_kpis_e_agregados(registros)

    gerar_html(dados, template_path, "pronta_resposta.html")
    print(f"Dashboard gerado: pronta_resposta.html ({dados['kpi']['total']} registros)")
    print(f"KPIs: {json.dumps(dados['kpi'], ensure_ascii=False, indent=2)}")


if __name__ == "__main__":
    main()
