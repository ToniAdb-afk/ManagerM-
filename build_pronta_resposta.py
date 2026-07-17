import pandas as pd
import numpy as np
import datetime as dt
import json
import re
import sys
import os

SRC = os.path.join(os.path.dirname(__file__), "manager_data", "Pronta Resposta.xlsx")
TEMPLATE = os.path.join(os.path.dirname(__file__), "pronta_resposta_template.html")
OUT = os.path.join(os.path.dirname(__file__), "pronta_resposta.html")

MESES_PT = {1:'Jan',2:'Fev',3:'Mar',4:'Abr',5:'Mai',6:'Jun',7:'Jul',8:'Ago',9:'Set',10:'Out',11:'Nov',12:'Dez'}

def to_minutes(v):
    if pd.isna(v):
        return None
    if isinstance(v, dt.timedelta):
        return v.total_seconds()/60
    if isinstance(v, dt.time):
        return v.hour*60 + v.minute + v.second/60
    if isinstance(v, dt.datetime):
        return v.hour*60 + v.minute + v.second/60
    return None

def time_str(v):
    if pd.isna(v):
        return None
    if isinstance(v, dt.timedelta):
        total = int(v.total_seconds())
        h = (total // 3600) % 24
        m = (total % 3600) // 60
        return f"{h:02d}:{m:02d}"
    if isinstance(v, (dt.time, dt.datetime)):
        return f"{v.hour:02d}:{v.minute:02d}"
    return None

def load():
    df = pd.read_excel(SRC, sheet_name='Sheet1', header=1)
    df['Tipo Serviço'] = df['Tipo Serviço'].astype(str).str.strip()
    df['Empresa'] = df['Empresa'].replace({'Seegsing': 'Segsing'})
    df['Data Acionamento'] = pd.to_datetime(df['Data Acionamento'])
    df['tempo_min'] = df['Tempo p/ Chegar'].apply(to_minutes)
    df['hora_acionamento'] = df['Acionamento'].apply(time_str)
    df['mes_key'] = df['Data Acionamento'].dt.to_period('M').astype(str)
    return df

def build_data(df):
    total = len(df)
    postos = int(df['AP'].nunique())
    cidades = int(df['CIDADE'].nunique())

    sla_counts = df['SLA'].value_counts(dropna=True)
    dentro = int(sla_counts.get('DENTRO DO SLA', 0))
    fora = int(sla_counts.get('FORA DO SLA', 0))
    sla_pct = round(100 * dentro / (dentro + fora), 1) if (dentro + fora) > 0 else None

    tempo_mediano = df['tempo_min'].median()
    tempo_medio = df['tempo_min'].mean()

    data_min = df['Data Acionamento'].min().strftime('%Y-%m-%d')
    data_max = df['Data Acionamento'].max().strftime('%Y-%m-%d')

    # Volume mensal
    vol = df.groupby('mes_key').size().sort_index()
    meses_labels = []
    for k in vol.index:
        ano, mes = k.split('-')
        meses_labels.append(f"{MESES_PT[int(mes)]}/{ano[2:]}")
    vol_valores = [int(x) for x in vol.values]

    queda_pct = None
    if len(vol_valores) >= 2 and vol_valores[0] > 0:
        queda_pct = round(100 * (vol_valores[0] - vol_valores[-1]) / vol_valores[0], 1)

    # Top cidades
    top_cidades = df['CIDADE'].value_counts().head(10)
    top_cidades_labels = top_cidades.index.tolist()
    top_cidades_valores = [int(x) for x in top_cidades.values]

    # Tipo de servico
    tipo_counts = df['Tipo Serviço'].value_counts()
    tipo_labels = tipo_counts.index.tolist()
    tipo_valores = [int(x) for x in tipo_counts.values]

    # Motivo (top 8)
    motivo_counts = df['Motivo'].value_counts().head(8)
    motivo_labels = motivo_counts.index.tolist()
    motivo_valores = [int(x) for x in motivo_counts.values]

    # Empresa prestadora
    empresa_counts = df['Empresa'].value_counts(dropna=True)
    empresa_labels = empresa_counts.index.tolist()
    empresa_valores = [int(x) for x in empresa_counts.values]

    # Solicitante
    solic_counts = df['SOLICITANTE'].value_counts()
    solic_labels = solic_counts.index.tolist()
    solic_valores = [int(x) for x in solic_counts.values]

    # Tabela detalhada
    registros = []
    for _, r in df.iterrows():
        registros.append({
            'ap': str(r['AP']),
            'cidade': r['CIDADE'],
            'uf': r['UF'],
            'tipo': r['Tipo Serviço'],
            'descricao': r['Descrição'],
            'motivo': r['Motivo'],
            'solicitante': r['SOLICITANTE'],
            'empresa': r['Empresa'] if pd.notna(r['Empresa']) else '-',
            'data': r['Data Acionamento'].strftime('%d/%m/%Y'),
            'data_iso': r['Data Acionamento'].strftime('%Y-%m-%d'),
            'hora': r['hora_acionamento'] or '-',
            'tempo_min': None if pd.isna(r['tempo_min']) else round(r['tempo_min'], 1),
            'sla': r['SLA'] if pd.notna(r['SLA']) else '-',
        })

    return {
        'kpi': {
            'total': total,
            'postos': postos,
            'cidades': cidades,
            'sla_pct': sla_pct,
            'dentro': dentro,
            'fora': fora,
            'tempo_mediano': None if pd.isna(tempo_mediano) else round(tempo_mediano, 0),
            'tempo_medio': None if pd.isna(tempo_medio) else round(tempo_medio, 0),
            'queda_pct': queda_pct,
            'data_min': data_min,
            'data_max': data_max,
        },
        'volume_mensal': {'labels': meses_labels, 'valores': vol_valores},
        'top_cidades': {'labels': top_cidades_labels, 'valores': top_cidades_valores},
        'tipo_servico': {'labels': tipo_labels, 'valores': tipo_valores},
        'motivo': {'labels': motivo_labels, 'valores': motivo_valores},
        'empresa': {'labels': empresa_labels, 'valores': empresa_valores},
        'solicitante': {'labels': solic_labels, 'valores': solic_valores},
        'registros': registros,
        'gerado_em': dt.datetime.now().strftime('%d/%m/%Y %H:%M'),
    }

def main():
    df = load()
    data = build_data(df)
    with open(TEMPLATE, 'r', encoding='utf-8') as f:
        template = f.read()
    payload = json.dumps(data, ensure_ascii=False)
    html = template.replace('__DASHBOARD_DATA__', payload)
    with open(OUT, 'w', encoding='utf-8') as f:
        f.write(html)
    print(f"OK -> {OUT}")
    print(f"Total de registros: {data['kpi']['total']}")
    print(f"SLA: {data['kpi']['sla_pct']}%")
    print(f"Queda de acionamentos (primeiro -> ultimo mes): {data['kpi']['queda_pct']}%")

if __name__ == '__main__':
    main()
