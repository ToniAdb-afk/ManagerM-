import pandas as pd
import numpy as np
import datetime as dt
import json
import os

SRC = os.path.join(os.path.dirname(__file__), "manager_data", "Manutencoes_Corretivas.xlsx")
TEMPLATE = os.path.join(os.path.dirname(__file__), "manutencao_corretiva_template.html")
OUT = os.path.join(os.path.dirname(__file__), "manutencao_corretiva.html")

MESES_PT = {1:'Jan',2:'Fev',3:'Mar',4:'Abr',5:'Mai',6:'Jun',7:'Jul',8:'Ago',9:'Set',10:'Out',11:'Nov',12:'Dez'}

PERIFERICO_NORM = {
    'SISMICO': 'SÍSMICO', 'SÍSMICO': 'SÍSMICO',
    'MAGNETICO': 'MAGNÉTICO', 'MAGNÉTICO': 'MAGNÉTICO',
    'CÂMERA': 'CÂMERA', 'CAMERA': 'CÂMERA',
}

def clean_tecnico(v):
    if pd.isna(v) or v == 0:
        return 'Não informado'
    s = str(v).replace('\xa0', ' ').strip()
    s = ' '.join(s.split())
    return s.title()

def clean_upper(v, default='Não informado'):
    if pd.isna(v):
        return default
    s = str(v).replace('\xa0', ' ').strip()
    s = ' '.join(s.split())
    return s.upper()

def load():
    df = pd.read_excel(SRC, sheet_name='Sheet1', header=1)
    df['Status'] = df['Status'].apply(lambda v: clean_upper(v, 'NÃO INFORMADO'))
    df['tecnico'] = df['Técnico'].apply(clean_tecnico)
    df['unidade'] = df['Nome da Unidade'].apply(lambda v: clean_upper(v, 'NÃO INFORMADA'))
    df['sistema'] = df['Sistema'].apply(lambda v: clean_upper(v, 'NÃO INFORMADO'))
    df['periferico'] = df['Periférico'].apply(lambda v: clean_upper(v, 'NÃO INFORMADO'))
    df['periferico'] = df['periferico'].replace(PERIFERICO_NORM)
    df['Data de Atendimento'] = pd.to_datetime(df['Data de Atendimento'])
    df['mes_key'] = df['Data de Atendimento'].dt.to_period('M').astype(str)
    return df

def top_counts(series, n=10):
    c = series.value_counts().head(n)
    return {'labels': c.index.tolist(), 'valores': [int(x) for x in c.values]}

def build_data(df):
    total = len(df)
    unidades = int(df['unidade'].nunique())
    tecnicos = int(df[df['tecnico'] != 'Não informado']['tecnico'].nunique())

    status_counts = df['Status'].value_counts()
    finalizado = int(status_counts.get('FINALIZADO', 0))
    agendado = int(status_counts.get('AGENDADO', 0))
    pendente = int(status_counts.get('PENDENTE', 0))
    outros = int(total - finalizado - agendado - pendente)
    pct_finalizado = round(100 * finalizado / total, 1) if total > 0 else None

    vol = df.groupby('mes_key').size().sort_index()
    meses_labels = []
    for k in vol.index:
        ano, mes = k.split('-')
        meses_labels.append(f"{MESES_PT[int(mes)]}/{ano[2:]}")
    vol_valores = [int(x) for x in vol.values]

    hoje = pd.Timestamp(dt.date.today())
    abertos = df[df['Status'].isin(['AGENDADO', 'PENDENTE'])]
    atrasados = int((abertos['Data de Atendimento'] < hoje).sum())

    sistema_counts = df['sistema'].value_counts()
    sistema_labels = sistema_counts.index.tolist()
    sistema_valores = [int(x) for x in sistema_counts.values]

    registros = []
    for _, r in df.iterrows():
        hist = r['Histórico'] if pd.notna(r['Histórico']) else ''
        hist = ' '.join(str(hist).split())
        registros.append({
            'ap': str(r['A2']),
            'unidade': r['unidade'],
            'tecnico': r['tecnico'],
            'data': r['Data de Atendimento'].strftime('%d/%m/%Y'),
            'data_iso': r['Data de Atendimento'].strftime('%Y-%m-%d'),
            'status': r['Status'],
            'sistema': r['sistema'],
            'periferico': r['periferico'],
            'historico': hist,
            'historico_curto': (hist[:110] + '…') if len(hist) > 110 else hist,
        })

    return {
        'kpi': {
            'total': total,
            'unidades': unidades,
            'tecnicos': tecnicos,
            'finalizado': finalizado,
            'agendado': agendado,
            'pendente': pendente,
            'outros': outros,
            'pct_finalizado': pct_finalizado,
            'atrasados': atrasados,
            'data_min': df['Data de Atendimento'].min().strftime('%Y-%m-%d'),
            'data_max': df['Data de Atendimento'].max().strftime('%Y-%m-%d'),
        },
        'volume_mensal': {'labels': meses_labels, 'valores': vol_valores},
        'status': {'labels': ['Finalizado', 'Agendado', 'Pendente'], 'valores': [finalizado, agendado, pendente]},
        'top_unidades': top_counts(df['unidade'], 10),
        'top_tecnicos': top_counts(df[df['tecnico'] != 'Não informado']['tecnico'], 10),
        'sistema': {'labels': sistema_labels, 'valores': sistema_valores},
        'top_periferico': top_counts(df[df['periferico'] != 'NÃO INFORMADO']['periferico'], 10),
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
    print(f"Finalizados: {data['kpi']['finalizado']} ({data['kpi']['pct_finalizado']}%)")
    print(f"Agendados: {data['kpi']['agendado']} | Pendentes: {data['kpi']['pendente']}")
    print(f"Em aberto e atrasados (data passada): {data['kpi']['atrasados']}")

if __name__ == '__main__':
    main()
