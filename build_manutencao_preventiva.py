import pandas as pd
import numpy as np
import datetime as dt
import json
import os

SRC = os.path.join(os.path.dirname(__file__), "manager_data", "Manutencoes_Preventivas.xlsx")
TEMPLATE = os.path.join(os.path.dirname(__file__), "manutencao_preventiva_template.html")
OUT = os.path.join(os.path.dirname(__file__), "manutencao_preventiva.html")

UF_VALIDAS = {'AC','AL','AP','AM','BA','CE','DF','ES','GO','MA','MT','MS','MG','PA','PB',
              'PR','PE','PI','RJ','RN','RS','RO','RR','SC','SP','SE','TO'}

def clean_upper(v, default='NÃO INFORMADO'):
    if pd.isna(v):
        return default
    s = str(v).replace('\xa0', ' ').strip()
    s = ' '.join(s.split())
    return s.upper() if s else default

def clean_uf(v):
    s = clean_upper(v)
    return s if s in UF_VALIDAS else 'NÃO INFORMADO'

def load():
    df = pd.read_excel(SRC, sheet_name='Sheet1', header=0)
    df['status'] = df['Status'].apply(clean_upper)
    df['uf'] = df['UF'].apply(clean_uf)
    df['cidade'] = df['CIDADE'].apply(lambda v: clean_upper(v, 'NÃO INFORMADA'))
    df['nome'] = df['Nome'].astype(str).str.strip()
    return df

def top_counts(series, n=10):
    c = series.value_counts().head(n)
    return {'labels': c.index.tolist(), 'valores': [int(x) for x in c.values]}

def build_data(df):
    total = len(df)
    cidades = int(df['cidade'].nunique())
    estados = int(df[df['uf'] != 'NÃO INFORMADO']['uf'].nunique())

    status_counts = df['status'].value_counts()
    finalizado = int(status_counts.get('FINALIZADO', 0))
    agendado = int(status_counts.get('AGENDADO', 0))
    outros = int(total - finalizado - agendado)
    pct_finalizado = round(100 * finalizado / total, 1) if total > 0 else None

    # progresso por UF (finalizado vs agendado), ordenado por volume total desc
    uf_grp = df.groupby('uf')['status'].value_counts().unstack(fill_value=0)
    for col in ['FINALIZADO', 'AGENDADO']:
        if col not in uf_grp.columns:
            uf_grp[col] = 0
    uf_grp['total'] = uf_grp['FINALIZADO'] + uf_grp['AGENDADO']
    uf_grp = uf_grp.sort_values('total', ascending=False).head(15)
    uf_labels = uf_grp.index.tolist()
    uf_finalizado = [int(x) for x in uf_grp['FINALIZADO'].tolist()]
    uf_agendado = [int(x) for x in uf_grp['AGENDADO'].tolist()]

    registros = []
    for _, r in df.iterrows():
        registros.append({
            'pa': str(r['PA']),
            'nome': r['nome'],
            'cidade': r['cidade'],
            'uf': r['uf'],
            'status': r['status'],
        })

    return {
        'kpi': {
            'total': total,
            'cidades': cidades,
            'estados': estados,
            'finalizado': finalizado,
            'agendado': agendado,
            'outros': outros,
            'pct_finalizado': pct_finalizado,
        },
        'status': {'labels': ['Finalizado', 'Agendado'], 'valores': [finalizado, agendado]},
        'uf_progresso': {'labels': uf_labels, 'finalizado': uf_finalizado, 'agendado': uf_agendado},
        'top_cidades': top_counts(df['cidade'], 10),
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
    print(f"Total de unidades: {data['kpi']['total']}")
    print(f"Finalizadas: {data['kpi']['finalizado']} ({data['kpi']['pct_finalizado']}%)")
    print(f"Agendadas: {data['kpi']['agendado']}")
    print(f"Estados: {data['kpi']['estados']} | Cidades: {data['kpi']['cidades']}")

if __name__ == '__main__':
    main()
