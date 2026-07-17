import pandas as pd
import numpy as np
import datetime as dt
import json
import os

SRC = os.path.join(os.path.dirname(__file__), "manager_data", "Faturamento.xlsx")
TEMPLATE = os.path.join(os.path.dirname(__file__), "valores_faturados_template.html")
OUT = os.path.join(os.path.dirname(__file__), "valores_faturados.html")

CATEGORIAS = ['Faturamento Recorrente', 'Manutenção Corretiva', 'Produtos', 'Pronta Resposta']
CAT_LABEL = {
    'Faturamento Recorrente': 'Recorrente',
    'Manutenção Corretiva': 'Manutenção Corretiva',
    'Produtos': 'Produtos',
    'Pronta Resposta': 'Pronta Resposta',
}

def load():
    df = pd.read_excel(SRC, sheet_name='Sheet1', header=0)
    for c in CATEGORIAS:
        df[c] = pd.to_numeric(df[c], errors='coerce').fillna(0.0)
    df['Total'] = df[CATEGORIAS].sum(axis=1)
    df['mes_num'] = range(len(df))
    return df

def build_data(df):
    total_geral = float(df['Total'].sum())
    media_mensal = float(df['Total'].mean())
    recorrente_total = float(df['Faturamento Recorrente'].sum())
    pct_recorrente = round(100 * recorrente_total / total_geral, 1) if total_geral > 0 else None

    idx_max = df['Total'].idxmax()
    idx_min = df['Total'].idxmin()
    mes_maior = {'mes': df.loc[idx_max, 'Mês Ref'], 'valor': float(df.loc[idx_max, 'Total'])}
    mes_menor = {'mes': df.loc[idx_min, 'Mês Ref'], 'valor': float(df.loc[idx_min, 'Total'])}

    variacao_pct = None
    if len(df) >= 2 and df['Total'].iloc[0] > 0:
        variacao_pct = round(100 * (df['Total'].iloc[-1] - df['Total'].iloc[0]) / df['Total'].iloc[0], 1)

    meses = df['Mês Ref'].tolist()

    composicao_mensal = {
        'labels': meses,
        'series': [
            {'nome': CAT_LABEL[c], 'valores': [round(float(v), 2) for v in df[c].tolist()]}
            for c in CATEGORIAS
        ]
    }

    distribuicao_categoria = {
        'labels': [CAT_LABEL[c] for c in CATEGORIAS],
        'valores': [round(float(df[c].sum()), 2) for c in CATEGORIAS]
    }

    evolucao_total = {
        'labels': meses,
        'valores': [round(float(v), 2) for v in df['Total'].tolist()]
    }

    registros = []
    for _, r in df.iterrows():
        registros.append({
            'mes': r['Mês Ref'],
            'mes_num': int(r['mes_num']),
            'recorrente': round(float(r['Faturamento Recorrente']), 2),
            'corretiva': round(float(r['Manutenção Corretiva']), 2),
            'produtos': round(float(r['Produtos']), 2),
            'pronta_resposta': round(float(r['Pronta Resposta']), 2),
            'total': round(float(r['Total']), 2),
        })

    return {
        'kpi': {
            'total_geral': round(total_geral, 2),
            'media_mensal': round(media_mensal, 2),
            'pct_recorrente': pct_recorrente,
            'mes_maior': mes_maior,
            'mes_menor': mes_menor,
            'variacao_pct': variacao_pct,
            'meses_count': len(df),
        },
        'composicao_mensal': composicao_mensal,
        'distribuicao_categoria': distribuicao_categoria,
        'evolucao_total': evolucao_total,
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
    print(f"Faturamento total: R$ {data['kpi']['total_geral']:,.2f}")
    print(f"Media mensal: R$ {data['kpi']['media_mensal']:,.2f}")
    print(f"% Recorrente: {data['kpi']['pct_recorrente']}%")
    print(f"Maior mes: {data['kpi']['mes_maior']}")
    print(f"Menor mes: {data['kpi']['mes_menor']}")

if __name__ == '__main__':
    main()
