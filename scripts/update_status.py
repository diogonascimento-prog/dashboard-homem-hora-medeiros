# -*- coding: utf-8 -*-
"""
Gera o STATUS.md (resumo AGREGADO do dashboard) a partir do HTML já gerado por
build_dashboard.py. Só totais por mês — nenhum nome, salário individual ou dado de pessoa entra
no arquivo, de propósito (o repositório não deve conter dados pessoais da folha).
"""
import json, re, os, datetime

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
HTML = os.path.join(ROOT, 'Dashboard Horas - Medeiros.html')
OUT = os.path.join(ROOT, 'STATUS.md')

with open(HTML, 'r', encoding='utf-8') as f:
    html = f.read()
m = re.search(r'<script id="raw-data" type="application/json">(.*?)</script>', html, re.S)
payload = json.loads(m.group(1))
rows = payload['rows']

def br(n, casas=2):
    return f'{n:,.{casas}f}'.replace(',', 'X').replace('.', ',').replace('X', '.')

meses = sorted({r['mesKey'] for r in rows})
label = {r['mesKey']: r['mes'] for r in rows}

linhas = []
tot = dict(pessoas=set(), normais=0.0, extras=0.0, salario=0.0, custoExtra=0.0)
for mk in meses:
    rs = [r for r in rows if r['mesKey'] == mk]
    normais = sum(r['horasMes'] for r in rs)
    extras = sum(r['extraTotal'] for r in rs)
    salario = sum(r['salario'] for r in rs)
    cext = sum(r['custoHoraExtra'] for r in rs)
    pessoas = {r['nome'] for r in rs}
    tot['pessoas'] |= pessoas
    tot['normais'] += normais; tot['extras'] += extras
    tot['salario'] += salario; tot['custoExtra'] += cext
    pct = extras / (normais + extras) * 100 if normais + extras else 0
    linhas.append(f'| {label[mk]} | {len(pessoas)} | {br(normais, 1)} | {br(extras, 1)} | {pct:.1f}% | R$ {br(salario)} | R$ {br(cext)} | R$ {br(salario + cext)} |')

horas_total = tot['normais'] + tot['extras']
custo_total = tot['salario'] + tot['custoExtra']
agora = datetime.datetime.now().strftime('%d/%m/%Y %H:%M')

md = f"""# Status do Dashboard — {payload.get('obraNome', 'Obra')}

> Gerado automaticamente por `scripts/update_status.py` em **{agora}** (a cada execução do `Atualizar Dashboard.bat`).
> Só totais agregados — nenhum dado individual de funcionário é versionado neste repositório.

## Resumo

| Indicador | Valor |
|---|---|
| Período | {label[meses[0]]} a {label[meses[-1]]} ({len(meses)} meses) |
| Pessoas com apontamento | {len(tot['pessoas'])} |
| Linhas pessoa-mês | {len(rows)} |
| Horas realizadas | {br(horas_total, 1)} h (normais {br(tot['normais'], 1)} h + extras {br(tot['extras'], 1)} h) |
| Horas extras sobre o total | {tot['extras'] / horas_total * 100:.1f}% |
| Custo realizado | R$ {br(custo_total)} (salário R$ {br(tot['salario'])} + hora extra R$ {br(tot['custoExtra'])}) |
| Hora extra sobre o custo | {tot['custoExtra'] / custo_total * 100:.1f}% |

## Por mês

| Mês | Pessoas | Horas normais | Horas extras | % extra | Salário | Custo hora extra | Custo total |
|---|---|---|---|---|---|---|---|
""" + '\n'.join(linhas) + '\n'

with open(OUT, 'w', encoding='utf-8', newline='\n') as f:
    f.write(md)
print('STATUS.md atualizado:', OUT)
