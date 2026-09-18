import openpyxl, json, re, sys, datetime
from funcao_map import CARGO_TO_SETOR, HIST_FUNC_TO_SETOR, CARGO_IS_OBRA, CARGO_TO_HISTFUNC, SETOR_ORDER, HIST_FUNC_ALIAS, SIENGE_FUNC_ALIAS

SRC = r"C:\Users\User\OneDrive - ENGETECNICA\Documentos\Homem Hora\HHT - Medeiros - Base de dados.xlsx"
OUT = r"C:\Users\User\OneDrive - ENGETECNICA\Documentos\Homem Hora\Dashboard Horas - Medeiros.html"
# Arquivo-fonte com uma aba por mês, de onde a "Base de dados" é montada via Power Query.
# Lemos Ex50%-Ex110%/Not.Tot. direto daqui (em vez de confiar na "Base de dados") porque o
# Excel do usuário não está propagando a atualização da query de forma confiável — a "Base de
# dados" ficou com cache desatualizado pra MAIO e JUNHO/2026 mesmo depois de várias atualizações.
SRC_MENSAL = r"C:\Users\User\OneDrive - ENGETECNICA\Documentos\Homem Hora\Base de dados\HHT - 5005 MEDEIROS .xlsx"

HOURS_PER_PERSON_MONTH = 220
TODAY = datetime.date.today().isoformat()
OBRA_NOME = 'SE Medeiros Neto II'

wb = openpyxl.load_workbook(SRC, data_only=True)
ws = wb['Base de dados']

def parse_hours(v):
    # A planilha já traz Ex50%-Ex110% e Not.Tot. como horas decimais diretas
    # (ex: 4.28 = 4h17min). Só a coluna "Normais", removida pelo usuário, vinha
    # no formato de fração de dia do Excel (por isso não há *24 aqui).
    if v is None:
        return 0.0
    if isinstance(v, str):
        v = v.strip()
        if not v:
            return 0.0
        m = re.match(r'^(\d+):(\d+)$', v)
        if m:
            return int(m.group(1)) + int(m.group(2)) / 60.0
        try:
            return float(v.replace(',', '.'))
        except Exception:
            return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    return 0.0

MES_MAP = {
    'FEV.2026': ('2026-02', 'FEV/2026'),
    'MAR.2026': ('2026-03', 'MAR/2026'),
    'ABRIL 2026': ('2026-04', 'ABR/2026'),
    'MAIO 2026': ('2026-05', 'MAI/2026'),
    'JUN. 2026': ('2026-06', 'JUN/2026'),
    'JUL.2026': ('2026-07', 'JUL/2026'),
}
MESES_ABREV_PT = ['JAN', 'FEV', 'MAR', 'ABR', 'MAI', 'JUN', 'JUL', 'AGO', 'SET', 'OUT', 'NOV', 'DEZ']

def mes_key_label(ano, mes):
    return f'{ano:04d}-{mes:02d}', f'{MESES_ABREV_PT[mes - 1]}/{ano}'

# A coluna "Mês de Refência" às vezes vem como data em vez de texto (o Excel reformata a coluna
# sozinho ao recarregar a consulta) — nesse caso dá pra calcular mes_key/mes_label direto do
# ano/mês da data, sem precisar de mapa nenhum (funciona pra qualquer mês, inclusive futuro).
# Quando vem como texto e não é nenhum dos 6 nomes acima (mês novo, ainda não cadastrado em
# MES_MAP), tenta reconhecer o mês pelas 3 primeiras letras (cobre abreviação e nome por
# extenso: "AGO", "AGOSTO", "SET.2026", "SETEMBRO 2026" etc.) — assim um mês novo na planilha
# não quebra o build inteiro só porque ninguém lembrou de cadastrar aqui antes.
def parse_mes_texto(valor):
    txt = valor.strip().upper()
    m_ano = re.search(r'(\d{4})', txt)
    if not m_ano:
        return None
    ano = int(m_ano.group(1))
    letras = re.sub(r'[^A-ZÇ]', '', txt)
    prefixo = letras[:3]
    if prefixo not in MESES_ABREV_PT:
        return None
    return mes_key_label(ano, MESES_ABREV_PT.index(prefixo) + 1)

# ---- Ex50%-Ex110%/Not.Tot. lidos direto das abas mensais do arquivo-fonte ----
# Mesma conversão que a query faz (Number.Round(v * 24, 2)): os valores vêm como fração de dia
# do Excel (datetime.timedelta quando >= 24h, datetime.time quando < 24h).
def to_hours(v):
    if v is None:
        return 0.0
    if isinstance(v, datetime.timedelta):
        return round(v.total_seconds() / 3600.0, 2)
    if isinstance(v, datetime.time):
        return round((v.hour * 3600 + v.minute * 60 + v.second) / 3600.0, 2)
    if isinstance(v, (int, float)):
        return round(float(v) * 24, 2)
    return 0.0

def build_monthly_hours_lookup():
    wbm = openpyxl.load_workbook(SRC_MENSAL, data_only=True)
    lookup = {}
    for sheet_name, (mes_key, _) in MES_MAP.items():
        if sheet_name not in wbm.sheetnames:
            continue
        wsm = wbm[sheet_name]
        col_map_m = {}
        for c in range(1, wsm.max_column + 1):
            v = wsm.cell(row=1, column=c).value
            # .strip() aqui é de propósito: já achei cabeçalho com espaço sobrando
            # ("Ex60% " em vez de "Ex60%") que quebra o match exato da query no Excel.
            if isinstance(v, str) and v.strip():
                col_map_m[v.strip()] = c
        c_nome_m = col_map_m.get('NOME')
        cols_ex_m = {k: col_map_m.get(label) for k, label in
                     [('ex50', 'Ex50%'), ('ex60', 'Ex60%'), ('ex70', 'Ex70%'),
                      ('ex80', 'Ex80%'), ('ex100', 'Ex100%'), ('ex110', 'Ex110%')]}
        c_nottot_m = col_map_m.get('Not.Tot.')
        if not c_nome_m:
            continue
        for r in range(2, wsm.max_row + 1):
            nome_m = wsm.cell(row=r, column=c_nome_m).value
            if not nome_m or not str(nome_m).strip():
                continue
            nome_m = str(nome_m).strip()
            vals = {k: (to_hours(wsm.cell(row=r, column=c).value) if c else 0.0) for k, c in cols_ex_m.items()}
            vals['nottot'] = to_hours(wsm.cell(row=r, column=c_nottot_m).value) if c_nottot_m else 0.0
            lookup[(mes_key, nome_m)] = vals
    return lookup

MONTHLY_HOURS = build_monthly_hours_lookup()
print(f'Horas extras lidas direto de {len(MES_MAP)} abas mensais: {len(MONTHLY_HOURS)} registros pessoa-mês.')

# Localiza a linha de cabeçalho procurando "NOME" nas primeiras linhas
# (o usuário mantém um painel de totais acima do cabeçalho, então a posição pode mudar).
HEADER_ALIASES = {
    'arquivo': 'Nome do Arquivo', 'mes': 'Mês de Refência', 'nome': 'NOME',
    'admissao': 'ADMISSÃO', 'funcao': 'FUNÇÃO', 'local': 'Local', 'cargoagrupado': 'Cargo Agrupado',
    'salario': 'SALÁRIO', 'salhora': 'Salário por Hora',
    'horasmes': 'Horas Mês',
    'ex50': 'Ex50%', 'vex50': 'Valor Hora 50%', 'ex60': 'Ex60%', 'vex60': 'Valor Hora 60%',
    'ex70': 'Ex70%', 'vex70': 'Valor Hora 70%', 'ex80': 'Ex80%', 'vex80': 'Valor Hora 80%',
    'ex100': 'Ex100%', 'vex100': 'Valor Hora 100%', 'ex110': 'Ex110%', 'vex110': 'Valor Hora 110%',
    'custohoraextra': 'Custo Hora Extra',
    'nottot': 'Not.Tot.',
    'setor': 'Setor', 'funcaosienge': 'Função x SIENGE',
}

def find_header_row(ws, max_scan=25):
    for r in range(1, max_scan + 1):
        for c in range(1, 8):
            v = ws.cell(row=r, column=c).value
            if isinstance(v, str) and v.strip().upper().startswith('NOME'):
                return r
    raise ValueError('Cabeçalho não encontrado nas primeiras linhas de "Base de dados"')

header_row = find_header_row(ws)
col_map = {}
for c in range(1, ws.max_column + 1):
    v = ws.cell(row=header_row, column=c).value
    if isinstance(v, str) and v.strip():
        col_map[v.strip().lower()] = c

def col(key):
    target = HEADER_ALIASES[key].strip().lower()
    if target in col_map:
        return col_map[target]
    for h, idx in col_map.items():
        if h.startswith(target[:5]):
            return idx
    raise KeyError(f'Coluna "{HEADER_ALIASES[key]}" não encontrada. Cabeçalhos disponíveis: {list(col_map)}')

def col_optional(key):
    """Como col(), mas retorna None em vez de estourar erro — pra colunas novas que só existem
    depois que a query do Power Query foi atualizada com o mapeamento de Setor/Sienge."""
    try:
        return col(key)
    except KeyError:
        return None

C_NOME, C_MES, C_FUNCAO = col('nome'), col('mes'), col('funcao')
C_LOCAL, C_CARGOAGR = col('local'), col('cargoagrupado')
C_SETOR = col_optional('setor')
C_FUNCAOSIENGE = col_optional('funcaosienge')
C_ADMISSAO = col('admissao')
C_SAL, C_SALHORA, C_HMES = col('salario'), col('salhora'), col('horasmes')
EX_TYPES = ['ex50', 'ex60', 'ex70', 'ex80', 'ex100', 'ex110']
C_EX = {k: col(k) for k in EX_TYPES}
C_NOTTOT = col('nottot')
# Multiplicadores exatos da query M (Salário/Hora × multiplicador × horas do percentual).
EX_MULT = {'ex50': 1.5, 'ex60': 1.6, 'ex70': 1.7, 'ex80': 1.8, 'ex100': 2.0, 'ex110': 2.1}
NOTURNO_MULT = 1.2  # adicional noturno = 20% a mais sobre o salário/hora
print(f'Cabeçalho encontrado na linha {header_row}. Colunas: {col_map}')

rows = []
n_linhas_em_branco = 0
n_linhas_demissao = 0
for r in range(header_row + 1, ws.max_row + 1):
    nome = ws.cell(row=r, column=C_NOME).value
    if not nome or not str(nome).strip():
        n_linhas_em_branco += 1
        continue
    mes_raw = ws.cell(row=r, column=C_MES).value
    if isinstance(mes_raw, (datetime.datetime, datetime.date)):
        mes_key, mes_label = mes_key_label(mes_raw.year, mes_raw.month)
    else:
        par = MES_MAP.get(mes_raw) or (parse_mes_texto(mes_raw) if isinstance(mes_raw, str) else None)
        mes_key, mes_label = par if par else (None, None)
    if mes_key is None:
        raise ValueError(f'Mês de Refência não reconhecido na linha {r}: {mes_raw!r}')
    funcao = ws.cell(row=r, column=C_FUNCAO).value
    funcao = funcao.strip() if isinstance(funcao, str) and funcao.strip() else 'N/D'
    cargo_agrupado = ws.cell(row=r, column=C_CARGOAGR).value
    cargo_agrupado = cargo_agrupado.strip() if isinstance(cargo_agrupado, str) and cargo_agrupado.strip() else funcao
    # "Demissão" não é um cargo real (é um marcador de desligamento) — a pedido do usuário,
    # essas linhas ficam de fora do dashboard inteiro, não só de uma tabela ou outra.
    if cargo_agrupado == 'Demissão':
        n_linhas_demissao += 1
        continue
    # Setor: lê direto da coluna "Setor" da planilha (calculada na query do Power Query, mesma
    # classificação oficial do cliente) — só cai pro mapa Python (funcao_map.py) se a coluna ainda
    # não existir nessa planilha (query antiga, não atualizada).
    setor_planilha = ws.cell(row=r, column=C_SETOR).value if C_SETOR else None
    setor_planilha = setor_planilha.strip() if isinstance(setor_planilha, str) and setor_planilha.strip() else None
    setor = setor_planilha or CARGO_TO_SETOR.get(cargo_agrupado, 'Não classificado')
    # A planilha ainda escreve "Transporte/Logística" na coluna Setor (Power Query antigo) —
    # renomeado pro dashboard mostrar só "Transporte", sem precisar editar a fonte no Excel.
    if setor == 'Transporte/Logística':
        setor = 'Transporte'
    # Função x SIENGE: lida direto da coluna homônima (preenchida pessoa por pessoa via Power
    # Query, confirmada pelo cliente linha a linha) — é a fonte oficial pra bater o executado
    # contra o orçamento do Sienge no JS. "Não Previsto" é um valor real (pessoa sem linha de
    # orçamento) e não None; None só acontece se a coluna nem existir (planilha antiga).
    funcao_sienge_planilha = ws.cell(row=r, column=C_FUNCAOSIENGE).value if C_FUNCAOSIENGE else None
    funcao_sienge_planilha = funcao_sienge_planilha.strip() if isinstance(funcao_sienge_planilha, str) and funcao_sienge_planilha.strip() else None
    if funcao_sienge_planilha:
        funcao_sienge_planilha = SIENGE_FUNC_ALIAS.get(funcao_sienge_planilha, funcao_sienge_planilha)
    # Obra x Matriz vem da regra "o cargo está no histograma?", não mais da coluna "Local" da planilha.
    local = 'Obra' if CARGO_IS_OBRA.get(cargo_agrupado, False) else 'Matriz'
    salario = ws.cell(row=r, column=C_SAL).value or 0
    sal_hora = ws.cell(row=r, column=C_SALHORA).value or 0
    horas_mes = ws.cell(row=r, column=C_HMES).value or 0

    # Ex50%-Ex110%/Not.Tot.: prioriza o valor lido direto da aba mensal (fonte da verdade);
    # só cai pra "Base de dados" se essa pessoa/mês não aparecer lá por algum motivo.
    fresh = MONTHLY_HOURS.get((mes_key, str(nome).strip()))
    if fresh:
        ex = {k: fresh[k] for k in EX_TYPES}
        nottot = fresh['nottot']
    else:
        ex = {k: parse_hours(ws.cell(row=r, column=C_EX[k]).value) for k in EX_TYPES}
        nottot = parse_hours(ws.cell(row=r, column=C_NOTTOT).value)
    # Valor Hora X% e Custo Hora Extra são recalculados a partir das horas acima (não lidos da
    # "Base de dados"), pra não herdar o mesmo cache desatualizado nas colunas derivadas.
    vex = {k: round(float(sal_hora) * EX_MULT[k] * ex[k], 2) for k in EX_TYPES}
    # Adicional noturno = salário/hora x 1.2 x Not.Tot. (20% a mais, confirmado pelo cliente).
    vex_noturno = round(float(sal_hora) * NOTURNO_MULT * nottot, 2)
    custo_hora_extra = sum(vex.values()) + vex_noturno
    extra_total = sum(ex.values()) + nottot
    admissao_raw = ws.cell(row=r, column=C_ADMISSAO).value
    admissao = admissao_raw.date().isoformat() if isinstance(admissao_raw, (datetime.datetime, datetime.date)) else None
    rows.append({
        'mesKey': mes_key, 'mes': mes_label, 'nome': str(nome).strip(), 'funcao': funcao,
        'funcaoBase': cargo_agrupado, 'setor': setor, 'funcaoSienge': funcao_sienge_planilha, 'local': local, 'admissao': admissao,
        'salario': round(float(salario), 2), 'salHora': round(float(sal_hora), 4), 'horasMes': float(horas_mes),
        'ex50': round(ex['ex50'], 2), 'ex60': round(ex['ex60'], 2), 'ex70': round(ex['ex70'], 2),
        'ex80': round(ex['ex80'], 2), 'ex100': round(ex['ex100'], 2), 'ex110': round(ex['ex110'], 2),
        'vex50': round(float(vex['ex50']), 2), 'vex60': round(float(vex['ex60']), 2), 'vex70': round(float(vex['ex70']), 2),
        'vex80': round(float(vex['ex80']), 2), 'vex100': round(float(vex['ex100']), 2), 'vex110': round(float(vex['ex110']), 2),
        'custoHoraExtra': round(float(custo_hora_extra), 2), 'vexNoturno': round(float(vex_noturno), 2),
        'notTot': round(nottot, 2), 'extraTotal': round(extra_total, 2),
    })

# ---------------- HIST-MO (histograma de mão de obra) ----------------
wsh = wb['HIST-MO']

def read_hist_table(header_row, first_data_row, last_data_row):
    months = []
    for c in range(3, 40):
        v = wsh.cell(row=header_row, column=c).value
        # Para no primeiro valor que não é data — a planilha tem colunas de resumo (ex: "Total
        # de Pessoas na obra") logo depois do último mês, não só célula vazia.
        if not isinstance(v, (datetime.datetime, datetime.date)):
            break
        months.append(v.strftime('%Y-%m'))
    agg = {}
    for r in range(first_data_row, last_data_row + 1):
        func = wsh.cell(row=r, column=1).value
        if not func or not str(func).strip():
            continue
        func = str(func).strip()
        func = HIST_FUNC_ALIAS.get(func, func)
        funcao_base, setor = func, HIST_FUNC_TO_SETOR.get(func, 'Não classificado')
        for i, mk in enumerate(months):
            v = wsh.cell(row=r, column=3 + i).value
            v = int(v) if isinstance(v, (int, float)) else 0
            key = (setor, funcao_base, mk)
            agg[key] = agg.get(key, 0) + v
    return months, agg

months_p, previsto_agg = read_hist_table(53, 56, 107)
_, contratado_agg = read_hist_table(111, 114, 163)
_, liberado_agg = read_hist_table(165, 168, 217)

# A planilha tem uma coluna de mês (JUL/2027) depois do fim real do cronograma, sem nenhum
# headcount planejado em nenhuma função (mesmo padrão do mês fantasma antes do início dos
# dados) — o fim de verdade do projeto é JUN/2027. Corta essa coluna vazia do fim.
while months_p and sum(v for k, v in previsto_agg.items() if k[2] == months_p[-1]) == 0:
    removido = months_p.pop()
    previsto_agg = {k: v for k, v in previsto_agg.items() if k[2] != removido}

def agg_to_list(agg):
    return [{'setor': k[0], 'funcaoBase': k[1], 'mesKey': k[2], 'headcount': v} for k, v in agg.items() if v]

# Coluna "Salario medio por função" que o usuário calculou direto na aba HIST-MO, do lado do
# histograma de headcount — vira a referência de salário médio pra TODO o Custo Previsto do
# dashboard (KPIs, Custo por Setor/Função, Curva de Horas etc.), EXCETO a tabela do Sienge, que
# usa só o relatório do Sienge (aba "Relatório"), sem essa reserva.
def read_salario_medio_hist(header_row, first_data_row, last_data_row):
    col_map_h = {}
    for c in range(1, 30):
        v = wsh.cell(row=header_row, column=c).value
        if isinstance(v, str) and v.strip():
            col_map_h[v.strip().lower()] = c
    c_sal = next((c for k, c in col_map_h.items() if k.startswith('salario medio') or k.startswith('salário medio') or k.startswith('salário médio')), None)
    if not c_sal:
        return {}
    out = {}
    for r in range(first_data_row, last_data_row + 1):
        func = wsh.cell(row=r, column=1).value
        if not func or not str(func).strip():
            continue
        func = str(func).strip()
        func = HIST_FUNC_ALIAS.get(func, func)
        sal = wsh.cell(row=r, column=c_sal).value
        if isinstance(sal, (int, float)) and sal > 0:
            out[func] = float(sal)
    return out

SALARIO_MEDIO_HIST = read_salario_medio_hist(53, 56, 107)
print(f'Salário médio por função lido da aba HIST-MO: {len(SALARIO_MEDIO_HIST)} funções.')

# ---------------- Relatório do Sienge (orçado por função, aba "Relatório") ----------------
# Essa aba é o BOQ exportado do Sienge: Código/Descrição/Un./Quantidade/Preço unitário/Preço
# total/Preço total reajustado, uma linha por função. Não tem coluna de "já gasto" — só o valor
# orçado. O "Saldo a Gastar" é calculado no JS (orçado − executado, usando os lançamentos da
# própria base), não vem pronto da planilha.
def read_sienge_orcado():
    if 'Relatório' not in wb.sheetnames:
        return []
    wsr = wb['Relatório']
    header_row = None
    for r in range(1, 15):
        for c in range(1, wsr.max_column + 1):
            v = wsr.cell(row=r, column=c).value
            if isinstance(v, str) and v.strip().lower().startswith('descri'):
                header_row = r
                break
        if header_row:
            break
    if not header_row:
        return []
    col_map_s = {}
    for c in range(1, wsr.max_column + 1):
        v = wsr.cell(row=header_row, column=c).value
        if isinstance(v, str) and v.strip():
            col_map_s[v.strip().lower()] = c
    c_desc = col_map_s.get('descrição') or col_map_s.get('descricao')
    c_orcado = col_map_s.get('preço total reajustado') or col_map_s.get('preco total reajustado')
    c_horas = col_map_s.get('quantidade')
    if not c_desc or not c_orcado:
        return []
    out = []
    for r in range(header_row + 1, wsr.max_row + 1):
        desc = wsr.cell(row=r, column=c_desc).value
        if not desc or not str(desc).strip():
            continue
        orcado = wsr.cell(row=r, column=c_orcado).value
        if not isinstance(orcado, (int, float)):
            continue
        horas = wsr.cell(row=r, column=c_horas).value if c_horas else None
        horas = round(float(horas), 2) if isinstance(horas, (int, float)) else 0.0
        out.append({'funcao': str(desc).strip(), 'orcado': round(float(orcado), 2), 'orcadoHoras': horas})
    return out

SIENGE_ORCADO = read_sienge_orcado()
print(f'Relatório do Sienge: {len(SIENGE_ORCADO)} funções orçadas lidas da aba "Relatório".')

# ---------- Avanço Físico (Curva S do planejamento) — lido de um arquivo externo separado que o
# Planejamento mantém (curvas semanais com 4 revisões de linha de base: BL0-BL3, mais uma linha
# de Tendência e uma de Real, cada uma com uma sub-curva por disciplina: GER/ENG/SUP/CeM/PEM/PEL/
# PCV/OBR/MNT/CMS/PO/MANF/MOS). Usamos a linha "CeM" — é essa a coluna por trás do card
# "Construção" da aba "CURVAS GERENCIAIS" da própria planilha do Planejamento (confirmado
# batendo os números: BL0-CeM e REAL-CeM de 23/08/2026 = 50,36%/44,81%, exatamente os valores que
# o cliente informou como "Construção" — 2026-08-27). Não confundir com "OBR", que na mesma aba
# de resumo aparece como "Obra Civil" — uma disciplina mais estreita, não o que o cliente pediu.
# "CeM"/Construção também bate melhor com a mão de obra em campo que este dashboard mede do que
# "GER" (projeto inteiro, inclui engenharia/suprimentos, fases sem gente na Obra lançando hora).
# E a linha de base BL0 como Planejado: é a única cuja data de conclusão bate com o fim do
# histograma (JUN/2027) — BL1/BL2 são revisões antigas já defasadas (BL2 "termina" em ABR/2026,
# muito antes do avanço real de ~62% em AGO/2026) e BL3 está vazia nesta planilha. Arquivo é
# opcional: se não existir/mudar de lugar, o dashboard simplesmente não mostra essa curva, sem
# quebrar o resto do build.
SRC_CURVAS = r"C:\Users\User\Downloads\SE MND - Curvas S - 23.08.2026.xlsm"

def ler_avanco_fisico():
    import warnings
    try:
        with warnings.catch_warnings():
            warnings.simplefilter('ignore')
            wb_c = openpyxl.load_workbook(SRC_CURVAS, data_only=True)
        ws_c = wb_c['BASELINES']
    except Exception as e:
        print(f'Aviso: não consegui ler a curva de avanço físico ({SRC_CURVAS}): {e}. Gráfico de avanço físico ficará vazio.')
        return {'planejado': [], 'realizado': []}

    def row_vals(r, maxcol=286):
        return [ws_c.cell(row=r, column=c).value for c in range(2, maxcol + 1)]

    hoje = datetime.date.today()

    def monthly_series(date_row, val_row):
        datas = row_vals(date_row)
        vals = row_vals(val_row)
        by_month = {}
        for d, v in zip(datas, vals):
            if d is None or not isinstance(v, (int, float)):
                continue
            mk = f'{d.year:04d}-{d.month:02d}'
            # Cumulativo: fica com o ponto mais recente de cada mês (valor de fim de mês) — MAS,
            # pro mês corrente (ainda em andamento), só considera semanas até hoje. A curva de
            # linha de base (BL0/Planejado) já vem preenchida até o fim do projeto inteiro, então
            # sem essa checagem o mês corrente pegava uma semana FUTURA daquele mesmo mês (ex:
            # 30/08 em vez de 23/08 com "hoje" = 27/08) — inflava o Previsto além do que já
            # deveria ter acontecido até agora. Mês inteiramente no futuro (nenhuma semana <=
            # hoje) continua pegando a última semana dele normalmente — sem essa curva não dá pra
            # desenhar o Previsto dos meses que ainda vêm.
            passou = d.date() <= hoje
            atual = by_month.get(mk)
            if atual is None:
                by_month[mk] = (d, v, passou)
            else:
                d_atual, v_atual, passou_atual = atual
                if passou and not passou_atual:
                    by_month[mk] = (d, v, True)
                elif passou == passou_atual and d > d_atual:
                    by_month[mk] = (d, v, passou)
        by_month = {mk: (d, v) for mk, (d, v, _) in by_month.items()}
        return [{'mesKey': mk, 'pct': round(v * 100, 2)} for mk, (d, v) in sorted(by_month.items())]

    planejado = monthly_series(1, 5)    # datas0 / BL0-CeM
    realizado = monthly_series(71, 75)  # datasR / REAL-CeM
    print(f'Avanço físico: {len(planejado)} meses de Planejado (BL0-CeM), {len(realizado)} meses de Realizado (REAL-CeM).')
    return {'planejado': planejado, 'realizado': realizado}

AVANCO_FISICO = ler_avanco_fisico()

payload = {
    'rows': rows,
    'histPrevisto': agg_to_list(previsto_agg),
    'histContratado': agg_to_list(contratado_agg),
    'histLiberado': agg_to_list(liberado_agg),
    'projectMonths': months_p,
    'hoursPerPersonMonth': HOURS_PER_PERSON_MONTH,
    'today': TODAY,
    'setorOrder': SETOR_ORDER,
    'obraNome': OBRA_NOME,
    'cargoToHistFunc': CARGO_TO_HISTFUNC,
    'siengeOrcado': SIENGE_ORCADO,
    'salarioMedioHist': SALARIO_MEDIO_HIST,
    'avancoFisico': AVANCO_FISICO,
}

data_json = json.dumps(payload, ensure_ascii=False)
print('rows:', len(rows))

HTML_TEMPLATE = r"""<!doctype html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Dashboard de Projetos - Medeiros</title>
<style>
/* Tipografia (a pedido do cliente, 2026-09-15: mesma tipografia/tamanhos/estilo de um
   layout de referência, mantendo só a nossa paleta de cores — ver comentário no :root). */
@import url("https://fonts.googleapis.com/css2?family=Urbanist:wght@300..800&family=Plus+Jakarta+Sans:wght@400..800&family=Space+Grotesk:wght@400..700&display=swap");

:root{
  --bg:#f4f6f8; --card:#ffffff; --sunken:#f7f8fa; --text:#1a2332; --muted:#5b6b7c; --border:#e2e8ef;
  --green:#86efac; --green-dark:#22c55e; --green-text:#166534;
  --red:#ef4444; --red-dark:#b91c1c; --red-text:#fff;
  --ex50:#fecaca; --ex60:#fca5a5; --ex70:#f87171; --ex80:#ef4444; --ex100:#dc2626; --ex110:#7f1d1d; --notTot:#a78bfa;
  --accent:#E8622C; --amber:#f59e0b; --previsto:#2563EB; --executado:#EF4444;

  /* ---- Tipografia: Urbanist (títulos) · Plus Jakarta Sans (UI) · Space Grotesk (números) ---- */
  --font-display:"Urbanist",-apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif;
  --font-ui:"Plus Jakarta Sans",-apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif;
  --font-figure:"Space Grotesk","Plus Jakarta Sans",ui-monospace,sans-serif;
  --ls-heading:-0.01em; --ls-label:0.01em; --ls-eyebrow:0.08em;

  /* ---- Raios ---- */
  --r-xs:6px; --r-sm:8px; --r-md:10px; --r-lg:14px; --r-xl:18px; --r-2xl:24px; --r-pill:999px;

  /* ---- Elevação (tingida da nossa tinta --text, não de verde) ---- */
  --sh-xs:0 1px 2px rgba(26,35,50,.06);
  --sh-sm:0 1px 2px rgba(26,35,50,.06), 0 2px 6px -2px rgba(26,35,50,.07);
  --sh-card:0 1px 2px rgba(26,35,50,.05), 0 10px 26px -18px rgba(26,35,50,.16);
  --sh-raised:0 1px 3px rgba(26,35,50,.08), 0 14px 32px -16px rgba(26,35,50,.18);
  --sh-pop:0 8px 18px -6px rgba(26,35,50,.16), 0 24px 56px -24px rgba(26,35,50,.24);
  --sh-accent:0 8px 20px -10px rgba(232,98,44,.5);
  --ring-focus:0 0 0 3px rgba(232,98,44,.35);
  --shadow:var(--sh-card);

  /* ---- Movimento ---- */
  --dur-fast:140ms; --dur-base:200ms; --dur-slow:320ms; --dur-slower:520ms;
  --ease-out:cubic-bezier(.16,1,.3,1);
  --press-scale:.97;

  /* ---- Layout ---- */
  --topbar-h:60px; --sidebar-w:232px;
}
@media (prefers-color-scheme: dark){
  :root:not([data-theme="light"]){ --bg:#0f1521; --card:#161d2b; --sunken:rgba(255,255,255,.04); --text:#e7ecf3; --muted:#93a1b5; --border:#263042;
  --green:#4ade80; --green-dark:#22c55e; --green-text:#052e13;
  --red:#f87171; --red-dark:#ef4444; --red-text:#1a0505;
  --ex50:#7f1d1d; --ex60:#991b1b; --ex70:#b91c1c; --ex80:#dc2626; --ex100:#ef4444; --ex110:#fca5a5; --notTot:#7c3aed;
  --accent:#F2824F; --amber:#fbbf24; --previsto:#5B8DEF; --executado:#f87171;
  --sh-xs:0 1px 2px rgba(0,0,0,.35); --sh-sm:0 1px 2px rgba(0,0,0,.35);
  --sh-card:0 1px 2px rgba(0,0,0,.35), 0 10px 26px -18px rgba(0,0,0,.6);
  --sh-raised:0 14px 32px -16px rgba(0,0,0,.65);
  --sh-pop:0 8px 18px -6px rgba(0,0,0,.5), 0 24px 56px -24px rgba(0,0,0,.7);
  --shadow:var(--sh-card);
  }
}
:root[data-theme="dark"]{
  --bg:#0f1521; --card:#161d2b; --sunken:rgba(255,255,255,.04); --text:#e7ecf3; --muted:#93a1b5; --border:#263042;
  --green:#4ade80; --green-dark:#22c55e; --green-text:#052e13;
  --red:#f87171; --red-dark:#ef4444; --red-text:#1a0505;
  --ex50:#7f1d1d; --ex60:#991b1b; --ex70:#b91c1c; --ex80:#dc2626; --ex100:#ef4444; --ex110:#fca5a5; --notTot:#7c3aed;
  --accent:#F2824F; --amber:#fbbf24; --previsto:#5B8DEF; --executado:#f87171;
  --sh-xs:0 1px 2px rgba(0,0,0,.35); --sh-sm:0 1px 2px rgba(0,0,0,.35);
  --sh-card:0 1px 2px rgba(0,0,0,.35), 0 10px 26px -18px rgba(0,0,0,.6);
  --sh-raised:0 14px 32px -16px rgba(0,0,0,.65);
  --sh-pop:0 8px 18px -6px rgba(0,0,0,.5), 0 24px 56px -24px rgba(0,0,0,.7);
  --shadow:var(--sh-card);
}
:root[data-theme="light"]{
  --bg:#f4f6f8; --card:#ffffff; --sunken:#f7f8fa; --text:#1a2332; --muted:#5b6b7c; --border:#e2e8ef;
  --green:#86efac; --green-dark:#22c55e; --green-text:#166534;
  --red:#ef4444; --red-dark:#b91c1c; --red-text:#fff;
  --ex50:#fecaca; --ex60:#fca5a5; --ex70:#f87171; --ex80:#ef4444; --ex100:#dc2626; --ex110:#7f1d1d; --notTot:#a78bfa;
  --accent:#E8622C; --amber:#f59e0b; --previsto:#2563EB; --executado:#EF4444;
  --shadow:var(--sh-card);
}
*{box-sizing:border-box;}
body{margin:0;background:var(--bg);color:var(--text);font-family:var(--font-ui);font-size:14px;line-height:1.55;-webkit-font-smoothing:antialiased;text-rendering:optimizeLegibility;}
::selection{background:color-mix(in srgb, var(--accent) 30%, var(--card));color:var(--text);}
:focus-visible{outline:none;box-shadow:var(--ring-focus);}
.wrap{max-width:1320px;margin:0 auto;width:100%;padding:18px 24px 44px;}
body.bi-active .wrap{padding-bottom:18px;}
.brand-mark{width:36px;height:36px;flex-shrink:0;}
.subtitle{color:var(--muted);font-size:12.5px;margin:0;}
.controls{display:flex;gap:8px;align-items:center;flex-wrap:wrap;}
select, .btn{font-family:var(--font-ui);font-size:12.5px;font-weight:600;padding:8px 14px;border-radius:var(--r-pill);border:1px solid var(--border);background:var(--card);color:var(--text);cursor:pointer;box-shadow:var(--sh-xs);transition:background-color var(--dur-fast) var(--ease-out),border-color var(--dur-fast) var(--ease-out),color var(--dur-fast) var(--ease-out),box-shadow var(--dur-base) var(--ease-out),transform var(--dur-base) var(--ease-out);}
select{padding-right:10px;}
select:hover, .btn:hover{background:var(--sunken);box-shadow:var(--sh-sm);}
.btn{display:inline-flex;align-items:center;gap:6px;letter-spacing:var(--ls-label);}
.btn:active{transform:scale(var(--press-scale));}
.btn.active{background:var(--accent);border-color:var(--accent);color:#fff;box-shadow:var(--sh-accent);}
.search{font-family:var(--font-ui);font-size:12.5px;padding:8px 14px;border-radius:var(--r-lg);border:1px solid var(--border);background:var(--card);color:var(--text);width:210px;}
.search::placeholder{color:var(--muted);}
.search:focus{outline:none;box-shadow:var(--ring-focus);border-color:var(--accent);}

/* ============================================================
   APP SHELL — sidebar fixa (esquerda) + área principal (topbar + conteúdo).
   Layout inspirado em software SaaS convencional (sidebar com ícones/seções),
   a pedido do cliente (2026-09-15) — antes era um menu horizontal no topo.
   ============================================================ */
.app{display:flex;min-height:100vh;}
.sidebar{width:232px;flex:none;position:sticky;top:0;align-self:flex-start;height:100vh;display:flex;flex-direction:column;gap:14px;padding:16px 12px;border-right:1px solid var(--border);background:var(--bg);z-index:35;}
.sidebar-brand{display:flex;align-items:center;gap:10px;padding:4px 8px 2px;}
.sidebar-brand .brand-mark{width:30px;height:30px;}
.sidebar-brand .sb-word{min-width:0;}
.sidebar-brand .sb-name{display:block;font-family:var(--font-display);font-size:15px;font-weight:700;letter-spacing:var(--ls-heading);color:var(--text);line-height:1.2;}
.sidebar-brand .sb-kicker{display:block;font-size:10px;letter-spacing:var(--ls-eyebrow);text-transform:uppercase;color:var(--muted);margin-top:2px;}
.sidebar-nav{display:flex;flex-direction:column;gap:14px;flex:1;min-height:0;overflow-y:auto;}
.nav-group{display:flex;flex-direction:column;gap:2px;}
.nav-group-label{padding:4px 10px 6px;font-size:10px;font-weight:700;letter-spacing:var(--ls-eyebrow);text-transform:uppercase;color:var(--muted);}
.pagenav-btn{display:flex;align-items:center;gap:10px;width:100%;height:38px;padding:0 10px;border:none;border-radius:var(--r-md);background:none;color:var(--muted);font-family:var(--font-ui);font-size:12.5px;font-weight:600;text-align:left;cursor:pointer;white-space:nowrap;transition:background-color var(--dur-fast) var(--ease-out),color var(--dur-fast) var(--ease-out);}
.pagenav-btn svg{width:17px;height:17px;flex:none;stroke-width:1.8;}
.pagenav-btn .nav-text{flex:1;min-width:0;overflow:hidden;text-overflow:ellipsis;}
.pagenav-btn:not(.active):hover{background:color-mix(in srgb, var(--accent) 8%, transparent);color:var(--text);}
.pagenav-btn.active{color:#fff;background:var(--accent);font-weight:700;}
.sidebar-foot{flex:none;display:flex;flex-direction:column;gap:8px;}
.sb-context{background:var(--card);border:1px solid var(--border);border-radius:var(--r-lg);padding:11px 12px;box-shadow:var(--sh-xs);}
.sb-context .sb-ctx-label{font-size:9.5px;font-weight:700;letter-spacing:var(--ls-eyebrow);text-transform:uppercase;color:var(--muted);margin-bottom:5px;}
.sb-context .subtitle{font-size:11px;line-height:1.5;}
.sb-actions{display:flex;align-items:center;justify-content:space-between;gap:8px;padding:0 4px;}
.sb-actions .sb-hint{font-size:10px;color:var(--muted);}
.icon-btn{width:34px;height:34px;flex:none;display:inline-flex;align-items:center;justify-content:center;border-radius:var(--r-pill);border:1px solid var(--border);background:var(--card);color:var(--muted);cursor:pointer;box-shadow:var(--sh-xs);transition:background-color var(--dur-fast) var(--ease-out),color var(--dur-fast) var(--ease-out),transform var(--dur-base) var(--ease-out);}
.icon-btn svg{width:16px;height:16px;stroke-width:1.8;}
.icon-btn:hover{background:color-mix(in srgb, var(--accent) 8%, transparent);color:var(--text);}
.icon-btn:active{transform:scale(var(--press-scale));}
.theme-toggle .ico-moon{display:none;}
:root[data-theme="dark"] .theme-toggle .ico-moon{display:block;}
:root[data-theme="dark"] .theme-toggle .ico-sun{display:none;}
@media (prefers-color-scheme: dark){
  :root:not([data-theme="light"]) .theme-toggle .ico-moon{display:block;}
  :root:not([data-theme="light"]) .theme-toggle .ico-sun{display:none;}
}
.main{flex:1;min-width:0;display:flex;flex-direction:column;}
.topbar{position:sticky;top:0;z-index:30;display:flex;align-items:center;gap:16px;min-height:var(--topbar-h,60px);padding:11px 24px;background:color-mix(in srgb, var(--bg) 85%, transparent);backdrop-filter:saturate(140%) blur(14px);-webkit-backdrop-filter:saturate(140%) blur(14px);border-bottom:1px solid var(--border);}
.topbar-title{min-width:0;flex:1;}
.topbar-title h1{font-family:var(--font-display);font-size:19px;font-weight:600;letter-spacing:var(--ls-heading);color:var(--text);margin:0;line-height:1.2;}
.topbar-title .page-sub{font-size:12px;color:var(--muted);margin:2px 0 0;line-height:1.4;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;}
.topbar .controls{flex:none;}
.menu-btn{display:none;}
.sidebar-scrim{position:fixed;inset:0;background:rgba(15,21,33,.42);opacity:0;pointer-events:none;transition:opacity var(--dur-base) var(--ease-out);z-index:34;}
@media (max-width:1024px){
  .sidebar{position:fixed;left:0;top:0;height:100vh;width:260px;background:var(--bg);transform:translateX(-100%);transition:transform var(--dur-slow) var(--ease-out);box-shadow:var(--sh-pop);}
  body.nav-open .sidebar{transform:translateX(0);}
  body.nav-open .sidebar-scrim{opacity:1;pointer-events:auto;}
  .menu-btn{display:inline-flex;}
  .topbar{padding:11px 16px;}
  .wrap{padding:16px 16px 40px;}
}
@media (max-width:760px){
  .topbar{flex-wrap:wrap;align-items:flex-start;}
  .topbar .controls{width:100%;}
  .topbar-title .page-sub{white-space:normal;}
}
.page{display:none;}
.page.active{display:block;}
.kpis{display:grid;grid-template-columns:repeat(auto-fit,minmax(200px,1fr));gap:14px;margin-bottom:18px;}
@media (min-width:1100px){ #page-inicio .kpis{grid-template-columns:repeat(3,1fr);} }
@media (max-width:620px){ .kpis{grid-template-columns:1fr;} }
.kpi{display:flex;flex-direction:column;background:var(--card);border:1px solid var(--border);border-radius:var(--r-xl);padding:16px 18px;box-shadow:var(--sh-card);position:relative;overflow:hidden;}
.kpi .label{font-size:11.5px;color:var(--muted);font-weight:600;letter-spacing:var(--ls-label);margin-bottom:8px;line-height:1.4;}
.kpi .value{font-family:var(--font-figure);font-size:21px;font-weight:700;letter-spacing:-.02em;line-height:1.2;}
@media (min-width:1100px){ #page-inicio .kpi .value{font-size:26px;} }
.kpi .sub{font-size:11.5px;color:var(--muted);margin-top:auto;padding-top:8px;}
.kpi-breakdown{display:flex;flex-direction:column;gap:6px;margin-top:auto;padding-top:10px;}
.kpi-breakdown .kpi-bd-item{display:flex;align-items:center;gap:7px;flex-wrap:wrap;font-family:var(--font-ui);font-size:14.5px;color:var(--text);background:var(--sunken);border:1px solid transparent;border-radius:var(--r-md);padding:5px 11px;line-height:1.3;}
.kpi-breakdown .kpi-bd-item .dot{width:9px;height:9px;border-radius:var(--r-pill);flex-shrink:0;}
.kpi-breakdown .kpi-bd-item b{font-family:var(--font-figure);font-weight:700;}
#page-bi .kpi-breakdown{gap:4px;margin-top:auto;padding-top:6px;}
#page-bi .kpi-breakdown .kpi-bd-item{font-size:11.5px;padding:3px 7px;gap:5px;border-radius:var(--r-xs);}
#page-bi .kpi-breakdown .kpi-bd-item .dot{width:7px;height:7px;}
.card{background:var(--card);border:1px solid var(--border);border-radius:var(--r-xl);padding:18px 20px;box-shadow:var(--sh-card);margin-bottom:14px;}
.card-head{display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:10px;margin-bottom:14px;}
.card-title{font-family:var(--font-ui);font-size:14px;font-weight:700;margin:0;letter-spacing:var(--ls-heading);color:var(--text);}
.card-title .muted{color:var(--muted);font-weight:500;font-size:12px;}
.legend{display:flex;gap:14px;flex-wrap:wrap;font-size:11.5px;color:var(--muted);}
.legend span.dot{display:inline-block;width:9px;height:9px;border-radius:3px;margin-right:5px;vertical-align:-1px;}
#chartWrap, #curveWrap, #finChartWrap, #projChartWrap, #projDiffChartWrap, #decisaoFinChartWrap, #decisaoProjChartWrap, #decisaoProjDiffChartWrap, #homeHistWrap, #homeCurvaSWrap{width:100%;overflow-x:auto;}
svg#chart, svg#curveChart, svg#finChart, svg#projChart, svg#projDiffChart, svg#decisaoFinChart, svg#decisaoProjChart, svg#decisaoProjDiffChart, svg#homeHistChart, svg#homeCurvaSChart{display:block;}
.bar{cursor:pointer;transition:opacity var(--dur-fast) var(--ease-out);}
.bar:hover{opacity:.82;}
.axis text{fill:var(--muted);font-family:var(--font-figure);font-size:10.5px;}
.axis line, .axis path{stroke:var(--border);}
.gridline{stroke:var(--border);stroke-dasharray:3,3;}
.tooltip{position:fixed;pointer-events:none;background:var(--text);color:var(--bg);font-family:var(--font-ui);font-size:12px;padding:9px 11px;border-radius:var(--r-sm);box-shadow:var(--sh-pop);opacity:0;transition:opacity .1s;z-index:50;max-width:250px;line-height:1.5;}
table{width:100%;border-collapse:collapse;font-size:12.5px;}
thead th{text-align:left;padding:9px 10px;background:var(--sunken);border-bottom:1px solid var(--border);color:var(--muted);font-weight:600;font-size:11px;letter-spacing:var(--ls-label);cursor:pointer;white-space:nowrap;user-select:none;transition:color var(--dur-fast) var(--ease-out);}
thead th:first-child{border-top-left-radius:var(--r-sm);border-bottom-left-radius:var(--r-sm);}
thead th:last-child{border-top-right-radius:var(--r-sm);border-bottom-right-radius:var(--r-sm);}
thead th:hover{color:var(--text);}
thead th.num, td.num{text-align:right;}
td.num{font-family:var(--font-figure);font-weight:500;}
tbody td{padding:9px 10px;border-bottom:1px solid var(--border);}
tbody tr{transition:background-color var(--dur-fast) var(--ease-out);}
tbody tr:hover{background:color-mix(in srgb, var(--accent) 6%, transparent);}
tbody tr.row-total{font-weight:700;background:color-mix(in srgb, var(--accent) 6%, transparent);}
tbody tr.row-total td{border-top:1px solid var(--border);border-bottom:none;}
tbody tr.row-total:hover{background:color-mix(in srgb, var(--accent) 6%, transparent);}
.sort-arrow{font-size:10px;opacity:.6;margin-left:3px;}
.rank{display:inline-flex;align-items:center;justify-content:center;width:22px;height:22px;border-radius:var(--r-pill);font-family:var(--font-figure);font-size:10.5px;font-weight:700;background:var(--sunken);color:var(--text);}
.rank.top1{background:#f59e0b;color:#3a2400;}
.rank.top2{background:#cbd5e1;color:#1e293b;}
.rank.top3{background:#d97706aa;color:#2c1400;}
.detail-btn{font-size:11px;padding:5px 11px;border-radius:var(--r-pill);border:1px solid var(--border);background:var(--card);color:var(--accent);cursor:pointer;font-weight:600;transition:background-color var(--dur-fast) var(--ease-out);}
.detail-btn:hover{background:var(--sunken);}
.detail-row td{background:var(--sunken);padding:14px 16px;}
.chip{display:inline-flex;align-items:center;gap:6px;font-size:11.5px;padding:6px 14px;border-radius:var(--r-pill);background:var(--sunken);border:1px solid transparent;margin:3px 6px 3px 0;}
.chip .dot{width:8px;height:8px;border-radius:var(--r-pill);}
.name-cell{font-weight:600;}
.tag{font-size:11px;color:var(--muted);}
.badge{display:inline-flex;align-items:center;font-size:10.5px;font-weight:600;padding:3px 10px;border-radius:var(--r-pill);letter-spacing:var(--ls-label);}
.badge.obra{background:color-mix(in srgb, var(--accent) 15%, transparent);color:var(--accent);}
.badge.matriz{background:color-mix(in srgb, var(--notTot) 18%, transparent);color:var(--notTot);}
.pill{font-family:var(--font-figure);font-size:11px;font-weight:700;padding:3px 9px;border-radius:var(--r-pill);background:var(--red);color:var(--red-text);}
.progress-track{width:100%;height:8px;border-radius:var(--r-pill);background:var(--border);overflow:hidden;}
.progress-fill{height:100%;border-radius:var(--r-pill);background:var(--accent);transition:width var(--dur-slower) var(--ease-out);}
.footnote{font-size:11px;color:var(--muted);line-height:1.6;margin-top:4px;}
.nav-cards{display:grid;grid-template-columns:repeat(auto-fit,minmax(220px,1fr));gap:14px;margin-bottom:18px;}
.nav-card{background:var(--card);border:1px solid var(--border);border-radius:var(--r-xl);padding:20px;box-shadow:var(--sh-card);cursor:pointer;transition:box-shadow var(--dur-base) var(--ease-out),background-color var(--dur-fast) var(--ease-out);text-align:left;}
.nav-card:hover{background:var(--sunken);box-shadow:var(--sh-raised);}
.nav-card:active{transform:scale(var(--press-scale));}
.nav-card .nc-icon{width:34px;height:34px;border-radius:var(--r-md);display:flex;align-items:center;justify-content:center;font-size:15px;font-weight:800;margin-bottom:12px;}
.nav-card .nc-title{font-family:var(--font-ui);font-size:14px;font-weight:700;margin-bottom:4px;color:var(--text);letter-spacing:var(--ls-heading);}
.nav-card .nc-desc{font-size:12px;color:var(--muted);line-height:1.5;}
.section-title{font-family:var(--font-display);font-size:20px;font-weight:600;margin:0 0 4px;letter-spacing:var(--ls-heading);color:var(--text);}
.section-sub{font-size:12.5px;color:var(--muted);margin:0 0 16px;}
.data-note{font-size:11.5px;color:var(--muted);background:var(--sunken);border:1px dashed var(--border);border-radius:var(--r-lg);padding:11px 13px;margin-top:6px;line-height:1.6;}
::-webkit-scrollbar{height:8px;width:8px;}
::-webkit-scrollbar-thumb{background:var(--border);border-radius:var(--r-pill);}
.tbl-toggle{background:var(--card);border:1px solid var(--border);border-radius:var(--r-xl);box-shadow:var(--sh-card);margin-bottom:14px;overflow:hidden;}
.tbl-toggle summary{list-style:none;cursor:pointer;padding:15px 20px;display:flex;align-items:center;justify-content:space-between;gap:10px;flex-wrap:wrap;transition:background-color var(--dur-fast) var(--ease-out);}
.tbl-toggle summary::-webkit-details-marker{display:none;}
.tbl-toggle summary .tt-title{font-family:var(--font-ui);font-size:14px;font-weight:700;letter-spacing:var(--ls-heading);color:var(--text);}
.tbl-toggle summary .tt-chevron{width:9px;height:9px;border-right:2px solid var(--muted);border-bottom:2px solid var(--muted);transform:rotate(-45deg);transition:transform var(--dur-base) var(--ease-out);flex-shrink:0;}
.tbl-toggle[open] summary .tt-chevron{transform:rotate(45deg);}
.tbl-toggle summary .tt-actions{display:flex;gap:8px;align-items:center;flex-wrap:wrap;margin-left:auto;}
.tbl-toggle .tt-body{padding:0 20px 20px;}
.tbl-toggle summary:hover{background:var(--sunken);}
.tbl-toggle summary:hover .tt-chevron{border-color:var(--text);}
.kpi-signal{font-weight:700;}
.section-block-title{font-size:11px;font-weight:600;letter-spacing:var(--ls-eyebrow);color:var(--muted);text-transform:uppercase;margin:28px 0 12px;padding-top:16px;border-top:1px solid var(--border);}
#page-bi{font-size:12.5px;}
#page-bi .section-title{margin:0 0 6px;}
.bi-layout{display:grid;grid-template-columns:168px 1fr;gap:12px;align-items:start;}
.bi-sidebar{background:var(--card);border:1px solid var(--border);border-radius:var(--r-lg);padding:12px;box-shadow:var(--sh-card);position:sticky;top:8px;}
.bi-filter-title{font-size:10.5px;font-weight:600;text-transform:uppercase;letter-spacing:var(--ls-eyebrow);color:var(--muted);margin-bottom:8px;}
.bi-filter-label{display:block;font-size:10px;font-weight:600;color:var(--muted);margin:10px 0 4px;letter-spacing:var(--ls-label);}
.bi-filter-label:first-of-type{margin-top:0;}
.bi-filter-select{width:100%;font-family:var(--font-ui);font-size:11px;padding:7px 8px;border-radius:var(--r-md);border:1px solid var(--border);background:var(--sunken);color:var(--text);}
/* Filtro multi-seleção (dropdown com checkboxes) — Mês/Setor (topo + Visão BI) */
.msf{position:relative;}
.msf-btn{font-family:var(--font-ui);font-size:12.5px;font-weight:600;padding:8px 14px;border-radius:var(--r-pill);border:1px solid var(--border);background:var(--card);color:var(--text);cursor:pointer;white-space:nowrap;box-shadow:var(--sh-xs);transition:background-color var(--dur-fast) var(--ease-out),border-color var(--dur-fast) var(--ease-out);}
.msf-btn::after{content:'⌄';margin-left:8px;font-size:11px;line-height:1;color:var(--muted);}
.msf-btn:hover{background:var(--sunken);border-color:var(--accent);}
.msf-panel{position:absolute;top:calc(100% + 6px);left:0;z-index:40;min-width:210px;max-height:280px;overflow-y:auto;background:var(--card);border:1px solid var(--border);border-radius:var(--r-lg);box-shadow:var(--sh-pop);padding:8px;}
.msf-panel-actions{display:flex;gap:6px;margin-bottom:6px;padding-bottom:6px;border-bottom:1px solid var(--border);}
.msf-panel-actions button{flex:1;font-family:var(--font-ui);font-size:10.5px;font-weight:600;padding:5px 8px;border-radius:var(--r-pill);border:1px solid var(--border);background:var(--card);color:var(--text);cursor:pointer;transition:background-color var(--dur-fast) var(--ease-out);}
.msf-panel-actions button:hover{background:color-mix(in srgb, var(--accent) 10%, transparent);border-color:var(--accent);}
.msf-option{display:flex;align-items:center;gap:8px;font-size:12px;padding:6px 6px;border-radius:var(--r-md);cursor:pointer;transition:background-color var(--dur-fast) var(--ease-out);}
.msf-option:hover{background:var(--sunken);}
.msf-option input{cursor:pointer;accent-color:var(--accent);}
.msf-empty{font-size:11.5px;color:var(--muted);padding:4px;}
.msf-block{width:100%;}
.msf-block .msf-btn{width:100%;text-align:left;font-size:11px;padding:7px 10px;display:flex;justify-content:space-between;align-items:center;}
.msf-block .msf-panel{left:0;right:0;min-width:0;width:100%;}
.bi-sidebar .btn{font-size:10.5px;padding:6px 12px;}
@media (max-width:900px){ .bi-layout{grid-template-columns:1fr;} .bi-sidebar{position:static;display:grid;grid-template-columns:repeat(4,1fr);gap:8px;align-items:end;} .bi-filter-title{grid-column:1/-1;margin-bottom:0;} .bi-filter-label{margin:0 0 3px;} }
#page-bi .kpis{margin-bottom:10px;gap:10px;grid-template-columns:repeat(auto-fit,minmax(168px,1fr));}
#page-bi .kpi{padding:9px 12px;border-radius:var(--r-lg);box-shadow:var(--sh-xs);}
#page-bi .kpi .label{font-size:9.5px;margin-bottom:3px;}
#page-bi .kpi .value{font-size:17px;}
#page-bi .kpi .sub{font-size:9.5px;margin-top:2px;}
.bi-grid{display:grid;grid-template-columns:repeat(3,1fr);gap:10px;}
@media (max-width:1100px){ .bi-grid{grid-template-columns:1fr 1fr;} }
@media (max-width:640px){ .bi-grid{grid-template-columns:1fr;} #page-bi .kpis{grid-template-columns:repeat(2,1fr);} }
.bi-grid .card{margin-bottom:0;padding:10px 13px;border-radius:var(--r-lg);}
.bi-grid .card-head{margin-bottom:4px;}
.bi-grid .card-title{font-size:11.5px;}
.bi-grid .card.tall{grid-row:span 2;}
@media (max-width:640px){ .bi-grid .card.tall{grid-row:span 1;} }
.bi-donut-wrap{display:flex;align-items:center;justify-content:center;padding:2px 0;}
.hbar-row{display:flex;align-items:center;gap:8px;margin-bottom:6px;}
.hbar-row .hbar-label{width:96px;flex-shrink:0;font-size:10px;color:var(--text);text-align:right;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;}
.hbar-row .hbar-track{flex:1;height:12px;border-radius:var(--r-pill);background:var(--sunken);border:1px solid var(--border);overflow:hidden;position:relative;}
.hbar-row .hbar-fill{height:100%;border-radius:var(--r-pill) 0 0 var(--r-pill);transition:width var(--dur-slower) var(--ease-out);}
.hbar-row .hbar-val{width:auto;flex-shrink:0;font-family:var(--font-figure);font-size:9.5px;font-weight:700;color:var(--muted);white-space:nowrap;}
</style>
</head>
<body>
<div class="app">

  <!-- ================= SIDEBAR ================= -->
  <aside class="sidebar" id="sidebar">
    <div class="sidebar-brand">
      <svg class="brand-mark" viewBox="0 0 36 36" xmlns="http://www.w3.org/2000/svg">
        <circle cx="14" cy="18" r="10.5" fill="none" stroke="var(--text)" stroke-width="2.4"/>
        <circle cx="23" cy="18" r="10.5" fill="none" stroke="#E8622C" stroke-width="2.4"/>
      </svg>
      <span class="sb-word">
        <span class="sb-name">__OBRA_NOME__</span>
        <span class="sb-kicker">Controle de Obra</span>
      </span>
    </div>

    <nav class="sidebar-nav">
      <div class="nav-group">
        <span class="nav-group-label">Menu Principal</span>
        <button class="pagenav-btn active" data-page="inicio" type="button"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M15 21v-8a1 1 0 0 0-1-1h-4a1 1 0 0 0-1 1v8"/><path d="M3 10a2 2 0 0 1 .709-1.528l7-5.999a2 2 0 0 1 2.582 0l7 5.999A2 2 0 0 1 21 10v9a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2z"/></svg><span class="nav-text">Início</span></button>
        <button class="pagenav-btn" data-page="bi" type="button"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M3 3v16a2 2 0 0 0 2 2h16"/><path d="M18 17V9"/><path d="M13 17V5"/><path d="M8 17v-3"/></svg><span class="nav-text">Visão BI</span></button>
        <button class="pagenav-btn" data-page="horas" type="button"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><circle cx="12" cy="12" r="10"/><polyline points="12 6 12 12 16 14"/></svg><span class="nav-text">Horas</span></button>
        <button class="pagenav-btn" data-page="financeiro" type="button"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><circle cx="12" cy="12" r="10"/><path d="M16 8h-6a2 2 0 1 0 0 4h4a2 2 0 1 1 0 4H8"/><path d="M12 18V6"/></svg><span class="nav-text">Financeiro</span></button>
      </div>
      <div class="nav-group">
        <span class="nav-group-label">Referência</span>
        <button class="pagenav-btn" data-page="glossario" type="button"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><circle cx="12" cy="12" r="10"/><path d="M9.09 9a3 3 0 0 1 5.83 1c0 2-3 3-3 3"/><path d="M12 17h.01"/></svg><span class="nav-text">Explicações</span></button>
      </div>
    </nav>

    <div class="sidebar-foot">
      <div class="sb-context">
        <div class="sb-ctx-label">Base de Dados</div>
        <p class="subtitle" id="subtitleRange">Período: --</p>
      </div>
      <div class="sb-actions">
        <span class="sb-hint">Tema</span>
        <button class="icon-btn theme-toggle" id="themeToggle" type="button" title="Alternar tema claro/escuro" aria-label="Alternar tema">
          <svg class="ico-sun" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><circle cx="12" cy="12" r="4"/><path d="M12 2v2"/><path d="M12 20v2"/><path d="m4.93 4.93 1.41 1.41"/><path d="m17.66 17.66 1.41 1.41"/><path d="M2 12h2"/><path d="M20 12h2"/><path d="m6.34 17.66-1.41 1.41"/><path d="m19.07 4.93-1.41 1.41"/></svg>
          <svg class="ico-moon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M12 3a6 6 0 0 0 9 9 9 9 0 1 1-9-9Z"/></svg>
        </button>
      </div>
    </div>
  </aside>
  <div class="sidebar-scrim" id="sidebarScrim"></div>

  <!-- ================= ÁREA PRINCIPAL ================= -->
  <div class="main">
    <header class="topbar">
      <button class="icon-btn menu-btn" id="menuBtn" type="button" title="Menu" aria-label="Abrir menu">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><line x1="4" x2="20" y1="6" y2="6"/><line x1="4" x2="20" y1="12" y2="12"/><line x1="4" x2="20" y1="18" y2="18"/></svg>
      </button>
      <div class="topbar-title">
        <h1 id="pageTitle">Início</h1>
        <p class="page-sub" id="pageSub">Panorama de efetivo, custo e avanço de obra.</p>
      </div>
      <div class="controls">
        <div class="msf" id="monthFilterWrap"><button class="msf-btn" id="monthFilterBtn" type="button" title="Mês"></button><div class="msf-panel" id="monthFilterPanel" hidden></div></div>
        <select id="localFilter" title="Local">
          <option value="all">Obra + Matriz</option>
          <option value="Obra">Só Obra</option>
          <option value="Matriz">Só Matriz</option>
        </select>
        <div class="msf" id="setorFilterWrap"><button class="msf-btn" id="setorFilterBtn" type="button" title="Setor"></button><div class="msf-panel" id="setorFilterPanel" hidden></div></div>
      </div>
    </header>

    <div class="wrap">
  <!-- ================= PÁGINA: INÍCIO ================= -->
  <div class="page active" id="page-inicio">
    <div class="kpis" id="kpisHome"></div>

    <div class="card">
      <div class="card-head">
        <h2 class="card-title">Efetivo Mensal &mdash; Previsto x Realizado</h2>
        <div style="display:flex;align-items:center;gap:12px;flex-wrap:wrap;">
          <div class="legend" id="homeHistLegend"></div>
          <button class="btn active" id="toggleHomeHistMode">Ver Horas</button>
        </div>
      </div>
      <div id="homeHistWrap"><svg id="homeHistChart"></svg></div>
    </div>

    <div class="card">
      <div class="card-head">
        <h2 class="card-title">Avanço Físico &mdash; Planejado x Realizado</h2>
        <div class="legend" id="homeCurvaSLegend"></div>
      </div>
      <div id="homeCurvaSWrap"><svg id="homeCurvaSChart"></svg></div>
    </div>

    <div class="card">
      <div class="card-head"><h2 class="card-title">Leitura Executiva</h2></div>
      <ul id="execSummary" style="margin:0;padding-left:18px;font-size:12.5px;line-height:1.9;color:var(--text);"></ul>
    </div>

    <div class="card">
      <div class="card-head"><h2 class="card-title">Navegar</h2></div>
      <div class="nav-cards" id="navCards"></div>
    </div>
  </div>

  <!-- ================= PÁGINA: VISÃO BI ================= -->
  <div class="page" id="page-bi">
    <div class="bi-layout">
      <aside class="bi-sidebar">
        <div class="bi-filter-title">Filtros</div>
        <label class="bi-filter-label">Mês</label>
        <div class="msf msf-block" id="biFiltroMesWrap"><button class="msf-btn" id="biFiltroMesBtn" type="button"></button><div class="msf-panel" id="biFiltroMesPanel" hidden></div></div>
        <label class="bi-filter-label">Setor</label>
        <div class="msf msf-block" id="biFiltroSetorWrap"><button class="msf-btn" id="biFiltroSetorBtn" type="button"></button><div class="msf-panel" id="biFiltroSetorPanel" hidden></div></div>
      </aside>

      <div class="bi-main">
        <div class="kpis" id="kpisBI"></div>

        <div class="bi-grid">
          <div class="card">
            <div class="card-head">
              <h2 class="card-title">Horas por Mês &mdash; Previsto x Realizado</h2>
              <div class="legend" id="biHorasLegend"></div>
            </div>
            <div id="biHorasWrap" style="width:100%;overflow-x:auto;"><svg id="biHorasChart"></svg></div>
          </div>

          <div class="card">
            <div class="card-head">
              <h2 class="card-title">Avanço Físico &mdash; Planejado x Realizado</h2>
              <div class="legend" id="biCurvaLegend"></div>
            </div>
            <div id="biCurvaWrap" style="width:100%;overflow-x:auto;"><svg id="biCurvaChart"></svg></div>
          </div>

          <div class="card tall">
            <div class="card-head"><h2 class="card-title">Top Funções &mdash; Horas Extras</h2></div>
            <div id="biRankingWrap"></div>
          </div>

          <div class="card">
            <div class="card-head">
              <h2 class="card-title">Custo por Mês &mdash; Previsto x Realizado</h2>
              <div class="legend" id="biCustoLegend"></div>
            </div>
            <div id="biCustoWrap" style="width:100%;overflow-x:auto;"><svg id="biCustoChart"></svg></div>
          </div>

          <div class="card">
            <div class="card-head">
              <h2 class="card-title">Horas Extras por Tipo &mdash; Previsto x Realizado</h2>
              <div class="legend" id="biExtrasLegend"></div>
            </div>
            <div id="biExtrasWrap" style="width:100%;overflow-x:auto;"><svg id="biExtrasChart"></svg></div>
          </div>
        </div>
      </div>
    </div>
  </div>

  <!-- ================= PÁGINA: HORAS ================= -->
  <div class="page" id="page-horas">
    <div class="kpis" id="kpisHoras"></div>

    <div class="card">
      <div class="card-head">
        <h2 class="card-title">Horas por Mês &mdash; Previsto x Realizado</h2>
        <div style="display:flex;align-items:center;gap:12px;flex-wrap:wrap;">
          <div class="legend" id="chartLegend"></div>
          <button class="btn" id="toggleDetail">Ver detalhe por tipo de hora extra</button>
        </div>
      </div>
      <div id="chartWrap"><svg id="chart"></svg></div>
    </div>

    <details class="tbl-toggle">
      <summary>
        <span class="tt-title">Horas por Setor &mdash; Previsto x Realizado</span>
        <span class="tt-actions">
          <button class="btn" data-csv="setorHorasTable" data-filename="resumo_por_setor_horas">Exportar CSV</button>
        </span>
        <span class="tt-chevron"></span>
      </summary>
      <div class="tt-body">
      <div style="overflow-x:auto;">
      <table id="setorHorasTable">
        <thead><tr>
          <th data-key="setor">Setor</th>
          <th data-key="pessoas" class="num">Func. c/ Apontamento</th>
          <th data-key="previsto" class="num">Horas Previstas</th>
          <th data-key="normais" class="num">Horas Normais</th>
          <th data-key="extra" class="num">Horas Extras</th>
          <th data-key="total" class="num">Horas Total</th>
          <th data-key="saldo" class="num">Diferença</th>
          <th data-key="pctExtra" class="num">% Extra</th>
        </tr></thead>
        <tbody></tbody>
      </table>
      </div>
      </div>
    </details>

    <details class="tbl-toggle">
      <summary>
        <span class="tt-title">Horas por Função &mdash; Previsto x Realizado</span>
        <span class="tt-actions">
          <input class="search" id="funcaoSearch" placeholder="Buscar função ou setor...">
          <button class="btn" data-csv="funcaoTable" data-filename="resumo_por_funcao">Exportar CSV</button>
        </span>
        <span class="tt-chevron"></span>
      </summary>
      <div class="tt-body">
      <div style="overflow-x:auto;">
      <table id="funcaoTable">
        <thead><tr>
          <th data-key="funcao">Função (Sienge / Cargo Agrupado)</th>
          <th data-key="setor">Setor</th>
          <th data-key="local">Local</th>
          <th data-key="pessoas" class="num">Func. c/ Apontamento</th>
          <th data-key="previsto" class="num">Horas Previstas</th>
          <th data-key="normais" class="num">Horas Normais</th>
          <th data-key="extra" class="num">Horas Extras</th>
          <th data-key="total" class="num">Horas Total</th>
          <th data-key="saldo" class="num">Diferença</th>
          <th data-key="pctExtra" class="num">% Extra</th>
        </tr></thead>
        <tbody></tbody>
      </table>
      </div>
      </div>
    </details>

    <details class="tbl-toggle">
      <summary>
        <span class="tt-title">Horas por Função &mdash; Orçado x Realizado - Sienge</span>
        <span class="tt-chevron"></span>
      </summary>
      <div class="tt-body">
      <div style="overflow-x:auto;">
      <table id="siengeHorasTableHoras">
        <thead><tr><th>Função</th><th class="num">Orçamento (h)</th><th class="num">Horas Normais</th><th class="num">Horas Extras</th><th class="num">Total Apontado (h)</th><th class="num">Diferença (c/ Hora Extra)</th><th class="num">Diferença (s/ Hora Extra)</th></tr></thead>
        <tbody></tbody>
      </table>
      </div>
      </div>
    </details>
  </div>

  <!-- ================= PÁGINA: FINANCEIRO ================= -->
  <div class="page" id="page-financeiro">
    <div class="kpis" id="kpisFin"></div>

    <div class="card">
      <div class="card-head">
        <h2 class="card-title">Custo por Mês &mdash; Previsto x Realizado</h2>
        <div class="legend" id="finChartLegend"></div>
      </div>
      <div id="finChartWrap"><svg id="finChart"></svg></div>
      <p class="footnote" id="finChartTotais" style="font-size:12.5px;"></p>
    </div>

    <details class="tbl-toggle">
      <summary>
        <span class="tt-title">Custo por Setor &mdash; Previsto x Realizado</span>
        <span class="tt-actions">
          <button class="btn" data-csv="finSetorTable" data-filename="custo_por_setor">Exportar CSV</button>
        </span>
        <span class="tt-chevron"></span>
      </summary>
      <div class="tt-body">
      <div style="overflow-x:auto;">
      <table id="finSetorTable">
        <thead><tr>
          <th data-key="setor">Setor</th>
          <th data-key="pessoas" class="num">Func. c/ Apontamento</th>
          <th data-key="custoPrevisto" class="num">Custo Previsto</th>
          <th data-key="custoEfetivo" class="num">Salário</th>
          <th data-key="custoExtra" class="num">Custo Hora Extra</th>
          <th data-key="custoTotal" class="num">Custo Total</th>
          <th data-key="custoPP" class="num">Custo Médio / Pessoa</th>
          <th data-key="custoPorHora" class="num">Custo / Hora</th>
          <th data-key="pctTotal" class="num">% do Total</th>
        </tr></thead>
        <tbody></tbody>
      </table>
      </div>
      </div>
    </details>

    <details class="tbl-toggle">
      <summary>
        <span class="tt-title">Custo por Tipo de Adicional</span>
        <span class="tt-chevron"></span>
      </summary>
      <div class="tt-body">
      <div style="overflow-x:auto;">
      <table id="extraTipoTable">
        <thead><tr>
          <th>Tipo</th>
          <th class="num">Horas</th>
          <th class="num">Custo (R$)</th>
          <th class="num">% do Custo de Hora Extra</th>
        </tr></thead>
        <tbody></tbody>
      </table>
      </div>
      </div>
    </details>

    <details class="tbl-toggle">
      <summary>
        <span class="tt-title">Custo por Função &mdash; Orçado x Realizado - Sienge</span>
        <span class="tt-chevron"></span>
      </summary>
      <div class="tt-body">
      <div style="overflow-x:auto;">
      <table id="siengeTable">
        <thead><tr><th>Função</th><th class="num">Orçamento</th><th class="num">Valor Hora Normal</th><th class="num">Valor Hora Extra</th><th class="num">Total Apontado</th><th class="num">Diferença (c/ Hora Extra)</th><th class="num">Diferença (s/ Hora Extra)</th></tr></thead>
        <tbody></tbody>
      </table>
      </div>
      </div>
    </details>

    <!-- Análise de Previsão (KPIs + Projeção + Curva S) oculta a pedido do cliente (2026-09-04) —
         ainda não vai ser usada. Não apagado de propósito, só escondido (atributo "hidden") pra
         voltar fácil depois: é só tirar o "hidden" da div abaixo. -->
    <div id="finAnalisePrevisaoWrap" hidden>
    <p class="section-block-title">Análise de Previsão</p>
    <details class="tbl-toggle" id="finPrevisaoDetails">
      <summary>
        <span class="tt-title">Ver Análise de Previsão (KPIs, Projeção e Curva S)</span>
        <span class="tt-chevron"></span>
      </summary>
      <div class="tt-body">
      <div class="kpis" id="kpisFinProjecao"></div>

      <div class="card">
        <div class="card-head">
          <h2 class="card-title">Projeção de Custo até o Fim da Obra</h2>
          <div class="legend" id="projChartLegend"></div>
        </div>
        <div id="projChartWrap"><svg id="projChart"></svg></div>
        <p class="footnote" id="projChartFootnote"></p>
        <details class="tbl-toggle" style="margin-top:12px;box-shadow:none;border:1px solid var(--border);">
          <summary>
            <span class="tt-title">Projeção de Custo &mdash; Detalhamento Mensal</span>
            <span class="tt-chevron"></span>
          </summary>
          <div class="tt-body">
          <div style="overflow-x:auto;">
          <table id="projTable">
            <thead><tr><th>Mês</th><th>Origem</th><th class="num">Custo do Mês</th><th class="num">Acumulado</th></tr></thead>
            <tbody></tbody>
          </table>
          </div>
          </div>
        </details>
      </div>

      <div class="card" style="margin-bottom:0;">
        <div class="card-head">
          <h2 class="card-title">Curva S: Previsto x Projeção ao Longo do Tempo</h2>
          <div class="legend" id="projDiffChartLegend"></div>
        </div>
        <div id="projDiffChartWrap"><svg id="projDiffChart"></svg></div>
        <p class="footnote" id="projDiffChartFootnote"></p>
      </div>
      </div>
    </details>
    </div>
  </div>

  <!-- ================= PÁGINA: EXPLICAÇÕES ================= -->
  <div class="page" id="page-glossario">
    <p class="section-sub">Toda a metodologia e as notas que antes ficavam dentro dos cards moraram pra cá &mdash; os cards das outras páginas ficaram só com título + visual. Organizado por página, na ordem em que os elementos aparecem nela.</p>

    <div class="card">
      <div class="card-head"><h2 class="card-title">Início</h2></div>
      <ul style="margin:0;padding-left:18px;font-size:12.5px;line-height:1.9;color:var(--text);">
        <li><strong>Gráfico 1 (Efetivo Mensal &mdash; Previsto x Realizado), modo Pessoas:</strong> Previsto = headcount do histograma de mão de obra por função/mês. Executado = pessoas distintas com hora lançada na folha naquele mês. Acima do Previsto sugere revisar efetivo.</li>
        <li><strong>Gráfico 1, modo Horas:</strong> Previsto = teto orçado pelo histograma de mão de obra (headcount &times; 220h/pessoa/mês). Executado = horas normais + extras realmente lançadas na folha. Ficar abaixo do Previsto é economia, não é uma meta a bater.</li>
        <li><strong>Gráfico 2 (Avanço Físico &mdash; Planejado x Realizado):</strong> as duas curvas (âmbar/violeta) são o Avanço Físico de Construção (planejado x realizado, linha "Construção" do resumo de curvas do Planejamento — não a linha "Geral" do projeto inteiro, que mistura fases sem gente lançando hora em campo), acumulado ao longo de todo o cronograma.</li>
        <li><strong>Leitura Executiva:</strong> gerado automaticamente a partir dos dados no filtro atual &mdash; não substitui sua leitura, é um ponto de partida.</li>
      </ul>
    </div>

    <div class="card">
      <div class="card-head"><h2 class="card-title">Aba Horas</h2></div>
      <ul style="margin:0;padding-left:18px;font-size:12.5px;line-height:1.9;color:var(--text);">
        <li><strong>Jornada padrão do mês (220h x 110h):</strong> o padrão é <strong>220h/pessoa/mês</strong>. Quando a pessoa trabalhou <strong>menos de 130h</strong> naquele mês (entrou ou saiu no meio do mês, por exemplo), consideramos a jornada daquele mês como <strong>110h</strong> em vez de 220h &mdash; vale só pra esse mês específico dessa pessoa, não muda a régua de ninguém mais. Isso afeta "Horas Normais"/"Horas Mês" em todas as tabelas e gráficos de Horas e Financeiro (inclusive o Salário por Hora nas colunas de hora extra, que sempre divide por 220, independente dessa regra).</li>
        <li><strong>Período da folha ponto (dia 21 ao dia 20) e centro de custo:</strong> a folha ponto de cada mês fecha do dia <strong>21 do mês anterior</strong> até o dia <strong>20 do mês corrente</strong>. Dentro desse período, até o dia <strong>15</strong> é definido em qual centro de custo (obra) a pessoa vai ficar registrada; a partir do dia <strong>16</strong>, quem não teve mudança definida permanece no centro de custo (obra) em que já estava antes. Essa regra é da rotina de fechamento da folha, fora do dashboard &mdash; citada aqui só pra explicar por que a mudança de obra de alguém pode não aparecer de imediato no mês em que ela realmente aconteceu.</li>
        <li><strong>Como ler os tipos de hora extra (Ex50%&ndash;Ex110%):</strong> toda hora extra aqui é tratada como <strong>desvio de jornada</strong>, não necessariamente hora extra "ruim" ou mal planejada. <strong>Ex100%</strong> = horas apontadas em sábados, domingos e feriados. Os demais percentuais (Ex50%, Ex60%, Ex70%, Ex80%, Ex110%) normalmente indicam que o funcionário veio de <strong>outra obra, em outro estado, e foi transferido</strong> pra este projeto &mdash; cada obra tem sua própria convenção coletiva de hora extra, então o percentual reflete a convenção do local de origem da pessoa, não um "nível" de gravidade da hora extra neste projeto.</li>
        <li><strong>Gráfico 1 (Horas por Mês &mdash; Previsto x Realizado):</strong> Barra da esquerda = Previsto (histograma da Obra e/ou 220h da Matriz, conforme filtro Local). Barra da direita = Realizado, empilhada com Horas Normais embaixo (220h por pessoa/mês, ou 110h no mês em que a pessoa trabalhou menos de 130h &mdash; ver nota acima) e Horas Extras em cima (Ex50%&ndash;Ex110% + Not.Tot., adicional noturno).</li>
        <li><strong>Tabela 1 (Horas por Setor):</strong> agrega todas as funções (cargos agrupados) do setor, somando Obra + Matriz conforme o filtro Local do topo. "Func. c/ Apontamento" = quantidade de funcionários distintos que registraram hora no(s) mês(es) do filtro atual (não é o quadro de pessoal total, é quem efetivamente apontou).</li>
        <li><strong>Tabela 2 (Horas por Função):</strong> Função consolidada = coluna "Função x SIENGE" da planilha quando a pessoa tem um valor real ali (função confirmada pessoa a pessoa; "Não Previsto" não conta) &mdash; sem isso, cai na coluna "Cargo Agrupado". O Previsto continua vindo sempre do histograma (por Cargo Agrupado), mesmo quando o rótulo da linha é a Função SIENGE. Diferença = Previsto &minus; Total (negativo = estourou o previsto).</li>
        <li><strong>Tabela 3 (Horas por Função &mdash; Orçado x Realizado - Sienge):</strong> mesma lógica da Tabela 4 da aba Financeiro &mdash; Orçamento vem SOMENTE da aba "Relatório" da planilha-fonte (coluna "Quantidade"), sem reserva de outra fonte. Reunida aqui também pra ficar tudo de horas num lugar só.</li>
      </ul>
    </div>

    <div class="card">
      <div class="card-head"><h2 class="card-title">Aba Financeiro</h2></div>
      <ul style="margin:0;padding-left:18px;font-size:12.5px;line-height:1.9;color:var(--text);">
        <li><strong>Metodologia geral:</strong> esta base não tem Receita, Cliente, Centro de Custo nem Margem &mdash; por isso esses indicadores não aparecem aqui. <strong>Custo Previsto</strong> = salário médio da função (histograma da Obra + efetivo real da Matriz) × pessoas previstas. <strong>Custo Realizado</strong> = salário-base + hora extra de quem apontou.</li>
        <li><strong>Gráfico 1 (Custo por Mês &mdash; Previsto x Realizado):</strong> Barra da esquerda = Previsto (mesma régua do "Custo Previsto" do topo: histograma da Obra × salário médio da função + efetivo real da Matriz). Barra da direita = Realizado, empilhada com Salário-base embaixo e Custo Hora Extra em cima (Ex50%&ndash;Ex110% &times; salário/hora + Adicional Noturno, salário/hora &times; 1,2). Linha tracejada = custo médio realizado por mês do período.</li>
        <li><strong>Tabela 2 (Custo por Setor):</strong> "Custo Previsto" = histograma (HIST-MO) — salário médio da função × pessoas previstas, somando só os meses do filtro Mês do topo (se "Todos os meses", soma FEV a JUN/2026; não é o empreendimento inteiro). "Salário", "Custo Hora Extra", "Custo Total" e "Func. c/ Apontamento" = dados realizados (aba "Base de dados"), também restritos ao mesmo filtro de Mês/Local/Setor do topo. "Custo Médio / Pessoa" e "Custo / Hora" são ponderados por pessoa-mês (não por pessoa distinta do período), pra não distorcer com quem entrou/saiu no meio do caminho.</li>
        <li><strong>Tabela 3 (Custo por Tipo de Adicional):</strong> hora extra aqui é tratada como desvio de jornada, não indicador de má gestão por si só. Ex100% = sábados, domingos e feriados; os demais percentuais geralmente indicam funcionário transferido de outra obra. Custo de cada percentual = soma da coluna "Valor Hora X%" da planilha. Adicional Noturno = salário/hora &times; 1,2 &times; horas de Not.Tot. (20% a mais) &mdash; já entra no "Custo Hora Extra" das outras páginas.</li>
        <li><strong>Tabela 4 (Custo por Função &mdash; Orçado x Realizado - Sienge):</strong> a lista de funções e o Orçamento vêm SOMENTE da aba "Relatório" da planilha-fonte (relatório do Sienge, coluna "Preço total reajustado") &mdash; sem reserva/estimativa de outra fonte. Funções sem linha nesse relatório (ex: "Encarregado de Terraplanagem") aparecem como "Sem orçamento" de propósito. A função de cada pessoa vem da coluna "Função x SIENGE" da planilha; Eletricista Comum + Força Controle são somados numa linha só. Valor Hora Normal = salário-base de quem apontou. Valor Hora Extra = custo de hora extra (Ex50%&ndash;Ex110% + adicional noturno). Total Apontado = soma dos dois, somando TODOS os meses lançados até agora (não filtra por Mês/Setor do topo; respeita só o filtro Local — Obra/Matriz/Todos —, já que o Orçamento do Sienge é um valor fixo do contrato inteiro, não fatiado por mês). Diferença (c/ Hora Extra) = Orçamento &minus; Total Apontado. Diferença (s/ Hora Extra) = Orçamento &minus; só o Valor Hora Normal. O % entre parênteses é a diferença sobre o Orçamento. Linha "Total" no rodapé: Diferença = soma do Orçamento &minus; soma do Total Apontado (não é a soma das diferenças de cada função).</li>
        <li><strong>Horas por Função &mdash; Orçado x Realizado - Sienge:</strong> só na aba Horas agora (pra ficar tudo de horas num lugar só) &mdash; mesma lógica da Tabela 4 acima, em horas: Orçamento = coluna "Quantidade" do relatório do Sienge.</li>
        <li hidden><strong>Análise de Previsão:</strong> "Projeção até Fim da Obra" = realizado até agora + histograma dos meses que faltam, mês a mês. "Previsto até o Fim" = orçamento do histograma pro projeto inteiro (Obra + Matriz), fixo, não recalculado com o realizado.</li>
      </ul>
    </div>
  </div>
    </div>
  </div>
</div>
<div class="tooltip" id="tooltip"></div>

<script id="raw-data" type="application/json">__DATA_JSON__</script>
<script>
__JS_BODY__
</script>
</body>
</html>
"""

JS_BODY = r"""
const PAYLOAD = JSON.parse(document.getElementById('raw-data').textContent);
const RAW = PAYLOAD.rows;
const HIST_PREVISTO = PAYLOAD.histPrevisto;
// Topografia passou a ser tratada como parte do setor Civil (pedido do cliente). Normalizado
// aqui, uma vez só, logo após carregar os dados — assim todo agregado que já existe (por setor,
// histograma, KPIs, gráficos, tabelas, curva) enxerga Topografia dentro de Civil sem precisar
// mexer em cada função individualmente. A função/cargo "Topógrafo" continua existindo
// normalmente — só o SETOR dela muda, o cargo em si não é apagado.
RAW.forEach(r=>{ if(r.setor==='Topografia') r.setor='Civil'; });
HIST_PREVISTO.forEach(h=>{ if(h.setor==='Topografia') h.setor='Civil'; });
const HIST_CONTRATADO = PAYLOAD.histContratado;
const HIST_LIBERADO = PAYLOAD.histLiberado;
const PROJECT_MONTHS = PAYLOAD.projectMonths;
const HOURS_PP = PAYLOAD.hoursPerPersonMonth;
const TODAY_KEY = PAYLOAD.today.slice(0,7);
const SETOR_ORDER = PAYLOAD.setorOrder.filter(s=>s!=='Topografia');
const OBRA_NOME = PAYLOAD.obraNome;
const CARGO_TO_HISTFUNC = PAYLOAD.cargoToHistFunc;
// Alguns Cargo Agrupado genéricos (ex: "Operador de Maquinas") cobrem várias funções do
// histograma ao mesmo tempo (confirmado pessoa a pessoa) — CARGO_TO_HISTFUNC guarda isso como
// lista nesses casos. Esse helper normaliza pra sempre devolver uma lista (vazia se não tem
// nenhuma), pra quem consome não precisar checar se é string, lista ou null toda vez.
function histFuncList(cargo){
  const hf = CARGO_TO_HISTFUNC[cargo];
  if(!hf) return [];
  return Array.isArray(hf) ? hf : [hf];
}
// Sentido inverso (histFunc -> Cargo(s) Agrupado): usado pra "vaga aberta" no Custo Previsto —
// quando o histograma planeja gente pra uma função num mês mas ainda não tem NINGUÉM real
// contratado com esse cargo naquele mês, precisamos de um nome de Cargo Agrupado pra rotular
// essa linha nas tabelas (Custo por Função etc.), já que não existe funcionário real pra copiar.
const HISTFUNC_TO_CARGOS = {};
// Processa os mapeamentos 1:1 (cargo específico) ANTES dos genéricos (lista, tipo "Operador de
// Maquinas") — assim, quando uma função do histograma tem tanto um cargo específico quanto
// aparece numa lista genérica (ex: "Operador de Retroescavadeira" sozinho E dentro da lista de
// "Operador de Maquinas"), o cargo específico fica em [0] e não perde o Previsto pro genérico.
const cargoEntries = Object.entries(CARGO_TO_HISTFUNC);
[...cargoEntries.filter(([,hf])=>hf && !Array.isArray(hf)), ...cargoEntries.filter(([,hf])=>Array.isArray(hf))]
  .forEach(([cargo,hf])=>{
    const hfs = Array.isArray(hf) ? hf : [hf];
    hfs.forEach(h=>{
      if(!HISTFUNC_TO_CARGOS[h]) HISTFUNC_TO_CARGOS[h] = [];
      HISTFUNC_TO_CARGOS[h].push(cargo);
    });
  });

// ============================================================
// RELATÓRIO DO SIENGE — ORÇADO POR FUNÇÃO (lido automaticamente da aba "Relatório")
// ============================================================
// PAYLOAD.siengeOrcado já vem pronto do Python, direto da aba "Relatório" da planilha-fonte
// (colunas "Descrição" e "Preço total reajustado") — não precisa colar nada manualmente aqui.
// Essa aba só tem o valor ORÇADO por função (não tem "já gasto"), então o Saldo a Gastar é
// calculado abaixo em siengeComparativo(): Orçado − Executado (soma de todos os lançamentos
// daquela função na base, salário+hora extra, acumulado até agora).
const SIENGE_ORCADO = PAYLOAD.siengeOrcado;

// Avanço Físico (curva do Planejamento, lida de um Excel externo — ver ler_avanco_fisico() no
// build_dashboard.py). Só a linha "Construção" (CeM — não "Geral" do projeto inteiro, não por
// Setor) — mesma curva do card "Construção" na aba "CURVAS GERENCIAIS" da planilha do
// Planejamento (pedido do cliente, 2026-08-27). Vem pronta em % por mês (0-100); os Maps são pra
// lookup rápido por mesKey
// na hora de desenhar a Curva S.
const AVANCO_FISICO = PAYLOAD.avancoFisico || {planejado:[], realizado:[]};
const AVANCO_FISICO_PLAN = new Map(AVANCO_FISICO.planejado.map(d=>[d.mesKey, d.pct]));
const AVANCO_FISICO_REAL = new Map(AVANCO_FISICO.realizado.map(d=>[d.mesKey, d.pct]));
// O nome da função no relatório do Sienge (coluna "Descrição") pode ser diferente do "Cargo
// Agrupado" usado neste dashboard — mapeie aqui: chave = nome no Sienge, valor = funcaoBase
// correspondente no dashboard. Critério: só entra no mapa quando existe um cargo já lançado na
// base que corresponde de forma inequívoca (mesma função, nome só varia por convenção). Funções
// do Sienge sem correspondência real no efetivo atual (ex: Médico, Pintor, Engenheiro genérico,
// Supervisor Administrativo/Elétrica/Montagem — a base só tem "Supervisor de Obras" e
// "Supervisor de Planejamento", nenhum dos dois claramente equivalente a essas 3) ficam de
// propósito fora do mapa: o resultado correto pra elas é "R$ 0 executado, saldo = orçado
// inteiro" (ninguém contratado ainda pra essa função, não é um bug de nomenclatura).
const SIENGE_TO_FUNCAO = {
  "Ajudante": "Ajudante de Obra",
  "Eletricista Força Controle": "Eletricista FC",
  "Encarregado Montagem": "Encarregado de Montagem",
  "Engenheiro de Segurança": "Engenheiro Seguranca Trabalho",
  "Técnico de Segurança": "Tecnico Seguranca do Trabalho",
  "Técnico Meio Ambiente": "Tecnico de Meio Ambiente",
  "Supervisor Civil": "Supervisor de Obras",
  "Enfermeira": "Enfermeiro Socorrista",
  "Zeladora": "Auxiliar de Serviços Gerais",
};
// Linhas do relatório do Sienge que na prática são a MESMA função na obra (só separadas em duas
// linhas no orçamento) — soma o orçado das duas e compara com o executado de quem exerce essa
// função de verdade na base (confirmado pelo cliente: Eletricista Comum e Eletricista Força
// Controle são o mesmo time, só a base só tem gente lançada como "Eletricista FC").
const SIENGE_GRUPOS = [
  { label: "Eletricista (Comum + Força Controle)", membros: ["Eletricista Comum", "Eletricista Força Controle"] },
];
// Funções que ainda NÃO estão no relatório do Sienge, mas o cliente já quer acompanhar separado
// (ex: Terraplanagem, que não tem linha de orçamento no Sienge). Casam por Cargo Agrupado, não
// pela coluna "Função x SIENGE" (ver siengeRowsFor). Ficam sem Previsto até existir uma
// referência (relatório do Sienge ou linha própria no histograma) —
// confirmado pelo cliente: "pode deixar sem previsto, depois iremos adicionar".
const SIENGE_EXTRAS = [
  // "Encarregado de terraplanagem" já é um Cargo Agrupado real da base (com mapeamento pro
  // histograma via CARGO_TO_HISTFUNC) — só não tinha linha própria nessa tabela do Sienge ainda.
  // "Supervisor de Terraplanagem" foi removido a pedido do cliente (2026-08-26) — não entra mais
  // nas tabelas Orçado x Realizado - Sienge (Horas e Financeiro).
  { label: "Encarregado de Terraplanagem", membros: ["Encarregado de terraplanagem"] },
];
// Pessoas cujo "Cargo Agrupado" na base não reflete a função real que exercem na obra (ex: Cicero
// está lançado como "Supervisor de Obras" mas na prática é "Supervisor de Montagem"; Clenilson
// está como "Gerente de Obras" mas na prática é o engenheiro da obra) NÃO precisam mais de
// override manual aqui — a coluna "Função x SIENGE" da planilha (preenchida pessoa por pessoa,
// ver r.funcaoSienge) já resolve isso na fonte. SIENGE_TO_FUNCAO acima só entra como fallback
// (por Cargo Agrupado) pra planilhas antigas, sem essa coluna ainda.
// Universo de meses válidos pra filtrar/ordenar monthsPresent e overlapMonths — usa PROJECT_MONTHS
// (o cronograma inteiro do histograma, já lido dinamicamente do HIST-MO) em vez de uma lista fixa
// de meses. Antes era uma lista de datas fixa no código (só FEV a JUL/2026) que exigia editar o
// código toda vez que um mês novo era lançado na "Base de dados" — agora qualquer mês dentro do
// cronograma do projeto (passado ou futuro) já é reconhecido sozinho, sem precisar mexer aqui.
const MONTH_ORDER = PROJECT_MONTHS;
const EXTRA_KEYS = ['ex50','ex60','ex70','ex80','ex100','ex110','notTot'];
const PCT_EXTRA_KEYS = ['ex50','ex60','ex70','ex80','ex100','ex110'];
const EXTRA_LABELS = {ex50:'50%', ex60:'60%', ex70:'70%', ex80:'80%', ex100:'100%', ex110:'110%', notTot:'Noturno'};
const EXTRA_COLORS = {ex50:'var(--ex50)', ex60:'var(--ex60)', ex70:'var(--ex70)', ex80:'var(--ex80)', ex100:'var(--ex100)', ex110:'var(--ex110)', notTot:'var(--notTot)'};
const MES_ABBR = ['JAN','FEV','MAR','ABR','MAI','JUN','JUL','AGO','SET','OUT','NOV','DEZ'];
const mesLabel = mk => `${MES_ABBR[parseInt(mk.slice(5,7),10)-1]}/${mk.slice(0,4)}`;
// Horizonte do histograma — lido do próprio dado (não fixo), pra não quebrar de novo se o
// usuário editar as datas do HIST-MO na planilha outra vez.
const PROJECT_START_LABEL = mesLabel(PROJECT_MONTHS[0]);
const PROJECT_END_LABEL = mesLabel(PROJECT_MONTHS[PROJECT_MONTHS.length-1]);
const SETOR_ESCRITORIO = null; // não existe mais — "Local" resolve isso agora

const fmtH = n => (n||0).toLocaleString('pt-BR',{minimumFractionDigits:1,maximumFractionDigits:1}) + 'h';
const fmtN = n => (n||0).toLocaleString('pt-BR',{minimumFractionDigits:0,maximumFractionDigits:0});
const fmtR = n => (n||0).toLocaleString('pt-BR',{style:'currency',currency:'BRL'});
const fmtPct = n => (n==null ? '—' : n.toFixed(1)+'%');
// Sinalização visual padrão dos KPIs "Previsto x Realizado": seta + cor + texto tipo "▼ 8,7% vs
// previsto". pct = realizado/previsto*100. <=100% (dentro do previsto) = verde/▼; >100%
// (estourou) = vermelho/▲; exatamente 100% = bolinha neutra. null (sem previsto pra comparar) =
// string vazia, o card cai pro texto de sub padrão.
function kpiSignal(pct){
  if(pct==null) return '';
  const delta = pct-100;
  const bom = delta<=0;
  const seta = delta<0 ? '▼' : (delta>0 ? '▲' : '●');
  const cor = bom ? 'var(--green-dark)' : 'var(--red)';
  return `<span class="kpi-signal" style="color:${cor};">${seta} ${Math.abs(delta).toFixed(1)}% vs previsto</span>`;
}

// Mês e Setor são multi-seleção (array; [] = sem filtro, equivalente ao
// antigo 'all'). Local continua single-select (string) — só tem 2 valores reais (Obra/Matriz),
// então "selecionar os dois" já é a mesma coisa que "Todos", e o código dele é todo baseado em
// branch (!=='Matriz'/!=='Obra'), não em filtro simples — não ganha nada virando array.
let state = {month:[], local:'all', setor:[], chartDetail:false, homeHistMode:'pessoas',
  funcaoSort:{key:'normais',dir:-1}, setorSort:{key:'custo',dir:-1},
  setorHorasSort:{key:'total',dir:-1}, finFuncaoSort:{key:'custoTotal',dir:-1}, extraTipoSort:{key:'valor',dir:-1},
  funcaoSearch:'', refMonth:null, funcaoPrevistoDe:null, funcaoPrevistoAte:null};

const monthsPresent = MONTH_ORDER.filter(m => RAW.some(r=>r.mesKey===m));
// Meses do cronograma (histograma) que vêm DEPOIS do último mês com dado lançado — não qualquer
// mês fora de monthsPresent. O histograma começa em JAN/2026 mas os lançamentos só começam em
// FEV/2026 (1 mês de defasagem ANTES do início dos dados); esse mês não deve entrar como "sem
// dado lançado" pra projeção/limite de ritmo, porque já passou e não faz sentido projetar gasto
// futuro pra ele — só os meses que ainda vêm pela frente (depois do último mês real) contam.
function mesesFuturos(){
  if(monthsPresent.length===0) return PROJECT_MONTHS.slice();
  const ultimoMes = monthsPresent[monthsPresent.length-1];
  const idx = PROJECT_MONTHS.indexOf(ultimoMes);
  return idx===-1 ? PROJECT_MONTHS.filter(m=>!monthsPresent.includes(m)) : PROJECT_MONTHS.slice(idx+1);
}
// "Tem baseline no histograma" = tem headcount de verdade lançado pra esse mês (algum registro
// em HIST_PREVISTO), não só "a coluna do mês existe no HIST-MO". A planilha atual tem colunas de
// JAN e FEV/2026 no histograma mas com headcount zerado em todas as funções (o ramp-up real só
// começa em MAR/2026) — sem essa checagem, FEV apareceria com "0h previsto" (parecendo economia
// de 100%) em vez de cair no mesmo fallback que a Matriz já usa (efetivo real como referência).
const histActiveMonths = new Set(HIST_PREVISTO.map(h=>h.mesKey));
// Não passa do último mês com dado real lançado (monthsPresent) — overlapMonths é "meses de
// referência até agora", não o cronograma inteiro. Sem esse teto, virar MONTH_ORDER pra
// PROJECT_MONTHS (pra reconhecer mês novo sozinho, sem editar código) faria overlapMonths incluir
// TODOS os meses futuros do histograma, inflando Previsto de agregados que usam overlapMonths
// direto pro cronograma inteiro em vez de só até hoje.
const ultimoMesPresente = monthsPresent[monthsPresent.length-1];
const overlapMonths = MONTH_ORDER.filter(m=>histActiveMonths.has(m) && (!ultimoMesPresente || m<=ultimoMesPresente));
state.refMonth = overlapMonths[overlapMonths.length-1];
// Período padrão da tabela "Custo Previsto por Função": todo o intervalo já lançado (pra já
// nascer comparando Previsto x Realizado de verdade, não só um mês isolado).
state.funcaoPrevistoDe = monthsPresent[0] || PROJECT_MONTHS[0];
state.funcaoPrevistoAte = monthsPresent[monthsPresent.length-1] || PROJECT_MONTHS[0];

// ---------- filtro multi-seleção (dropdown com checkboxes) ----------
// Reaproveitado pelos filtros de Mês/Setor (topo + espelho na Visão BI) e Função (só
// Visão BI). msfInstances guarda toda instância criada; refreshAllMsf() redesenha TODAS de uma
// vez depois de qualquer mudança — é assim que o topo e a barra da BI ficam sincronizados sem
// precisar de um "sync" manual: as duas instâncias de Mês (por ex.) leem/escrevem o mesmo
// state.month, então redesenhar todo mundo já reflete a mudança nos dois lugares.
const msfInstances = [];
const openMsfPanels = new Set();
function closeAllMsfPanels(){ openMsfPanels.forEach(p=>p.hidden=true); openMsfPanels.clear(); }
document.addEventListener('click', ()=>closeAllMsfPanels());
function refreshAllMsf(){ msfInstances.forEach(inst=>inst.render()); }
function makeMultiSelectFilter({btnId, panelId, getOptions, getSelected, setSelected, allLabel, itemNoun}){
  // getOptions() => [{value,label}]. getSelected()/setSelected(arr) leem/escrevem o array no state.
  const btn = document.getElementById(btnId);
  const panel = document.getElementById(panelId);
  if(!btn || !panel) return {render(){}};
  function render(){
    const options = getOptions();
    const selected = getSelected();
    const selectedLabels = options.filter(o=>selected.includes(o.value)).map(o=>o.label);
    btn.textContent = selected.length===0 ? allLabel : selected.length===1 ? selectedLabels[0] : `${selected.length} ${itemNoun||'selecionados'}`;
    panel.innerHTML = '';
    const actions = document.createElement('div');
    actions.className = 'msf-panel-actions';
    const btnAll = document.createElement('button');
    btnAll.type = 'button'; btnAll.textContent = 'Selecionar todos';
    btnAll.addEventListener('click', e=>{ e.stopPropagation(); setSelected(options.map(o=>o.value)); refreshAllMsf(); renderAll(); });
    const btnClear = document.createElement('button');
    btnClear.type = 'button'; btnClear.textContent = 'Limpar';
    btnClear.addEventListener('click', e=>{ e.stopPropagation(); setSelected([]); refreshAllMsf(); renderAll(); });
    actions.appendChild(btnAll); actions.appendChild(btnClear);
    panel.appendChild(actions);
    if(options.length===0){
      const empty = document.createElement('div');
      empty.className = 'msf-empty'; empty.textContent = 'Sem opções no filtro atual.';
      panel.appendChild(empty);
    }
    options.forEach(opt=>{
      const optLabel = document.createElement('label');
      optLabel.className = 'msf-option';
      const cb = document.createElement('input');
      cb.type = 'checkbox';
      cb.checked = selected.includes(opt.value);
      cb.addEventListener('change', ()=>{
        const cur = getSelected();
        const next = cb.checked ? [...cur, opt.value] : cur.filter(v=>v!==opt.value);
        setSelected(next);
        refreshAllMsf();
        renderAll();
      });
      const span = document.createElement('span');
      span.textContent = opt.label;
      optLabel.appendChild(cb); optLabel.appendChild(span);
      panel.appendChild(optLabel);
    });
  }
  btn.addEventListener('click', e=>{
    e.stopPropagation();
    const wasHidden = panel.hidden;
    closeAllMsfPanels();
    if(wasHidden){ panel.hidden = false; openMsfPanels.add(panel); }
  });
  panel.addEventListener('click', e=>e.stopPropagation());
  const instance = {render};
  msfInstances.push(instance);
  return instance;
}

// ---------- filtros globais ----------
const setoresPresentes = [...new Set(RAW.map(r=>r.setor))].sort();
const monthMsfTop = makeMultiSelectFilter({
  btnId:'monthFilterBtn', panelId:'monthFilterPanel',
  getOptions: () => monthsPresent.map(m=>({value:m, label:RAW.find(r=>r.mesKey===m).mes})),
  getSelected: () => state.month, setSelected: v=>{ state.month = v; },
  allLabel:'Todos os meses', itemNoun:'meses selecionados'
});

const localSel = document.getElementById('localFilter');
localSel.addEventListener('change', ()=>{ state.local = localSel.value; renderAll(); });

const setorMsfTop = makeMultiSelectFilter({
  btnId:'setorFilterBtn', panelId:'setorFilterPanel',
  getOptions: () => setoresPresentes.map(s=>({value:s, label:s})),
  getSelected: () => state.setor, setSelected: v=>{ state.setor = v; },
  allLabel:'Todos os setores', itemNoun:'setores selecionados'
});

const firstMes = RAW.find(r=>r.mesKey===monthsPresent[0]).mes;
const lastMes = RAW.find(r=>r.mesKey===monthsPresent[monthsPresent.length-1]).mes;
document.getElementById('subtitleRange').textContent = `${OBRA_NOME}  •  ${firstMes} a ${lastMes}  •  ${new Set(RAW.map(r=>r.nome)).size} pessoas`;

// ---------- navegação entre páginas ----------
document.querySelectorAll('.pagenav-btn').forEach(btn=>{
  btn.addEventListener('click', ()=>{ goToPage(btn.dataset.page); });
});
function goToPage(page){
  document.querySelectorAll('.pagenav-btn').forEach(b=>b.classList.toggle('active', b.dataset.page===page));
  document.querySelectorAll('.page').forEach(p=>p.classList.toggle('active', p.id==='page-'+page));
  document.body.classList.toggle('bi-active', page==='bi');
  renderChart(); renderHomeHistograma(); renderBIAvancoFisicoChart({wrap:'homeCurvaSWrap', svg:'homeCurvaSChart', legend:'homeCurvaSLegend'}); renderFinChart(); renderProjChart(); renderCurvaSChart(); renderBIPage();
}

function filteredRows(){
  return RAW.filter(r=>
    (state.month.length===0 || state.month.includes(r.mesKey)) &&
    (state.local==='all' || r.local===state.local) &&
    (state.setor.length===0 || state.setor.includes(r.setor))
  );
}
// obraRows() NÃO aplica o filtro Local (a página Avanço da Obra é sempre só Obra,
// então um filtro Local=Matriz não pode zerar essa página — só respeita Mês e Setor).
function obraRows(){
  return RAW.filter(r=>
    (state.month.length===0 || state.month.includes(r.mesKey)) &&
    (state.setor.length===0 || state.setor.includes(r.setor)) &&
    r.local==='Obra'
  );
}

function sortArr(arr, key, dir){
  return arr.slice().sort((a,b)=>{
    let va=a[key], vb=b[key];
    if(va==null && vb==null) return 0;
    if(va==null) return 1;
    if(vb==null) return -1;
    if(typeof va==='string') return va.localeCompare(vb)*dir;
    return (va-vb)*dir;
  });
}
// Soma as colunas numéricas das linhas JÁ FILTRADAS/BUSCADAS que a tabela está mostrando (mesma
// lógica de qualquer planilha: soma da coluna visível) — null/undefined não entram na soma (ex:
// "Previsto" de uma função só-Matriz sem baseline no histograma), então uma célula "—" na tabela
// vira uma contribuição de 0 no total, não vira erro. Não deduplica pessoa entre grupos (se
// alguém aparecesse em 2 Setores/Funções no mesmo período contaria nos dois) — não visto na base
// atual, mas é a limitação de somar por grupo em vez de recalcular do zero pelas linhas cruas.
function sumCols(arr, keys){
  const t = {};
  keys.forEach(k=>t[k]=0);
  arr.forEach(a=>keys.forEach(k=>{ if(a[k]!=null) t[k]+=a[k]; }));
  return t;
}
function updateSortArrows(tableId, sort){
  document.querySelectorAll('#'+tableId+' thead th[data-key]').forEach(th=>{
    th.querySelectorAll('.sort-arrow').forEach(a=>a.remove());
    if(th.dataset.key===sort.key){
      const arr = document.createElement('span');
      arr.className='sort-arrow'; arr.textContent = sort.dir===1?'▲':'▼';
      th.appendChild(arr);
    }
  });
}
function wireSort(tableId, stateKey, renderFn){
  document.querySelectorAll('#'+tableId+' thead th[data-key]').forEach(th=>{
    th.addEventListener('click', ()=>{
      const key = th.dataset.key;
      if(state[stateKey].key===key) state[stateKey].dir *= -1; else state[stateKey] = {key, dir:-1};
      renderFn();
    });
  });
}

// ---------- Exportar CSV ----------
function tableToCSV(tableId){
  const table = document.getElementById(tableId);
  const rows = [...table.querySelectorAll('tr')].filter(tr=>!tr.classList.contains('detail-row'));
  return rows.map(tr=>[...tr.children].filter(c=>c.tagName!=='TH' || c.textContent.trim()!=='').map(td=>{
    const txt = td.textContent.replace(/\s+/g,' ').trim().replace(/"/g,'""');
    return `"${txt}"`;
  }).join(';')).join('\r\n');
}
function downloadCSV(tableId, filename){
  const csv = '﻿' + tableToCSV(tableId);
  const blob = new Blob([csv], {type:'text/csv;charset=utf-8;'});
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url; a.download = (filename||tableId)+'.csv';
  document.body.appendChild(a); a.click(); document.body.removeChild(a);
  URL.revokeObjectURL(url);
}
document.querySelectorAll('[data-csv]').forEach(btn=>{
  btn.addEventListener('click', ()=>downloadCSV(btn.dataset.csv, btn.dataset.filename));
});

// ---------- Tabelas colapsáveis: impede que clicar num controle (busca, exportar, ordenar,
// seletor de período) dentro do <summary> também abra/feche o <details> ----------
document.querySelectorAll('.tbl-toggle summary').forEach(s=>{
  s.querySelectorAll('button,input,select,a').forEach(el=>{
    ['click','mousedown','keydown'].forEach(evt=>el.addEventListener(evt, e=>e.stopPropagation()));
  });
});
// "Análise de Previsão" (Financeiro) tem gráfico SVG dentro do bloco recolhível — enquanto
// fechado, wrap.clientWidth fica 0 e renderProjChart()/renderCurvaSChart() nem desenham nada
// (mesma guarda de sempre pra não desenhar gráfico invisível). Redesenha na hora que abre.
const finPrevisaoDetailsEl = document.getElementById('finPrevisaoDetails');
if(finPrevisaoDetailsEl){
  finPrevisaoDetailsEl.addEventListener('toggle', ()=>{
    if(finPrevisaoDetailsEl.open){ renderProjChart(); renderCurvaSChart(); }
  });
}

// ---------- Matriz: efetivo mensal (não tem histograma) ----------
// Ambas as funções abaixo respeitam o filtro Setor do topo (state.setor), para o
// Previsto ficar consistente com o Realizado quando o usuário filtra por setor.
function matrizHeadcountByMonth(mk){
  return new Set(RAW.filter(r=>r.local==='Matriz' && r.mesKey===mk &&
    (state.setor.length===0 || state.setor.includes(r.setor))).map(r=>r.nome)).size;
}
function previstoObraHoras(months){
  return HIST_PREVISTO.filter(h=>months.includes(h.mesKey) && (state.setor.length===0 || state.setor.includes(h.setor)))
    .reduce((s,h)=>s+h.headcount*HOURS_PP,0);
}
function previstoMatrizHoras(months){
  return months.reduce((s,mk)=>s+matrizHeadcountByMonth(mk)*HOURS_PP,0);
}
// Previsto combinado respeitando os filtros "Local" e "Setor": Obra usa histograma (só meses
// com baseline), Matriz usa 220h (todo mês tem baseline, inclusive FEV/2026).
function previstoCombinado(months){
  let total = 0;
  // histActiveMonths (não overlapMonths) — overlapMonths só cobre a janela de MONTH_ORDER (meses
  // com folha real), mas o histograma de Obra tem baseline pra obra inteira, inclusive meses
  // futuros que ainda não têm lançamento (usado pelo histograma da Início, que mostra até o fim
  // do projeto). overlapMonths continua sendo o certo pra tudo que é "mês de referência atual".
  if(state.local!=='Matriz') total += previstoObraHoras(months.filter(m=>histActiveMonths.has(m)));
  if(state.local!=='Obra') total += previstoMatrizHoras(months);
  return total;
}
// Mesma régua do previstoCombinado, mas em quantidade de pessoas (headcount) em vez de horas —
// usado pelo toggle "Ver quantidade de pessoas" do histograma da Início. Só faz sentido por mês
// individual (headcount não soma entre meses sem contar a mesma pessoa várias vezes).
function previstoPessoasMes(mk){
  let total = 0;
  if(state.local!=='Matriz' && histActiveMonths.has(mk)){
    total += HIST_PREVISTO.filter(h=>h.mesKey===mk && (state.setor.length===0||state.setor.includes(h.setor))).reduce((s,h)=>s+h.headcount,0);
  }
  if(state.local!=='Obra') total += matrizHeadcountByMonth(mk);
  return total;
}

// Previsto por Cargo Agrupado (a pedido do cliente, 2026-08-26: "Horas Previstas é só o que está
// no histograma; se for Matriz, o efetivo real da Matriz"): Obra usa SOMENTE o histograma
// (CARGO_TO_HISTFUNC) — cargo de Obra sem função lá não tem Previsto (null, mostra "—"), não
// inventa um número a partir do efetivo real. Matriz nunca tem histograma, então sempre usa
// "efetivo real daquele cargo naquele mês × 220h".
function previstoPorFuncao(cargoAgrupado, months, local){
  const histFuncs = histFuncList(cargoAgrupado);
  if(local==='Obra'){
    if(histFuncs.length===0) return null;
    const overlap = months.filter(m=>overlapMonths.includes(m));
    return HIST_PREVISTO.filter(h=>histFuncs.includes(h.funcaoBase) && overlap.includes(h.mesKey))
      .reduce((s,h)=>s+h.headcount*HOURS_PP,0);
  }
  return months.reduce((s,mk)=>{
    const pessoas = new Set(RAW.filter(r=>r.funcaoBase===cargoAgrupado && r.mesKey===mk && r.local===local).map(r=>r.nome)).size;
    return s + pessoas*HOURS_PP;
  }, 0);
}
// ================= CUSTO PREVISTO (A5+A8+B1) =================
// Regra única usada em TODA a base de custo do dashboard:
// - Custo Previsto de uma Função num Mês = salário médio da função (prioriza a coluna "Salario
//   medio por função" da aba HIST-MO — mesma referência do Custo Previsto do projeto inteiro;
//   só cai pro salário real médio pago NAQUELE MÊS, tirado da folha, quando a função não tem
//   essa referência no HIST-MO — ver avgSalarioParaPrevisto) × nº de pessoas PREVISTAS da função
//   naquele mês (A5): usa o histograma quando a função tem uma linha lá (Obra); quando não tem —
//   Matriz nunca tem — usa o efetivo real daquele mês como proxy (mesmo raciocínio já usado pra
//   Horas Previstas da Matriz).
// - Custo médio por pessoa (A8): nunca "custo total ÷ pessoas distintas do período inteiro"
//   (isso trata quem trabalhou 1 mês como se tivesse trabalhado o período todo). É sempre
//   "custo total ÷ pessoa-mês" — como cada linha da base já é uma pessoa-mês, isso equivale a
//   dividir pela CONTAGEM DE LINHAS do grupo, não pela quantidade de nomes distintos.

// Usado só como RESERVA do Custo Previsto (quando a função não tem referência na aba HIST-MO).
function avgSalarioEfetivoPorFuncaoMes(funcaoBase, mesKey){
  const rs = RAW.filter(r=>r.funcaoBase===funcaoBase && r.mesKey===mesKey &&
    (state.local==='all'||r.local===state.local) && (state.setor.length===0||state.setor.includes(r.setor)));
  if(rs.length===0) return 0;
  return rs.reduce((s,r)=>s+r.salario,0)/rs.length;
}

function headcountPrevistoPorFuncaoMes(funcaoBase, mesKey, local){
  const histFuncs = histFuncList(funcaoBase);
  if(local==='Obra' && histFuncs.length>0 && histActiveMonths.has(mesKey)){
    const entries = HIST_PREVISTO.filter(h=>histFuncs.includes(h.funcaoBase) && h.mesKey===mesKey && (state.setor.length===0||state.setor.includes(h.setor)));
    if(entries.length>0) return entries.reduce((s,h)=>s+h.headcount,0);
    // Mês dentro do range de datas do histograma mas sem headcount pra ESSA função específica
    // nesse mês (0 lançado) — cai pro efetivo real, mesmo fallback usado abaixo pra Matriz.
  }
  return new Set(RAW.filter(r=>r.funcaoBase===funcaoBase && r.mesKey===mesKey && r.local===local &&
    (state.setor.length===0||state.setor.includes(r.setor))).map(r=>r.nome)).size;
}

// Salário usado no Custo Previsto: prioriza a coluna "Salario medio por função" que o cliente
// calculou na aba HIST-MO (mesma fonte do Custo Previsto do projeto inteiro, na Suporte à
// Decisão — PAYLOAD.salarioMedioHist) — assim os dois "Custo Previsto" do dashboard usam a
// mesma referência salarial. Só cai pro salário real médio pago naquele mês (folha) quando a
// função não tem referência na aba HIST-MO, pra não zerar o previsto de funções sem essa linha.
function avgSalarioParaPrevisto(funcaoBase, mesKey){
  const histFuncs = histFuncList(funcaoBase);
  const avgMap = avgSalarioPorHistFunc();
  if(histFuncs.length===1 && avgMap.has(histFuncs[0]) && avgMap.get(histFuncs[0])>0) return avgMap.get(histFuncs[0]);
  if(histFuncs.length>1){
    // Cargo com mais de uma função do histograma (ex: "Operador de Maquinas") — média ponderada
    // pelo headcount de cada função naquele mês, pra custoPrevistoPorFuncaoMes (que multiplica
    // isso pelo headcount TOTAL somado) dar o mesmo resultado que somar função por função.
    let totalHc = 0, totalCusto = 0;
    histFuncs.forEach(hf=>{
      const hc = HIST_PREVISTO.filter(h=>h.funcaoBase===hf && h.mesKey===mesKey).reduce((s,h)=>s+h.headcount,0);
      const sal = avgMap.get(hf)||0;
      if(hc>0 && sal>0){ totalHc += hc; totalCusto += hc*sal; }
    });
    if(totalHc>0) return totalCusto/totalHc;
  }
  return avgSalarioEfetivoPorFuncaoMes(funcaoBase, mesKey);
}
function custoPrevistoPorFuncaoMes(funcaoBase, mesKey, local){
  return avgSalarioParaPrevisto(funcaoBase, mesKey) * headcountPrevistoPorFuncaoMes(funcaoBase, mesKey, local);
}

// Soma o Custo Previsto de todas as funções presentes no filtro atual (Mês/Local/Setor),
// agrupando pela chave que a chamada pedir (setor, função, local, ou um total único).
function custoPrevistoAgrupado(groupKeyFn, monthsOverride){
  const months = monthsOverride || (state.month.length===0 ? monthsPresent : state.month);
  const map = new Map();
  const avgMap = avgSalarioPorHistFunc();
  months.forEach(mk=>{
    // ---- OBRA: direto do histograma — TODA linha do HIST-MO com headcount planejado pra esse
    // mês entra, preenchida ou não. É o "teto orçado" de verdade (igual a planilha HIST-MO do
    // cliente): uma vaga planejada mas ainda sem ninguém contratado continua pesando no Previsto
    // (sem isso, Previsto e Executado ficam os dois zerados juntos, escondendo a vaga aberta).
    if(state.local!=='Matriz'){
      HIST_PREVISTO.filter(h=>h.mesKey===mk && h.headcount>0 && (state.setor.length===0||state.setor.includes(h.setor))).forEach(h=>{
        const cargos = HISTFUNC_TO_CARGOS[h.funcaoBase] || [h.funcaoBase];
        const fb = cargos[0];
        let avgCusto = avgMap.get(h.funcaoBase)||0;
        if(avgCusto===0) avgCusto = avgSalarioEfetivoPorFuncaoMes(fb, mk); // sem ref no HIST-MO: cai pro salário real do mês
        if(avgCusto===0) return; // sem referência salarial de nenhuma fonte — não dá pra estimar
        const key = groupKeyFn({funcaoBase: fb, setor: h.setor, local: 'Obra'});
        map.set(key, (map.get(key)||0) + avgCusto*h.headcount);
      });
    }
    // ---- MATRIZ: o histograma não cobre Matriz, então continua usando o efetivo real do mês ×
    // salário real médio pago naquele mês (mesma régua de sempre).
    if(state.local!=='Obra'){
      const rowsThisMonth = RAW.filter(r=>r.mesKey===mk && r.local==='Matriz' && (state.setor.length===0||state.setor.includes(r.setor)));
      const byFuncao = new Map();
      rowsThisMonth.forEach(r=>{ if(!byFuncao.has(r.funcaoBase)) byFuncao.set(r.funcaoBase, r); });
      byFuncao.forEach((sampleRow, fb)=>{
        const cp = custoPrevistoPorFuncaoMes(fb, mk, 'Matriz');
        const key = groupKeyFn(sampleRow);
        map.set(key, (map.get(key)||0) + cp);
      });
    }
  });
  return map;
}
function custoPrevistoTotalFiltro(){
  let total = 0;
  custoPrevistoAgrupado(()=>'x').forEach(v=>total+=v);
  return total;
}
// Previsto de UM mês específico, ignorando o filtro Mês do topo (respeita só Local/Setor) — pra
// desenhar a barra de Previsto do gráfico "Evolução do Custo por Mês", que sempre mostra todos
// os meses lançados lado a lado, igual o gráfico de Horas já faz.
function custoPrevistoParaMes(mk){
  let total = 0;
  custoPrevistoAgrupado(()=>'x', [mk]).forEach(v=>total+=v);
  return total;
}

// ================= PÁGINA: INÍCIO =================
function setorAgg2(rows){
  const map = new Map();
  rows.forEach(r=>{
    if(!map.has(r.setor)) map.set(r.setor, {setor:r.setor, normais:0, extra:0, custo:0});
    const a = map.get(r.setor);
    a.normais+=r.horasMes; a.extra+=r.extraTotal; a.custo += r.salario+r.custoHoraExtra;
  });
  return [...map.values()].map(a=>({...a, total:a.normais+a.extra, pctExtra: (a.normais+a.extra)>0 ? a.extra/(a.normais+a.extra)*100 : 0}));
}

function renderHome(){
  const rows = filteredRows();
  const totalPessoas = new Set(rows.map(r=>r.nome)).size;
  const pessoasObra = new Set(rows.filter(r=>r.local==='Obra').map(r=>r.nome)).size;
  const pessoasMatriz = new Set(rows.filter(r=>r.local==='Matriz').map(r=>r.nome)).size;
  const months = state.month.length===0 ? monthsPresent : state.month;
  const horasPrevistas = previstoCombinado(months);
  const totalNormaisHome = rows.reduce((s,r)=>s+r.horasMes,0);
  const totalExtraHome = rows.reduce((s,r)=>s+r.extraTotal,0);
  const horasRealizadas = totalNormaisHome + totalExtraHome;
  const custoExtraH = rows.reduce((s,r)=>s+r.custoHoraExtra,0);
  const custoSalarioH = rows.reduce((s,r)=>s+r.salario,0);
  const custoPrevisto = custoPrevistoTotalFiltro(); // A5: salário médio da função × pessoas previstas
  const custoTotal = custoSalarioH + custoExtraH; // Custo Realizado — real e independente do Previsto

  const kpis = [
    {label:'Nº de Obras', value:'1', sub:OBRA_NOME, color:'var(--accent)'},
    {label:'Funcionários com Apontamento', value: fmtN(totalPessoas), sub: `${pessoasObra} Obra · ${pessoasMatriz} Matriz`, color:'var(--accent)'},
    {label:'Horas Previstas', value: fmtH(horasPrevistas), color:'var(--accent)'},
    {label:'Horas Realizadas', value: fmtH(horasRealizadas), breakdown: [{label:'Normal', value:fmtH(totalNormaisHome), color:'var(--green)'}, {label:'Extra', value:fmtH(totalExtraHome), color:'var(--red)'}], color:'var(--accent)'},
    {label:'Custo Previsto', value: fmtR(custoPrevisto), color:'var(--accent)'},
    {label:'Custo Realizado', value: fmtR(custoTotal), breakdown: [{label:'Salário', value:fmtR(custoSalarioH), color:'var(--green)'}, {label:'Extra', value:fmtR(custoExtraH), color:'var(--red)'}], color:'var(--accent)'},
  ];
  document.getElementById('kpisHome').innerHTML = kpis.map(k=>`
    <div class="kpi"><div class="label">${k.label}</div><div class="value" style="color:${k.color}">${k.value}</div>${k.breakdown ? `<div class="kpi-breakdown">${k.breakdown.map(b=>`<span class="kpi-bd-item"><span class="dot" style="background:${b.color}"></span>${b.label} <b>${b.value}</b></span>`).join('')}</div>` : (k.sub ? `<div class="sub">${k.sub}</div>` : '')}</div>
  `).join('');

  renderExecSummary(rows);

  const navs = [
    {page:'bi', icon:'BI', title:'Visão BI', desc:'Só gráficos e cards comparativos, sem tabela.', color:'var(--previsto)'},
    {page:'horas', icon:'H', title:'Dashboard de Horas', desc:'Previsto x Realizado, horas normais/extras, Sienge Horas.', color:'var(--green-dark)'},
    {page:'financeiro', icon:'R$', title:'Dashboard Financeiro', desc:'Custo de folha: normal e hora extra, por setor e por local.', color:'var(--accent)'},
  ];
  document.getElementById('navCards').innerHTML = navs.map(n=>`
    <button class="nav-card" data-goto="${n.page}" style="cursor:pointer;">
      <div class="nc-icon" style="background:color-mix(in srgb, ${n.color} 16%, transparent);color:${n.color};">${n.icon}</div>
      <div class="nc-title">${n.title}</div>
      <div class="nc-desc">${n.desc}</div>
    </button>
  `).join('');
  document.querySelectorAll('[data-goto]').forEach(btn=>btn.addEventListener('click', ()=>goToPage(btn.dataset.goto)));

}

function renderExecSummary(rows){
  const el = document.getElementById('execSummary');
  if(rows.length===0){ el.innerHTML = '<li>Nenhum dado no filtro atual.</li>'; return; }
  const bullets = [];

  // A3: removida a tolerância de 15% (não havia sido definida com o cliente) — mostra só o
  // indicador base (índice de hora extra por setor), sem julgar se está "bom" ou "ruim".
  const setores = setorAgg2(rows).filter(a=>a.total>0).sort((a,b)=>b.pctExtra-a.pctExtra);
  const setorMaiorHE = setores[0];
  if(setorMaiorHE) bullets.push(`Setor com maior índice de hora extra: <strong>${setorMaiorHE.setor}</strong> (${setorMaiorHE.pctExtra.toFixed(1)}% do total de horas).`);

  const setorCusto = [...setores].sort((a,b)=>b.custo-a.custo)[0];
  if(setorCusto) bullets.push(`Setor de maior custo de folha: <strong>${setorCusto.setor}</strong> (${fmtR(setorCusto.custo)}).`);

  el.innerHTML = bullets.map(b=>`<li>${b}</li>`).join('');
}

// Histograma mês a mês da Início: só Previsto x Executado (sem normal/extra empilhado — versão
// "limpa" pedida pelo cliente). Segue o mesmo padrão de renderChart()/monthlyAgg(): sempre
// mostra todos os meses lançados (monthsPresent), sem recortar por state.month — o filtro de Mês
// já não recorta o gráfico equivalente da página Horas, então mantemos a mesma consistência aqui.
// Local e Setor do topo são respeitados normalmente.
function renderHomeHistograma(){
  const wrap = document.getElementById('homeHistWrap');
  if(!wrap || wrap.clientWidth===0) return;
  const isPessoas = state.homeHistMode==='pessoas';
  const fmtVal = isPessoas ? fmtN : fmtH;
  // Rótulo do valor em cima da barra usa formato compacto (18 meses lado a lado não cabe
  // "23.320,0h" por extenso sem colidir) — o tooltip continua mostrando o valor completo.
  const fmtValLabel = isPessoas ? fmtN : (v => (v||0).toLocaleString('pt-BR',{notation:'compact',maximumFractionDigits:0})+' h');
  // Mostra o horizonte inteiro do histograma (PROJECT_MONTHS), não só os meses já lançados —
  // pedido do cliente pra ver o previsto até o fim da obra, não só até o mês atual. Meses
  // futuros (isFuturo) entram com Executado ainda zerado/tracejado, não é dado real.
  const data = PROJECT_MONTHS.map(mk=>{
    const rows = RAW.filter(r=>r.mesKey===mk &&
      (state.local==='all' || r.local===state.local) && (state.setor.length===0 || state.setor.includes(r.setor)));
    const executado = isPessoas ? new Set(rows.map(r=>r.nome)).size : rows.reduce((s,r)=>s+r.horasMes+r.extraTotal,0);
    const previsto = isPessoas ? previstoPessoasMes(mk) : previstoCombinado([mk]);
    const hasPrevisto = (state.local==='Matriz') || (state.local==='all') || (state.local==='Obra' && histActiveMonths.has(mk));
    const isFuturo = !monthsPresent.includes(mk);
    return {mes: (rows[0]||{mes:mesLabel(mk)}).mes, previsto, executado, hasPrevisto, isFuturo};
  });
  const svg = document.getElementById('homeHistChart');
  const W = Math.max(wrap.clientWidth, 1080), H = 300;
  const marginL = 60, marginR = 20, marginT = 20, marginB = 40;
  const plotW = W - marginL - marginR, plotH = H - marginT - marginB;
  svg.setAttribute('viewBox', `0 0 ${W} ${H}`);
  svg.setAttribute('width', W); svg.setAttribute('height', H);
  svg.innerHTML = '';
  const ns = 'http://www.w3.org/2000/svg';
  const mkEl = (tag, attrs) => { const el = document.createElementNS(ns, tag); for(const k in attrs) el.setAttribute(k, attrs[k]); return el; };
  const tooltip = document.getElementById('tooltip');

  if(data.length===0){ document.getElementById('homeHistLegend').innerHTML=''; return; }

  const groupW = plotW / data.length;
  const barW = Math.min(46, groupW*0.32);
  const maxVal = Math.max(...data.map(d=>Math.max(d.previsto, d.executado)), 10) * 1.15;

  const ticks = 5;
  for(let i=0;i<=ticks;i++){
    const val = maxVal/ticks*i;
    const y = marginT + plotH - (val/maxVal)*plotH;
    svg.appendChild(mkEl('line', {x1:marginL, x2:W-marginR, y1:y, y2:y, class:'gridline'}));
    const t = mkEl('text', {x:marginL-8, y:y+4, 'text-anchor':'end', class:'axis'});
    t.textContent = Math.round(val).toLocaleString('pt-BR');
    svg.appendChild(t);
  }
  svg.appendChild(mkEl('line', {x1:marginL, x2:marginL, y1:marginT, y2:marginT+plotH, class:'axis'}));

  function addTooltipHandlers(el, html){
    el.addEventListener('mousemove', (e)=>{
      tooltip.style.opacity=1; tooltip.innerHTML = html;
      tooltip.style.left = (e.clientX+14)+'px'; tooltip.style.top = (e.clientY+10)+'px';
    });
    el.addEventListener('mouseleave', ()=>{ tooltip.style.opacity=0; });
  }

  // Horizonte inteiro (até 18 meses): rótulo curto ("JAN/26") + fonte menor pra caber todo mês
  // sem sobrepor, em vez de pular rótulo. O skip só entra como rede de segurança se o container
  // for mesmo assim estreito demais.
  const estLabelW = 46;
  const maxLabels = Math.max(1, Math.floor(plotW/estLabelW));
  const labelSkip = Math.max(1, Math.ceil(data.length/maxLabels));
  data.forEach((d,i)=>{
    const gx = marginL + groupW*i;
    const cx1 = gx + groupW/2 - barW - 4;
    const cx2 = gx + groupW/2 + 4;
    const showLabel = (i % labelSkip === 0 || i===data.length-1);
    if(showLabel){
      const lbl = mkEl('text', {x:gx+groupW/2, y:marginT+plotH+20, 'text-anchor':'middle', class:'axis', style:'font-size:8.5px;'});
      lbl.textContent = d.mes.replace('/20','/');
      svg.appendChild(lbl);
    }

    function valueLabel(x, y, texto, cor){
      const t = mkEl('text', {x, y:y-5, 'text-anchor':'middle', style:`font-size:8px;font-weight:700;fill:${cor||'var(--muted)'};`});
      t.textContent = texto;
      svg.appendChild(t);
    }

    if(d.hasPrevisto){
      const hPrev = (d.previsto/maxVal)*plotH;
      const rPrev = mkEl('rect', {x:cx1, y:marginT+plotH-hPrev, width:barW, height:hPrev, rx:4, fill:'var(--previsto)', class:'bar'});
      addTooltipHandlers(rPrev, `<strong>${d.mes} — Previsto</strong><br>${fmtVal(d.previsto)}${isPessoas?' pessoas':''}`);
      svg.appendChild(rPrev);
      if(showLabel) valueLabel(cx1+barW/2, marginT+plotH-hPrev, fmtValLabel(d.previsto), 'var(--previsto)');
    } else {
      const rNoPrev = mkEl('rect', {x:cx1, y:marginT+plotH-14, width:barW, height:14, rx:4, fill:'none', stroke:'var(--previsto)', 'stroke-width':1.5, 'stroke-dasharray':'3,2'});
      addTooltipHandlers(rNoPrev, `<strong>${d.mes} — Previsto</strong><br>Sem baseline do histograma.`);
      svg.appendChild(rNoPrev);
    }

    if(d.isFuturo){
      // Mês ainda não lançado na folha — tracejado, não é "executou zero".
      const rNoExec = mkEl('rect', {x:cx2, y:marginT+plotH-14, width:barW, height:14, rx:4, fill:'none', stroke:'var(--executado)', 'stroke-width':1.5, 'stroke-dasharray':'3,2'});
      addTooltipHandlers(rNoExec, `<strong>${d.mes} — Executado</strong><br>Mês ainda não lançado na folha.`);
      svg.appendChild(rNoExec);
    } else {
      const hExec = (d.executado/maxVal)*plotH;
      const rExec = mkEl('rect', {x:cx2, y:marginT+plotH-hExec, width:barW, height:hExec, rx:4, fill:'var(--executado)', class:'bar'});
      addTooltipHandlers(rExec, `<strong>${d.mes} — Executado</strong><br>${fmtVal(d.executado)}${isPessoas?' pessoas':''}`);
      svg.appendChild(rExec);
      if(showLabel) valueLabel(cx2+barW/2, marginT+plotH-hExec, fmtValLabel(d.executado), 'var(--text)');
    }
  });

  document.getElementById('homeHistLegend').innerHTML = `<span><span class="dot" style="background:var(--previsto)"></span>Previsto</span><span><span class="dot" style="background:var(--executado)"></span>Executado</span>`;

  const footnoteEl = document.getElementById('homeHistFootnote');
  if(footnoteEl){
    const horizonteTxt = ` Mostra o histograma inteiro, de ${PROJECT_START_LABEL} a ${PROJECT_END_LABEL} — meses tracejados ainda não têm lançamento na folha.`;
    footnoteEl.textContent = (isPessoas
      ? 'Efetivo (nº de pessoas). Previsto = headcount do histograma de mão de obra por função/mês. Executado = pessoas distintas com hora lançada na folha naquele mês. Acima do Previsto sugere revisar efetivo.'
      : 'Previsto = teto orçado pelo histograma de mão de obra (headcount × 220h/pessoa/mês). Executado = horas normais + extras realmente lançadas na folha. Ficar abaixo do Previsto é economia, não é uma meta a bater.') + horizonteTxt;
  }
}

// ================= STATUS DO CRONOGRAMA (Obra) =================
function computeStatus(){
  const totalPrevistoProjeto = HIST_PREVISTO.filter(h=>state.setor.length===0||state.setor.includes(h.setor)).reduce((s,h)=>s+h.headcount*HOURS_PP,0);
  const executadoObra = RAW.filter(r=>r.local==='Obra' && (state.setor.length===0||state.setor.includes(r.setor))).reduce((s,r)=>s+r.horasMes+r.extraTotal,0);
  const avancoPct = totalPrevistoProjeto>0 ? executadoObra/totalPrevistoProjeto*100 : 0;
  const projectStart = new Date(PROJECT_MONTHS[0]+'-01T00:00:00');
  const [ly,lm] = PROJECT_MONTHS[PROJECT_MONTHS.length-1].split('-').map(Number);
  const projectEnd = new Date(ly, lm, 0);
  const today = new Date(PAYLOAD.today+'T00:00:00');
  const diasPrevistos = Math.round((projectEnd-projectStart)/86400000)+1;
  const diasRealizados = Math.max(0, Math.min(diasPrevistos, Math.round((today-projectStart)/86400000)+1));
  const tempoDecorridoPct = diasPrevistos>0 ? diasRealizados/diasPrevistos*100 : 0;
  const desvio = avancoPct - tempoDecorridoPct;
  let label, cls, color;
  if(desvio >= -5){ label='No Prazo'; cls='ok'; color='var(--green-dark)'; }
  else if(desvio >= -15){ label='Atenção'; cls='warn'; color='var(--amber)'; }
  else { label='Atrasado'; cls='bad'; color='var(--red)'; }
  return {label, cls, color, sub: `${desvio>=0?'+':''}${desvio.toFixed(1)}pp vs. cronograma`,
    totalPrevistoProjeto, executadoObra, avancoPct, diasPrevistos, diasRealizados, tempoDecorridoPct, desvio};
}

// ================= PÁGINA: HORAS =================
function renderKPIsHoras(){
  const rows = filteredRows();
  const totalNormais = rows.reduce((s,r)=>s+r.horasMes,0);
  const totalExtra = rows.reduce((s,r)=>s+r.extraTotal,0);
  const realizado = totalNormais+totalExtra;
  const months = state.month.length===0 ? monthsPresent : state.month;
  const previsto = previstoCombinado(months);
  const saldo = previsto - realizado;
  const pctExec = previsto>0 ? realizado/previsto*100 : null;
  const pctExtra = realizado>0 ? (totalExtra/realizado*100) : 0;
  // Decomposição da diferença: quanto do desvio vem do Normal (Previsto − Normal realizado) e
  // quanto vem da Hora Extra (que não tem previsto, então é sempre um desconto puro do saldo).
  // Os dois somam exatamente o saldo total: (previsto−totalNormais) + (−totalExtra) = saldo.
  const saldoNormal = previsto - totalNormais;
  const saldoExtraParte = -totalExtra;
  // % Executado decomposto em pontos percentuais: quanto do desvio (pctExec−100) vem do Normal
  // sozinho passar (ou não) do previsto, e quanto vem da Hora Extra (que não tem previsto próprio,
  // então sempre soma pontos ao desvio). Os dois somam exatamente pctExec−100 — mesma lógica da
  // decomposição em horas/R$ acima, só que em % do previsto: (totalNormais−previsto)/previsto*100
  // + totalExtra/previsto*100 = (realizado−previsto)/previsto*100 = pctExec−100.
  const pctNormalPP = previsto>0 ? (totalNormais-previsto)/previsto*100 : null;
  const pctExtraPP = previsto>0 ? totalExtra/previsto*100 : null;

  const kpis = [
    {label:'Horas Previstas', value: fmtH(previsto), color:'var(--accent)'},
    {label:'Horas Realizadas', value: fmtH(realizado), breakdown: [{label:'Normal', value:fmtH(totalNormais), color:'var(--green)'}, {label:'Extra', value:fmtH(totalExtra), color:'var(--red)'}], color:'var(--accent)'},
    {label:'Horas Extras', value: fmtH(totalExtra), sub: pctExtra.toFixed(1)+'% do realizado', color:'var(--red)'},
    {label:'Diferença (Previsto − Realizado)', value: fmtH(Math.abs(saldo)) + (saldo>=0?' de folga':' estourado'), breakdown: [{label:'Normal', value:(saldoNormal>=0?'+':'')+fmtH(saldoNormal), color: saldoNormal>=0?'var(--green)':'var(--red)'}, {label:'Extra', value:(saldoExtraParte>=0?'+':'')+fmtH(saldoExtraParte), color:'var(--red)'}], color: saldo>=0 ? 'var(--green-dark)' : 'var(--red)'},
    {label:'% Executado', value: fmtPct(pctExec), sub: pctExec==null ? kpiSignal(pctExec) : null, breakdown: pctExec==null ? null : [{label:'Normal', value:(pctNormalPP>=0?'+':'')+pctNormalPP.toFixed(1)+'%', color: pctNormalPP>=0?'var(--red)':'var(--green)'}, {label:'Extra', value:'+'+pctExtraPP.toFixed(1)+'%', color:'var(--red)'}], color: pctExec==null ? 'var(--accent)' : pctExec<=100 ? 'var(--green-dark)' : 'var(--red)'},
  ];
  document.getElementById('kpisHoras').innerHTML = kpis.map(k=>`
    <div class="kpi"><div class="label">${k.label}</div><div class="value" style="color:${k.color}">${k.value}</div>${k.breakdown ? `<div class="kpi-breakdown">${k.breakdown.map(b=>`<span class="kpi-bd-item"><span class="dot" style="background:${b.color}"></span>${b.label} <b>${b.value}</b></span>`).join('')}</div>` : (k.sub ? `<div class="sub">${k.sub}</div>` : '')}</div>
  `).join('');
}

function monthlyAgg(){
  return monthsPresent.map(mk=>{
    const rows = RAW.filter(r=>r.mesKey===mk &&
      (state.local==='all' || r.local===state.local) && (state.setor.length===0 || state.setor.includes(r.setor)));
    const agg = {mesKey:mk, mes: (rows[0]||{mes:mesLabel(mk)}).mes, normais:0, extraTotal:0, custoHoraExtra:0};
    EXTRA_KEYS.forEach(k=>agg[k]=0);
    rows.forEach(r=>{
      agg.normais += r.horasMes; agg.extraTotal += r.extraTotal;
      agg.custoHoraExtra += r.custoHoraExtra;
      EXTRA_KEYS.forEach(k=>agg[k]+=r[k]);
    });
    agg.hasPrevisto = state.local==='Obra' ? overlapMonths.includes(mk) : true;
    agg.previsto = agg.hasPrevisto ? previstoCombinado([mk]) : 0;
    return agg;
  });
}

function renderChartLegend(){
  const el = document.getElementById('chartLegend');
  let html = `<span><span class="dot" style="background:var(--previsto)"></span>Previsto</span><span><span class="dot" style="background:var(--green)"></span>Normais (realizado)</span>`;
  if(!state.chartDetail){
    html += `<span><span class="dot" style="background:var(--red)"></span>Extras (realizado)</span>`;
  } else {
    html += EXTRA_KEYS.map(k=>`<span><span class="dot" style="background:${EXTRA_COLORS[k]}"></span>Extra ${EXTRA_LABELS[k]}</span>`).join('');
  }
  el.innerHTML = html;
}

function renderChart(){
  const wrap = document.getElementById('chartWrap');
  if(!wrap || wrap.clientWidth===0) return;
  const data = monthlyAgg();
  const svg = document.getElementById('chart');
  const W = Math.max(wrap.clientWidth, 560), H = 340;
  const marginL = 60, marginR = 20, marginT = 20, marginB = 46;
  const plotW = W - marginL - marginR, plotH = H - marginT - marginB;
  svg.setAttribute('viewBox', `0 0 ${W} ${H}`);
  svg.setAttribute('width', W); svg.setAttribute('height', H);
  svg.innerHTML = '';

  const groupW = plotW / data.length;
  const barW = Math.min(46, groupW*0.32);
  const maxVal = Math.max(...data.map(d=>Math.max(d.normais+d.extraTotal, d.previsto)), 10) * 1.15;

  const ns = 'http://www.w3.org/2000/svg';
  const mk = (tag, attrs) => { const el = document.createElementNS(ns, tag); for(const k in attrs) el.setAttribute(k, attrs[k]); return el; };

  const ticks = 5;
  for(let i=0;i<=ticks;i++){
    const val = maxVal/ticks*i;
    const y = marginT + plotH - (val/maxVal)*plotH;
    svg.appendChild(mk('line', {x1:marginL, x2:W-marginR, y1:y, y2:y, class:'gridline'}));
    const t = mk('text', {x:marginL-8, y:y+4, 'text-anchor':'end', class:'axis'});
    t.textContent = Math.round(val).toLocaleString('pt-BR');
    svg.appendChild(t);
  }
  svg.appendChild(mk('line', {x1:marginL, x2:marginL, y1:marginT, y2:marginT+plotH, class:'axis'}));

  const tooltip = document.getElementById('tooltip');

  data.forEach((d, i)=>{
    const gx = marginL + groupW*i;
    const cx1 = gx + groupW/2 - barW - 4;
    const cx2 = gx + groupW/2 + 4;

    const lbl = mk('text', {x:gx+groupW/2, y:marginT+plotH+22, 'text-anchor':'middle', class:'axis'});
    lbl.textContent = d.mes;
    svg.appendChild(lbl);

    function addTooltipHandlers(el, html){
      el.addEventListener('mousemove', (e)=>{
        tooltip.style.opacity=1; tooltip.innerHTML = html;
        tooltip.style.left = (e.clientX+14)+'px'; tooltip.style.top = (e.clientY+10)+'px';
      });
      el.addEventListener('mouseleave', ()=>{ tooltip.style.opacity=0; });
    }

    // Rótulo de valor acima da barra — deixa o comparativo visível sem precisar passar o mouse.
    function valueLabel(x, y, texto, cor){
      const t = mk('text', {x, y:y-5, 'text-anchor':'middle', style:`font-size:9.5px;font-weight:700;fill:${cor||'var(--muted)'};`});
      t.textContent = texto;
      svg.appendChild(t);
    }

    // Barra 1: Previsto
    if(d.hasPrevisto){
      const hPrev = (d.previsto/maxVal)*plotH;
      const rPrev = mk('rect', {x:cx1, y:marginT+plotH-hPrev, width:barW, height:hPrev, rx:4, fill:'var(--previsto)', class:'bar'});
      addTooltipHandlers(rPrev, `<strong>${d.mes} — Previsto</strong><br>${fmtH(d.previsto)}`);
      svg.appendChild(rPrev);
      valueLabel(cx1+barW/2, marginT+plotH-hPrev, fmtH(d.previsto), 'var(--previsto)');
    } else {
      const rNoPrev = mk('rect', {x:cx1, y:marginT+plotH-14, width:barW, height:14, rx:4, fill:'none', stroke:'var(--previsto)', 'stroke-width':1.5, 'stroke-dasharray':'3,2'});
      addTooltipHandlers(rNoPrev, `<strong>${d.mes} — Previsto</strong><br>Sem baseline do histograma (antes de ${PROJECT_START_LABEL} ou filtro "Só Obra").`);
      svg.appendChild(rNoPrev);
    }

    // Barra 2: Realizado empilhado (Normais embaixo, Extras em cima)
    const hNorm = (d.normais/maxVal)*plotH;
    const rNorm = mk('rect', {x:cx2, y:marginT+plotH-hNorm, width:barW, height:hNorm, fill:'var(--green)', class:'bar'});
    addTooltipHandlers(rNorm, `<strong>${d.mes} — Normais</strong><br>${fmtH(d.normais)}`);
    svg.appendChild(rNorm);

    if(!state.chartDetail){
      const hEx = (d.extraTotal/maxVal)*plotH;
      const rEx = mk('rect', {x:cx2, y:marginT+plotH-hNorm-hEx, width:barW, height:hEx, rx:3, fill:'var(--red)', class:'bar'});
      addTooltipHandlers(rEx, `<strong>${d.mes} — Extras (total)</strong><br>${fmtH(d.extraTotal)}<br>Custo Hora Extra: ${fmtR(d.custoHoraExtra)}`);
      svg.appendChild(rEx);
    } else {
      let yCursor = marginT+plotH-hNorm;
      EXTRA_KEYS.forEach(k=>{
        const val = d[k];
        if(val<=0) return;
        const h = (val/maxVal)*plotH;
        yCursor -= h;
        const rSeg = mk('rect', {x:cx2, y:yCursor, width:barW, height:h, fill:EXTRA_COLORS[k], class:'bar'});
        addTooltipHandlers(rSeg, `<strong>${d.mes} — Extra ${EXTRA_LABELS[k]}</strong><br>${fmtH(val)}`);
        svg.appendChild(rSeg);
      });
    }
    // Total do Realizado (Normais + Extras) — sempre acima da barra, mesmo que a barra seja
    // empilhada em várias cores, pra dar o número redondo de comparação com o Previsto.
    const hRealTotal = ((d.normais+d.extraTotal)/maxVal)*plotH;
    valueLabel(cx2+barW/2, marginT+plotH-hRealTotal, fmtH(d.normais+d.extraTotal), 'var(--text)');
  });

  renderChartLegend();
}

const toggleDetailBtn = document.getElementById('toggleDetail');
if(toggleDetailBtn) toggleDetailBtn.addEventListener('click', (e)=>{
  state.chartDetail = !state.chartDetail;
  e.target.classList.toggle('active', state.chartDetail);
  e.target.textContent = state.chartDetail ? 'Ver total de horas extras' : 'Ver detalhe por tipo de hora extra';
  renderChart();
});

const toggleHomeHistBtn = document.getElementById('toggleHomeHistMode');
if(toggleHomeHistBtn) toggleHomeHistBtn.addEventListener('click', (e)=>{
  state.homeHistMode = state.homeHistMode==='horas' ? 'pessoas' : 'horas';
  e.target.classList.toggle('active', state.homeHistMode==='pessoas');
  e.target.textContent = state.homeHistMode==='pessoas' ? 'Ver Horas' : 'Ver Efetivo (pessoas)';
  renderHomeHistograma();
});

// ---------- Função table (página Horas) ----------
// Rótulo da linha (a pedido do cliente, 2026-08-26): usa "Função x SIENGE" quando a pessoa tem um
// valor real ali (função confirmada pessoa a pessoa, casos como Cicero/Clenilson onde o Cargo
// Agrupado da base não reflete o cargo real) — "Não Previsto" não conta como valor real (a
// planilha às vezes grava essa mesma sentinela sem o til, "Nao Previsto" — trata as duas iguais).
// Sem Função x SIENGE (planilha antiga) ou com qualquer uma das duas grafias, cai no Cargo
// Agrupado, como antes.
const SEM_FUNCAO_SIENGE = new Set(['Não Previsto', 'Nao Previsto']);
// Reclassificações manuais (a pedido do cliente, 2026-08-26): esses 4 Cargos Agrupados não têm
// função própria no histograma, mas o cliente confirmou que a função de verdade é a mesma de um
// grupo que já existe na tabela — contam junto (rótulo igual ao do grupo real, então somam horas
// na MESMA linha) em vez de aparecerem sozinhos sem Previsto. O valor à direita é o rótulo de
// destino já usado hoje na tabela (confirmado nos dados: quem é "Ajudante de Obra" já aparece
// como "Ajudante" via Função x SIENGE; "Tecnico Seguranca do Trabalho" já aparece como "Técnico
// de Segurança"; "Tecnico de Qualidade" aparece sem acento mesmo).
const CARGO_CONTA_COMO = {
  'Supervisor Segurança do Trabalho': 'Técnico de Segurança',
  'Auxiliar de Escritorio': 'Ajudante',
  'Analista de Qualidade': 'Tecnico de Qualidade',
  'Auxiliar Administrativo': 'Ajudante',
  // "Tecnico de Qualidade" e "Técnico de Qualidade" são o mesmo cargo, só com/sem acento na
  // planilha — as duas grafias já apontam pro mesmo "Tec. De Qualidade" no histograma
  // (CARGO_TO_HISTFUNC), confirmado pelo cliente (2026-08-26): junta na grafia sem acento.
  'Técnico de Qualidade': 'Tecnico de Qualidade',
  // "Op de Munck" e "Operador de Munck" são o mesmo cargo (confirmado pelo cliente, 2026-08-26) —
  // as duas já apontavam pro mesmo "Operado de Munk" no histograma, mas ficavam em linhas
  // separadas (uma delas sempre de fora do Previsto, por sorteio de ordem). Junta em "Op de Munck".
  'Operador de Munck': 'Op de Munck',
  // "Motorista de Caminhao" e "Motorista de Caminhão Basculante" são a mesma função (motorista de
  // caçamba), confirmado pelo cliente (2026-08-26) — os dois Cargos Agrupados já apontam pro
  // mesmo "Motorista de caçamba" no histograma (CARGO_TO_HISTFUNC), então juntam numa linha só em
  // vez de aparecerem separados disputando o mesmo Previsto.
  'Motorista de Caminhao': 'Motorista de Caçamba',
  'Motorista de Caminhão Basculante': 'Motorista de Caçamba',
};
function funcaoAggLabel(r){
  if(CARGO_CONTA_COMO[r.funcaoBase]) return CARGO_CONTA_COMO[r.funcaoBase];
  return (r.funcaoSienge && !SEM_FUNCAO_SIENGE.has(r.funcaoSienge)) ? r.funcaoSienge : r.funcaoBase;
}
function funcaoAgg(){
  const rows = filteredRows();
  const months = state.month.length===0 ? monthsPresent : state.month;
  const overlap = months.filter(m=>overlapMonths.includes(m));
  function previstoHistFunc(hf){
    return HIST_PREVISTO.filter(h=>h.funcaoBase===hf && overlap.includes(h.mesKey)).reduce((s,h)=>s+h.headcount*HOURS_PP,0);
  }

  const map = new Map();
  rows.forEach(r=>{
    const label = funcaoAggLabel(r);
    const key = label+'|'+r.local;
    if(!map.has(key)) map.set(key, {funcao:label, setor:r.setor, local:r.local, pessoasSet:new Set(), normais:0, extra:0, custo:0, cargosBase:new Set()});
    const a = map.get(key);
    a.pessoasSet.add(r.nome); a.normais+=r.horasMes; a.extra+=r.extraTotal; a.custo+=r.custoHoraExtra;
    a.cargosBase.add(r.funcaoBase);
  });
  const linhas = [...map.values()];

  // Previsto continua SEMPRE vindo do histograma — não muda de fonte só porque o rótulo virou
  // Função SIENGE. O problema: uma função do histograma (HIST_PREVISTO) pode ser "tocada" por
  // mais de um Cargo Agrupado/linha ao mesmo tempo, de dois jeitos — (a) split: um Cargo Agrupado
  // cobre uma LISTA de funções do histograma (CARGO_TO_HISTFUNC) e, pessoa a pessoa, cada uma virou
  // o rótulo (Função SIENGE) de uma linha diferente (ex: "Supervisor de Obras" → Supervisor Civil /
  // Supervisor de Montagem); (b) colisão: dois Cargos Agrupados DIFERENTES apontam pra MESMA função
  // do histograma (ex: "Almoxarife" e "Auxiliar de Almoxarifado", ou duas grafias do mesmo cargo
  // como "Tecnico"/"Técnico de Qualidade"). Sem tratar isso, a função do histograma seria somada
  // mais de uma vez e o Total desta tabela ficaria maior que o Previsto real do histograma (foi
  // exatamente o que o cliente reportou em 2026-08-26).
  //
  // Resolve com uma fila de reivindicações: cada (linha, Cargo Agrupado, função do histograma)
  // tenta reivindicar aquela função; só a PRIMEIRA da fila fica com o Previsto dela — as demais
  // ficam sem essa parcela (nunca em dobro, nunca inventando uma divisão que o histograma não dá
  // pra sustentar). Prioridade de quem entra primeiro na fila: (1) o rótulo da própria linha é
  // exatamente aquela função do histograma (split explícito por Função SIENGE — o caso mais
  // específico/confiável); (2) o Cargo Agrupado tem o MESMO nome da função do histograma (dono
  // "natural" da linha); (3) qualquer outro, na ordem em que aparece na planilha.
  const claims = [];
  linhas.forEach(a=>{
    if(a.local!=='Obra') return;
    a.cargosBase.forEach(cargo=>{
      histFuncList(cargo).forEach(hf=>{
        // Só reivindica se a função do histograma for do MESMO setor da linha — sem isso, um
        // cargo genérico com lista longa no CARGO_TO_HISTFUNC (ex: "Operador de Maquinas", que
        // cobre funções de Terraplanagem E de Equipamentos) puxava o Previsto inteiro das duas
        // pra dentro da própria linha, inflando o Previsto do setor onde a pessoa está classificada
        // (reportado pelo cliente em 2026-08-27: tabela de Terraplanagem somando 9.900h em vez dos
        // 5.500h do histograma). A função que sobrar sem dono cai como "vaga aberta" no setor certo
        // dela (bloco abaixo, "Vagas totalmente abertas").
        const histSetor = HIST_PREVISTO.find(h=>h.funcaoBase===hf)?.setor;
        if(histSetor !== a.setor) return;
        const prioridade = a.funcao===hf ? 0 : cargo===hf ? 1 : 2;
        claims.push({hf, rowKey:a.funcao+'|'+a.local, prioridade});
      });
    });
  });
  claims.sort((x,y)=>x.prioridade-y.prioridade);
  const donoDoHistFunc = new Map(); // histFunc -> rowKey vencedor
  claims.forEach(c=>{ if(!donoDoHistFunc.has(c.hf)) donoDoHistFunc.set(c.hf, c.rowKey); });

  const resultado = linhas.map(a=>{
    const total = a.normais+a.extra;
    let previsto = null;
    const hfsContados = new Set(); // uma função do histograma só conta 1x por linha, mesmo que
    // mais de um Cargo Agrupado deste grupo (ver CARGO_CONTA_COMO) aponte pra ela.
    a.cargosBase.forEach(cargo=>{
      const hfs = a.local==='Obra' ? histFuncList(cargo) : [];
      if(hfs.length>0){
        hfs.forEach(hf=>{
          if(hfsContados.has(hf)) return;
          if(donoDoHistFunc.get(hf) !== (a.funcao+'|'+a.local)) return; // outra linha já ficou com essa função
          hfsContados.add(hf);
          previsto = (previsto||0) + previstoHistFunc(hf);
        });
      } else {
        const p = previstoPorFuncao(cargo, months, a.local);
        if(p!=null) previsto = (previsto||0) + p;
      }
    });
    return {
      funcao:a.funcao, setor:a.setor, local:a.local, pessoas:a.pessoasSet.size,
      previsto, normais:a.normais, extra:a.extra, total, saldo: previsto==null ? null : previsto-total,
      pctExtra: total>0 ? a.extra/total*100 : 0, custo:a.custo,
    };
  });

  // Vagas totalmente abertas: função do histograma com Previsto mas ZERO gente real lançada em
  // nenhum Cargo Agrupado do filtro atual (não reivindicada por ninguém na fila acima) — mesmo
  // tratamento que "Custo por Função" (Financeiro) já dá a esse caso. Sem essa linha, o Total desta
  // tabela ficava menor que o Previsto oficial do histograma (a pedido do cliente, 2026-08-26, pra
  // bater com os cards do topo/tabela de Setor).
  if(state.local!=='Matriz'){
    const histFuncsAtivos = [...new Set(HIST_PREVISTO.filter(h=>overlap.includes(h.mesKey) && (state.setor.length===0||state.setor.includes(h.setor))).map(h=>h.funcaoBase))];
    histFuncsAtivos.forEach(hf=>{
      if(donoDoHistFunc.has(hf)) return;
      const p = previstoHistFunc(hf);
      if(p<=0) return;
      const histEntry = HIST_PREVISTO.find(h=>h.funcaoBase===hf);
      resultado.push({
        funcao:hf, setor: histEntry?histEntry.setor:'Não classificado', local:'Obra', pessoas:0,
        previsto:p, normais:0, extra:0, total:0, saldo:p, pctExtra:0, custo:0,
      });
    });
  }
  return resultado;
}

// ---------- Setor table (página Horas) ----------
function setorHorasAgg(){
  const rows = filteredRows();
  const months = state.month.length===0 ? monthsPresent : state.month;
  const map = new Map();
  rows.forEach(r=>{
    if(!map.has(r.setor)) map.set(r.setor, {setor:r.setor, pessoasSet:new Set(), normais:0, extra:0});
    const a = map.get(r.setor);
    a.pessoasSet.add(r.nome); a.normais+=r.horasMes; a.extra+=r.extraTotal;
  });
  const effectiveMonths = state.local==='Obra' ? months.filter(m=>overlapMonths.includes(m)) : months;
  return [...map.values()].map(a=>{
    const total = a.normais+a.extra;
    let previsto = 0;
    if(state.local!=='Matriz') previsto += HIST_PREVISTO.filter(h=>h.setor===a.setor && effectiveMonths.includes(h.mesKey)).reduce((s,h)=>s+h.headcount*HOURS_PP,0);
    if(state.local!=='Obra') previsto += months.reduce((s,mk)=>s+new Set(RAW.filter(r=>r.setor===a.setor && r.local==='Matriz' && r.mesKey===mk).map(r=>r.nome)).size*HOURS_PP,0);
    return {
      setor:a.setor, pessoas:a.pessoasSet.size, previsto, normais:a.normais, extra:a.extra, total,
      saldo: previsto-total, pctExtra: total>0 ? a.extra/total*100 : 0,
    };
  });
}

function renderSetorHorasTable(){
  const arr = sortArr(setorHorasAgg(), state.setorHorasSort.key, state.setorHorasSort.dir);
  const rowsHtml = arr.map(a=>`
    <tr>
      <td>${a.setor}</td>
      <td class="num">${fmtN(a.pessoas)}</td>
      <td class="num">${fmtH(a.previsto)}</td>
      <td class="num">${fmtH(a.normais)}</td>
      <td class="num">${fmtH(a.extra)}</td>
      <td class="num">${fmtH(a.total)}</td>
      <td class="num" style="color:${a.saldo<0?'var(--red)':'var(--green-dark)'}">${fmtH(a.saldo)}</td>
      <td class="num">${a.pctExtra.toFixed(1)}%</td>
    </tr>
  `).join('');
  const totalHtml = arr.length===0 ? '' : (()=>{
    const t = sumCols(arr, ['pessoas','previsto','normais','extra','total','saldo']);
    const pctExtra = t.total>0 ? t.extra/t.total*100 : 0; // recalculado (soma de % daria "média de médias" errada)
    return `<tr class="row-total">
      <td>Total</td>
      <td class="num">${fmtN(t.pessoas)}</td>
      <td class="num">${fmtH(t.previsto)}</td>
      <td class="num">${fmtH(t.normais)}</td>
      <td class="num">${fmtH(t.extra)}</td>
      <td class="num">${fmtH(t.total)}</td>
      <td class="num" style="color:${t.saldo<0?'var(--red)':'var(--green-dark)'}">${fmtH(t.saldo)}</td>
      <td class="num">${pctExtra.toFixed(1)}%</td>
    </tr>`;
  })();
  document.querySelector('#setorHorasTable tbody').innerHTML = (rowsHtml + totalHtml) || '<tr><td colspan="8" style="text-align:center;color:var(--muted);padding:20px;">Nenhum resultado</td></tr>';
  updateSortArrows('setorHorasTable', state.setorHorasSort);
}
wireSort('setorHorasTable', 'setorHorasSort', renderSetorHorasTable);

function renderFuncaoTable(){
  let arr = funcaoAgg();
  if(state.funcaoSearch){
    const q = state.funcaoSearch.toLowerCase();
    arr = arr.filter(a=>a.funcao.toLowerCase().includes(q) || a.setor.toLowerCase().includes(q));
  }
  arr = sortArr(arr, state.funcaoSort.key, state.funcaoSort.dir);
  const tbody = document.querySelector('#funcaoTable tbody');
  const rowsHtml = arr.map(a=>`
    <tr>
      <td>${a.funcao}</td>
      <td><span class="tag">${a.setor}</span></td>
      <td><span class="badge ${a.local==='Obra'?'obra':'matriz'}">${a.local}</span></td>
      <td class="num">${fmtN(a.pessoas)}</td>
      <td class="num">${a.previsto==null ? '—' : fmtH(a.previsto)}</td>
      <td class="num">${fmtH(a.normais)}</td>
      <td class="num">${fmtH(a.extra)}</td>
      <td class="num">${fmtH(a.total)}</td>
      <td class="num" style="color:${a.saldo==null?'inherit':a.saldo<0?'var(--red)':'var(--green-dark)'}">${a.saldo==null?'—':fmtH(a.saldo)}</td>
      <td class="num">${a.pctExtra.toFixed(1)}%</td>
    </tr>
  `).join('');
  const totalHtml = arr.length===0 ? '' : (()=>{
    const t = sumCols(arr, ['pessoas','previsto','normais','extra','total']);
    // Diferença = Previsto somado − Total somado (não a soma das diferenças linha a linha) —
    // algumas linhas ficam com Previsto "—" (função do histograma que outra linha já reivindicou,
    // ver funcaoAgg) mas o Total (Realizado) delas continua contando; somar só os saldo!=null
    // bateria menor que a conta feita "de olho" em cima dos dois totais ao lado (mesmo ajuste já
    // feito nas tabelas Sienge).
    const pctExtra = t.total>0 ? t.extra/t.total*100 : 0;
    const saldoTotal = t.previsto - t.total;
    return `<tr class="row-total">
      <td>Total</td>
      <td>—</td>
      <td>—</td>
      <td class="num">${fmtN(t.pessoas)}</td>
      <td class="num">${fmtH(t.previsto)}</td>
      <td class="num">${fmtH(t.normais)}</td>
      <td class="num">${fmtH(t.extra)}</td>
      <td class="num">${fmtH(t.total)}</td>
      <td class="num" style="color:${saldoTotal<0?'var(--red)':'var(--green-dark)'}">${fmtH(saldoTotal)}</td>
      <td class="num">${pctExtra.toFixed(1)}%</td>
    </tr>`;
  })();
  tbody.innerHTML = (rowsHtml + totalHtml) || '<tr><td colspan="10" style="text-align:center;color:var(--muted);padding:20px;">Nenhum resultado</td></tr>';
  updateSortArrows('funcaoTable', state.funcaoSort);
}
const funcaoSearchEl = document.getElementById('funcaoSearch');
if(funcaoSearchEl) funcaoSearchEl.addEventListener('input', e=>{ state.funcaoSearch = e.target.value; renderFuncaoTable(); });
wireSort('funcaoTable', 'funcaoSort', renderFuncaoTable);

// ================= PÁGINA: FINANCEIRO =================
// Projeção de custo até o fim da obra: realizado até agora (filtro atual) + histograma mês a mês
// dos meses que ainda vêm pela frente (depois do último mês com dado lançado) — segue o
// ramp-up/ramp-down real do plano, não uma média fixa repetida. Não projeta gasto pra meses que já
// passaram.
function projecaoCustoInfo(){
  const rows = filteredRows();
  const custoSalarioTotal = rows.reduce((s,r)=>s+r.salario,0);
  const custoExtra = rows.reduce((s,r)=>s+r.custoHoraExtra,0);
  const custoRealizado = custoSalarioTotal+custoExtra;
  const futuros = mesesFuturos();
  const custoProjetadoTotal = custoRealizado + custoPrevistoObraTotal(futuros) + custoMatrizFixoMensal()*futuros.length;
  // Previsto Até o Fim = orçamento do histograma pro projeto inteiro (Obra, todo o cronograma, +
  // Matriz fixo × todos os meses) — o valor orçado na planilha, fixo, não recalculado com o realizado.
  const previstoProjeto = custoPrevistoOrcadoTotal();
  return {custoProjetadoTotal, previstoProjeto, desvioProjecao: custoProjetadoTotal-previstoProjeto, mesesRestantes: futuros.length};
}

function renderKPIsFin(){
  const rows = filteredRows();
  const custoSalarioTotal = rows.reduce((s,r)=>s+r.salario,0);
  const custoExtra = rows.reduce((s,r)=>s+r.custoHoraExtra,0);
  const custoRealizado = custoSalarioTotal+custoExtra;
  const custoPrevisto = custoPrevistoTotalFiltro(); // A5: salário médio da função × pessoas previstas
  const pctExec = custoPrevisto>0 ? custoRealizado/custoPrevisto*100 : null;

  // A7: custo médio por mês do próprio filtro (custo realizado do período ÷ nº de meses no filtro).
  const mesesFiltro = state.month.length===0 ? monthsPresent.length : state.month.length;
  const custoMedioPorMes = mesesFiltro>0 ? custoRealizado/mesesFiltro : 0;

  const {custoProjetadoTotal, previstoProjeto, desvioProjecao, mesesRestantes} = projecaoCustoInfo();
  const pctProjecao = previstoProjeto>0 ? custoProjetadoTotal/previstoProjeto*100 : null;

  // Bloco do topo: mesma lógica "Previsto x Realizado" da página Horas (Previsto, Realizado,
  // Extra, Diferença, % Executado). Diferença = Previsto − Realizado (positivo = ainda sobra
  // orçamento; negativo = estourou) — mesmo sentido usado no card de Horas.
  const diferenca = custoPrevisto - custoRealizado;
  // Mesma decomposição da página Horas: quanto da diferença vem do Salário (Previsto − Salário
  // realizado) e quanto vem da Hora Extra (sem previsto, sempre um desconto puro do saldo).
  const diferencaSalario = custoPrevisto - custoSalarioTotal;
  const diferencaExtraParte = -custoExtra;
  // % Executado decomposto em pontos percentuais (mesma lógica da página Horas): quanto do desvio
  // (pctExec−100) vem do Salário sozinho passar do previsto e quanto vem da Hora Extra (sem
  // previsto próprio, sempre soma pontos ao desvio). Somam exatamente pctExec−100.
  const pctSalarioPP = custoPrevisto>0 ? (custoSalarioTotal-custoPrevisto)/custoPrevisto*100 : null;
  const pctExtraCustoPP = custoPrevisto>0 ? custoExtra/custoPrevisto*100 : null;
  const kpisTopo = [
    {label:'Custo Previsto', value: fmtR(custoPrevisto), color:'var(--accent)'},
    {label:'Custo Realizado', value: fmtR(custoRealizado), breakdown: [{label:'Salário', value:fmtR(custoSalarioTotal), color:'var(--green)'}, {label:'Extra', value:fmtR(custoExtra), color:'var(--red)'}], color:'var(--accent)'},
    {label:'Custo Hora Extra', value: fmtR(custoExtra), color:'var(--red)'},
    {label:'Diferença (Previsto − Realizado)', value: fmtR(Math.abs(diferenca)) + (diferenca>=0?' de folga':' estourado'), breakdown: [{label:'Salário', value:(diferencaSalario>=0?'+':'')+fmtR(diferencaSalario), color: diferencaSalario>=0?'var(--green)':'var(--red)'}, {label:'Extra', value:(diferencaExtraParte>=0?'+':'')+fmtR(diferencaExtraParte), color:'var(--red)'}], color: diferenca>=0 ? 'var(--green-dark)' : 'var(--red)'},
    {label:'% Executado', value: fmtPct(pctExec), sub: pctExec==null ? kpiSignal(pctExec) : null, breakdown: pctExec==null ? null : [{label:'Salário', value:(pctSalarioPP>=0?'+':'')+pctSalarioPP.toFixed(1)+'%', color: pctSalarioPP>=0?'var(--red)':'var(--green)'}, {label:'Extra', value:'+'+pctExtraCustoPP.toFixed(1)+'%', color:'var(--red)'}], color: pctExec==null ? 'var(--accent)' : pctExec<=100 ? 'var(--green-dark)' : 'var(--red)'},
  ];
  document.getElementById('kpisFin').innerHTML = kpisTopo.map(k=>`
    <div class="kpi"><div class="label">${k.label}</div><div class="value" style="color:${k.color}">${k.value}</div>${k.breakdown ? `<div class="kpi-breakdown">${k.breakdown.map(b=>`<span class="kpi-bd-item"><span class="dot" style="background:${b.color}"></span>${b.label} <b>${b.value}</b></span>`).join('')}</div>` : (k.sub ? `<div class="sub">${k.sub}</div>` : '')}</div>
  `).join('');

  renderKPIsProjecaoBlock('kpisFinProjecao', {custoMedioPorMes, mesesFiltro, custoProjetadoTotal, mesesRestantes, previstoProjeto, desvioProjecao, pctProjecao});
}
function renderKPIsProjecaoBlock(elId, v){
  const el = document.getElementById(elId);
  if(!el) return;
  const kpisProjecao = [
    {label:'Custo Médio por Mês', value: fmtR(v.custoMedioPorMes), color:'var(--accent)'},
    {label:'Projeção até Fim da Obra', value: fmtR(v.custoProjetadoTotal), color:'var(--accent)'},
    {label:'Previsto até o Fim', value: fmtR(v.previstoProjeto), color:'var(--accent)'},
    {label:'Diferença (Projeção x Previsto)', value: fmtR(v.desvioProjecao), color: v.desvioProjecao>0 ? 'var(--red)' : 'var(--green-dark)'},
    {label:'% do Previsto até o Fim', value: v.pctProjecao==null?'—':v.pctProjecao.toFixed(1)+'%', color: v.pctProjecao==null ? 'var(--accent)' : v.pctProjecao<=100 ? 'var(--green-dark)' : 'var(--red)'},
  ];
  el.innerHTML = kpisProjecao.map(k=>`
    <div class="kpi"><div class="label">${k.label}</div><div class="value" style="color:${k.color}">${k.value}</div>${k.breakdown ? `<div class="kpi-breakdown">${k.breakdown.map(b=>`<span class="kpi-bd-item"><span class="dot" style="background:${b.color}"></span>${b.label} <b>${b.value}</b></span>`).join('')}</div>` : (k.sub ? `<div class="sub">${k.sub}</div>` : '')}</div>
  `).join('');
}

function finMonthlyAgg(){
  return monthsPresent.map(mk=>{
    const rows = RAW.filter(r=>r.mesKey===mk &&
      (state.local==='all' || r.local===state.local) && (state.setor.length===0 || state.setor.includes(r.setor)));
    const salarioBase = rows.reduce((s,r)=>s+r.salario,0);
    return {
      mesKey:mk, mes: (rows[0]||{mes:mesLabel(mk)}).mes,
      salarioBase,
      extra: rows.reduce((s,r)=>s+r.custoHoraExtra,0),
      previsto: custoPrevistoParaMes(mk),
    };
  });
}

function renderFinChart(ids){
  ids = ids || {wrap:'finChartWrap', svg:'finChart', legend:'finChartLegend', totais:'finChartTotais'};
  const wrap = document.getElementById(ids.wrap);
  if(!wrap || wrap.clientWidth===0) return;
  const data = finMonthlyAgg();
  const svg = document.getElementById(ids.svg);
  const W = Math.max(wrap.clientWidth, ids.minWidth||560), H = ids.height||320;
  const marginL = ids.marginL||74, marginR = 20, marginT = 20, marginB = ids.height ? 26 : 40;
  const plotW = W - marginL - marginR, plotH = H - marginT - marginB;
  svg.setAttribute('viewBox', `0 0 ${W} ${H}`);
  svg.setAttribute('width', W); svg.setAttribute('height', H);
  svg.innerHTML = '';
  const ns = 'http://www.w3.org/2000/svg';
  const mk = (tag, attrs) => { const el = document.createElementNS(ns, tag); for(const k in attrs) el.setAttribute(k, attrs[k]); return el; };
  const tooltip = document.getElementById('tooltip');

  // Realizado x Previsto — mesmo padrão do gráfico de Horas: uma barra de Previsto e uma barra
  // empilhada de Realizado, lado a lado, por mês.
  const totals = data.map(d=>d.salarioBase+d.extra);
  const maxVal = Math.max(...totals, ...data.map(d=>d.previsto), 10) * 1.15;
  const groupW = plotW/data.length;
  const barW = Math.min(46, groupW*0.32);
  // A7: linha de referência com o custo médio realizado por mês do período (não pesa pelo filtro
  // Mês, já que aqui sempre olhamos todos os meses lançados lado a lado).
  const custoMedioMes = totals.length>0 ? totals.reduce((s,v)=>s+v,0)/totals.length : 0;

  const ticks = ids.height ? 3 : 5;
  for(let i=0;i<=ticks;i++){
    const val = maxVal/ticks*i;
    const y = marginT + plotH - (val/maxVal)*plotH;
    svg.appendChild(mk('line', {x1:marginL, x2:W-marginR, y1:y, y2:y, class:'gridline'}));
    const t = mk('text', {x:marginL-8, y:y+4, 'text-anchor':'end', class:'axis', style: ids.height?'font-size:9px;':''});
    t.textContent = val.toLocaleString('pt-BR', {notation:'compact', maximumFractionDigits:1});
    svg.appendChild(t);
  }
  svg.appendChild(mk('line', {x1:marginL, x2:marginL, y1:marginT, y2:marginT+plotH, class:'axis'}));

  function addTooltipHandlers(el, html){
    el.addEventListener('mousemove', e=>{ tooltip.style.opacity=1; tooltip.innerHTML = html; tooltip.style.left=(e.clientX+14)+'px'; tooltip.style.top=(e.clientY+10)+'px'; });
    el.addEventListener('mouseleave', ()=>tooltip.style.opacity=0);
  }

  const SEG_LABELS = {salarioBase:'Salário-base', extra:'Custo Hora Extra'};
  const segs = [['salarioBase','var(--green)'],['extra','var(--red)']];
  data.forEach((d,i)=>{
    const gx = marginL + groupW*i;
    const cx1 = gx + groupW/2 - barW - 4;
    const cx2 = gx + groupW/2 + 4;
    const lbl = mk('text', {x:gx+groupW/2, y:marginT+plotH+18, 'text-anchor':'middle', class:'axis', style: ids.height?'font-size:9px;':''});
    lbl.textContent = ids.height ? (d.mes||'').split('/')[0] : d.mes;
    svg.appendChild(lbl);

    // Rótulo de valor acima da barra — deixa o comparativo visível sem precisar passar o mouse.
    const valFontSize = ids.height ? 8 : 9.5;
    function valueLabel(x, y, texto, cor){
      const t = mk('text', {x, y:y-4, 'text-anchor':'middle', style:`font-size:${valFontSize}px;font-weight:700;fill:${cor||'var(--muted)'};`});
      t.textContent = texto;
      svg.appendChild(t);
    }

    // Barra 1: Previsto (azul, padrão do dashboard)
    if(d.previsto>0){
      const hPrev = (d.previsto/maxVal)*plotH;
      const rPrev = mk('rect', {x:cx1, y:marginT+plotH-hPrev, width:barW, height:hPrev, rx:4, fill:'var(--previsto)', class:'bar'});
      addTooltipHandlers(rPrev, `<strong>${d.mes} — Previsto</strong><br>${fmtR(d.previsto)}`);
      svg.appendChild(rPrev);
      if(!ids.height) valueLabel(cx1+barW/2, marginT+plotH-hPrev, fmtR(d.previsto), 'var(--previsto)');
    } else {
      const rNoPrev = mk('rect', {x:cx1, y:marginT+plotH-14, width:barW, height:14, rx:4, fill:'none', stroke:'var(--previsto)', 'stroke-width':1.5, 'stroke-dasharray':'3,2'});
      addTooltipHandlers(rNoPrev, `<strong>${d.mes} — Previsto</strong><br>Sem baseline do histograma nesse mês.`);
      svg.appendChild(rNoPrev);
    }

    // Barra 2: Realizado, empilhada (Salário-base + Custo Hora Extra)
    let yCursor = marginT+plotH;
    segs.forEach(([key,color])=>{
      const val = d[key];
      if(val<=0) return;
      const h = (val/maxVal)*plotH;
      yCursor -= h;
      const rSeg = mk('rect', {x:cx2, y:yCursor, width:barW, height:h, fill:color});
      addTooltipHandlers(rSeg, `<strong>${d.mes} — ${SEG_LABELS[key]}</strong><br>${fmtR(val)}<br><span style="opacity:.8;">Realizado: ${fmtR(d.salarioBase+d.extra)}</span>`);
      svg.appendChild(rSeg);
    });
    if(!ids.height && d.salarioBase+d.extra>0) valueLabel(cx2+barW/2, yCursor, fmtR(d.salarioBase+d.extra), 'var(--text)');
  });

  if(custoMedioMes>0 && custoMedioMes<=maxVal && !ids.height){
    const yRef = marginT + plotH - (custoMedioMes/maxVal)*plotH;
    svg.appendChild(mk('line', {x1:marginL, x2:W-marginR, y1:yRef, y2:yRef, stroke:'var(--amber)', 'stroke-width':1.5, 'stroke-dasharray':'5,3'}));
    const refLbl = mk('text', {x:W-marginR, y:yRef-5, 'text-anchor':'end', class:'axis', fill:'var(--amber)'});
    refLbl.textContent = `Média realizado: ${fmtR(custoMedioMes)}/mês`;
    svg.appendChild(refLbl);
  }

  document.getElementById(ids.legend).innerHTML = `<span><span class="dot" style="background:var(--previsto)"></span>Previsto</span><span><span class="dot" style="background:var(--green)"></span>Salário-base</span><span><span class="dot" style="background:var(--red)"></span>Custo Hora Extra</span>` + (ids.height ? '' : `<span><span class="dot" style="background:var(--amber)"></span>Custo Médio / Mês (linha tracejada)</span>`);

  if(ids.totais){
    const totalPrevisto = data.reduce((s,d)=>s+d.previsto,0);
    const totalRealizado = data.reduce((s,d)=>s+d.salarioBase+d.extra,0);
    const desvioChart = totalRealizado - totalPrevisto;
    const totaisEl = document.getElementById(ids.totais);
    const termoDesvio = ids.termoDesvio || 'Desvio';
    const corDesvio = desvioChart>0?'var(--red)':'var(--green-dark)';
    if(totaisEl) totaisEl.innerHTML = `<strong>Total do período (${data.length} ${data.length===1?'mês':'meses'}):</strong> Previsto <strong>${fmtR(totalPrevisto)}</strong> &middot; Realizado <strong>${fmtR(totalRealizado)}</strong> &middot; ${termoDesvio} <strong style="color:${corDesvio}">${fmtR(desvioChart)}</strong>`;
  }
}

// ---------- Projeção de custo até o fim da obra (Realizado x Projetado) ----------
// Projetado = histograma mês a mês dos meses que faltam (não mais uma média fixa repetida) —
// segue o ramp-up/ramp-down real do plano. "previsto" (por mês, pros dois lados da linha) é a
// mesma régua do histograma, calculada pra TODOS os meses (inclusive os já realizados), pra
// desenhar a curva de referência mês a mês por cima das barras.
function projTimeline(){
  const timeline = [...monthsPresent, ...mesesFuturos()];
  const custoMatrizMes = custoMatrizFixoMensal();
  // Meses do cronograma que ficam ANTES do início dos dados (ex: JAN/2026) não entram como barra
  // (não faz sentido "projetar" um mês que já passou sem nunca ter sido lançado), mas o orçamento
  // deles ainda conta pro total — senão o acumulado de Previsto desta curva fica menor que o
  // "Previsto até o Fim" do KPI, que soma o cronograma inteiro. Começa o acumulado já com esse
  // pedaço somado, pra bater exato no fim.
  const mesesForaTimeline = PROJECT_MONTHS.filter(m=>!timeline.includes(m));
  let cum = 0;
  let cumPrevisto = mesesForaTimeline.reduce((s,mk)=>s+custoPrevistoObraTotal([mk])+custoMatrizMes, 0);
  return timeline.map(mk=>{
    const realizado = monthsPresent.includes(mk);
    const previsto = custoPrevistoObraTotal([mk]) + custoMatrizMes;
    const custo = realizado
      ? RAW.filter(r=>r.mesKey===mk && (state.local==='all'||r.local===state.local) && (state.setor.length===0||state.setor.includes(r.setor)))
          .reduce((s,r)=>s+r.salario+r.custoHoraExtra,0)
      : previsto;
    cum += custo;
    cumPrevisto += previsto;
    return {mesKey:mk, mes:mesLabel(mk), custo, cum, realizado, previsto, cumPrevisto, diferenca: cum-cumPrevisto};
  });
}

// ---------- Curva S: Previsto (histograma) x Realizado+Projeção, acumulados ao longo do tempo.
// Duas curvas cumulativas (o formato de "S" vem do ramp-up/pico/ramp-down do histograma) com a
// área entre elas sombreada — o desvio em qualquer ponto do tempo é a distância vertical entre
// as duas linhas, não um número só no fim.
function renderCurvaSChart(ids){
  ids = ids || {wrap:'projDiffChartWrap', svg:'projDiffChart', legend:'projDiffChartLegend', footnote:'projDiffChartFootnote'};
  const wrap = document.getElementById(ids.wrap);
  if(!wrap || wrap.clientWidth===0) return;
  const data = projTimeline();
  const svg = document.getElementById(ids.svg);
  // Avanço Físico (curva do Planejamento, linha "Construção") só entra se o Excel externo tiver dado
  // pro período mostrado.
  const hasAvanco = data.some(d=>AVANCO_FISICO_PLAN.has(d.mesKey) || AVANCO_FISICO_REAL.has(d.mesKey));
  const W = Math.max(wrap.clientWidth, ids.minWidth||900), H = ids.height||300;
  const marginL = ids.marginL||84, marginR = ids.height?36:52, marginT = 20, marginB = ids.height ? 26 : 40;
  const plotW = W - marginL - marginR, plotH = H - marginT - marginB;
  svg.setAttribute('viewBox', `0 0 ${W} ${H}`);
  svg.setAttribute('width', W); svg.setAttribute('height', H);
  svg.innerHTML = '';
  const ns = 'http://www.w3.org/2000/svg';
  const mk = (tag, attrs) => { const el = document.createElementNS(ns, tag); for(const k in attrs) el.setAttribute(k, attrs[k]); return el; };
  const tooltip = document.getElementById('tooltip');
  if(data.length===0) return;

  // Uma régua vertical só, em % do Previsto total do projeto (não R$ à esquerda e % à direita em
  // escalas independentes) — assim o Custo e o Avanço Físico ficam na MESMA altura visual quando
  // estão no mesmo %, e uma curva "acima de 100%" realmente aparece mais alta que uma de 90%. O
  // eixo esquerdo só traduz cada marca pro R$ equivalente; o direito mostra a marca em %.
  const totalPrevistoFinal = data[data.length-1].cumPrevisto;
  const pctFor = v => totalPrevistoFinal>0 ? (v/totalPrevistoFinal*100) : 0;
  const pctsTodos = [
    ...data.map(d=>pctFor(d.cum)), ...data.map(d=>pctFor(d.cumPrevisto)),
    ...Array.from(AVANCO_FISICO_PLAN.values()), ...Array.from(AVANCO_FISICO_REAL.values())
  ];
  const maxPct = Math.max(100, ...pctsTodos) * 1.08;
  const yForPct = pct => marginT + plotH - (pct/maxPct)*plotH;
  const yFor = v => yForPct(pctFor(v));
  const groupW = plotW/data.length;
  const xFor = i => marginL + groupW*i + groupW/2;

  const ticks = 5;
  for(let i=0;i<=ticks;i++){
    const pctTick = maxPct/ticks*i;
    const y = yForPct(pctTick);
    svg.appendChild(mk('line', {x1:marginL, x2:W-marginR, y1:y, y2:y, class:'gridline'}));
    const valR = pctTick/100*totalPrevistoFinal;
    const t = mk('text', {x:marginL-8, y:y+4, 'text-anchor':'end', class:'axis', style: ids.height?'font-size:9px;':''});
    t.textContent = valR.toLocaleString('pt-BR', {notation:'compact', maximumFractionDigits:1});
    svg.appendChild(t);
    const t2 = mk('text', {x:W-marginR+6, y:y+4, 'text-anchor':'start', class:'axis', style: ids.height?'font-size:9px;':''});
    t2.textContent = pctTick.toFixed(0)+'%';
    svg.appendChild(t2);
  }
  svg.appendChild(mk('line', {x1:marginL, x2:marginL, y1:marginT, y2:marginT+plotH, class:'axis'}));
  svg.appendChild(mk('line', {x1:W-marginR, x2:W-marginR, y1:marginT, y2:marginT+plotH, class:'axis'}));

  const pCum = data.map((d,i)=>({x:xFor(i), y:yFor(d.cum), d}));
  const pPrev = data.map((d,i)=>({x:xFor(i), y:yFor(d.cumPrevisto), d}));

  // Rótulo do mês só cabe se tiver espaço — em vez de um limiar fixo de quantidade de meses,
  // calcula quantos rótulos cabem no espaço real disponível (evita sobreposição em cards
  // estreitos, tipo a Visão BI, sem tirar rótulos à toa nas páginas largas). Fonte menor no modo
  // compacto deixa mais rótulos caberem sem embolar.
  const labelFontSize = ids.height ? 8.5 : 10.5;
  const estLabelW = ids.height ? 22 : 32;
  const maxLabels = Math.max(1, Math.floor(plotW/estLabelW));
  const labelSkip = Math.max(1, Math.ceil(data.length/maxLabels));
  data.forEach((d,i)=>{
    if(i % labelSkip === 0 || i===data.length-1){
      const lbl = mk('text', {x:xFor(i), y:marginT+plotH+18, 'text-anchor':'middle', class:'axis', style:`font-size:${labelFontSize}px;`});
      lbl.textContent = d.mes.replace('/20','/');
      svg.appendChild(lbl);
    }
  });

  // Área sombreada entre as duas curvas, um trapézio por par de meses consecutivos — vermelho
  // (cor do Executado) = trecho acima do previsto, azul (cor do Previsto) = trecho abaixo.
  for(let i=0;i<data.length-1;i++){
    const corTrecho = (data[i].diferenca + data[i+1].diferenca)/2 > 0 ? 'var(--executado)' : 'var(--previsto)';
    const poly = mk('polygon', {
      points: `${pCum[i].x},${pCum[i].y} ${pCum[i+1].x},${pCum[i+1].y} ${pPrev[i+1].x},${pPrev[i+1].y} ${pPrev[i].x},${pPrev[i].y}`,
      fill: corTrecho, 'fill-opacity':'0.12'
    });
    svg.appendChild(poly);
  }

  // Avanço Físico (curva do Planejamento, linha "Construção") — mesma régua da
  // Curva S de custo (0-100%+), pra comparar se a obra anda no mesmo ritmo do gasto de verdade.
  // Cores diferentes de Previsto/Realizado (azul/vermelho já usados pela curva de custo) pra não
  // confundir as duas curvas: âmbar = Avanço Físico Planejado, violeta = Avanço Físico Realizado.
  if(hasAvanco){
    const corFisPlan = 'var(--amber)', corFisReal = 'var(--notTot)';
    const pFisPlan = data.map((d,i)=>({x:xFor(i), y:yForPct(AVANCO_FISICO_PLAN.get(d.mesKey)), d, pct:AVANCO_FISICO_PLAN.get(d.mesKey)})).filter(p=>p.pct!=null);
    const pFisReal = data.map((d,i)=>({x:xFor(i), y:yForPct(AVANCO_FISICO_REAL.get(d.mesKey)), d, pct:AVANCO_FISICO_REAL.get(d.mesKey)})).filter(p=>p.pct!=null);
    if(pFisPlan.length>1) svg.appendChild(mk('polyline', {points: pFisPlan.map(p=>`${p.x},${p.y}`).join(' '), fill:'none', stroke:corFisPlan, 'stroke-width':1.5, 'stroke-dasharray':'2,3'}));
    if(pFisReal.length>1) svg.appendChild(mk('polyline', {points: pFisReal.map(p=>`${p.x},${p.y}`).join(' '), fill:'none', stroke:corFisReal, 'stroke-width':1.5}));
    const addFisDots = (pontos, tipo, cor) => pontos.forEach(p=>{
      const dot = mk('circle', {cx:p.x, cy:p.y, r:2.5, fill:cor});
      dot.addEventListener('mousemove', e=>{
        tooltip.style.opacity=1;
        const custoTxt = `<br><span style="opacity:.85;">Curva S de Custo &mdash; devia ter gasto ${pctFor(p.d.cumPrevisto).toFixed(1)}% &middot; gastou ${pctFor(p.d.cum).toFixed(1)}%</span>`;
        tooltip.innerHTML = `<strong>${p.d.mes} — Avanço Físico ${tipo}</strong><br>${p.pct.toFixed(1)}%${custoTxt}`;
        tooltip.style.left=(e.clientX+14)+'px'; tooltip.style.top=(e.clientY+10)+'px';
      });
      dot.addEventListener('mouseleave', ()=>tooltip.style.opacity=0);
      svg.appendChild(dot);
    });
    addFisDots(pFisPlan, 'Planejado', corFisPlan);
    addFisDots(pFisReal, 'Realizado', corFisReal);
  }

  // Linha Previsto (histograma), tracejada, do início ao fim — azul (padrão do dashboard).
  svg.appendChild(mk('polyline', {points: pPrev.map(p=>`${p.x},${p.y}`).join(' '), fill:'none', stroke:'var(--previsto)', 'stroke-width':2.5, 'stroke-dasharray':'6,3'}));
  // Linha Realizado+Projeção: sólida até o último mês com dado, tracejada dali em diante — vermelho.
  const idxUltimoReal = data.reduce((acc,d,i)=>d.realizado?i:acc, -1);
  if(idxUltimoReal>=0){
    svg.appendChild(mk('polyline', {points: pCum.slice(0,idxUltimoReal+1).map(p=>`${p.x},${p.y}`).join(' '), fill:'none', stroke:'var(--executado)', 'stroke-width':2.5}));
  }
  if(idxUltimoReal < data.length-1){
    svg.appendChild(mk('polyline', {points: pCum.slice(Math.max(idxUltimoReal,0)).map(p=>`${p.x},${p.y}`).join(' '), fill:'none', stroke:'var(--executado)', 'stroke-width':2.5, 'stroke-dasharray':'6,3'}));
  }

  // Tooltip no formato "devia ter gasto X% e gastamos Y%" — mesma régua de % usada no Avanço
  // Físico, pra dar pra comparar as duas coisas igual (pedido do cliente). As duas linhas
  // (Previsto e Realizado) mostram o MESMO conteúdo completo, não só o próprio lado, pra não
  // precisar passar o mouse duas vezes.
  const addDots = (pontos, corFn) => pontos.forEach(p=>{
    const dot = mk('circle', {cx:p.x, cy:p.y, r:4, fill:corFn(p.d)});
    dot.addEventListener('mousemove', e=>{
      tooltip.style.opacity=1;
      const pctPrevisto = pctFor(p.d.cumPrevisto);
      const pctReal = pctFor(p.d.cum);
      const deltaPct = pctReal - pctPrevisto;
      const cor = deltaPct>0 ? 'var(--red)' : 'var(--green-dark)';
      const fisPlan = AVANCO_FISICO_PLAN.get(p.d.mesKey);
      const fisReal = AVANCO_FISICO_REAL.get(p.d.mesKey);
      const fisTxt = (fisPlan!=null || fisReal!=null)
        ? `<br><span style="opacity:.85;">Avanço Físico &mdash; Planejado: ${fisPlan!=null?fisPlan.toFixed(1)+'%':'—'} &middot; Realizado: ${fisReal!=null?fisReal.toFixed(1)+'%':'—'}</span>`
        : '';
      tooltip.innerHTML = `<strong>${p.d.mes} — Curva S de Custo</strong><br>Devia ter gasto: <strong>${pctPrevisto.toFixed(1)}%</strong> (${fmtR(p.d.cumPrevisto)})<br>Gastou: <strong>${pctReal.toFixed(1)}%</strong> (${fmtR(p.d.cum)}) <strong style="color:${cor}">(${deltaPct>=0?'+':''}${deltaPct.toFixed(1)}pp)</strong>${fisTxt}`;
      tooltip.style.left=(e.clientX+14)+'px'; tooltip.style.top=(e.clientY+10)+'px';
    });
    dot.addEventListener('mouseleave', ()=>tooltip.style.opacity=0);
    svg.appendChild(dot);
  });
  addDots(pPrev, ()=>'var(--previsto)');
  addDots(pCum, ()=>'var(--executado)');

  // % de avanço financeiro em cada ponto — pedido do cliente pra comparar com o Avanço Físico
  // direto no olho, sem precisar passar o mouse ponto a ponto. Previsto à esquerda do próprio
  // ponto, Realizado à direita — deslocamento horizontal (não só vertical) pra não embolar quando
  // as duas linhas estão coladas ou quando o Realizado cruza o Previsto. Frequência é a METADE
  // dos rótulos de mês (não a mesma) — com um rótulo de % em cada mês, a régua ficava poluída
  // demais (números um em cima do outro nos trechos em que Previsto e Realizado andam próximos).
  const pctLblSize = 8.5;
  const pctLabelSkip = labelSkip * 2;
  data.forEach((d,i)=>{
    if(ids.height) return; // card compacto da Visão BI fica sem rótulo de % pra não embolar
    if(!(i % pctLabelSkip === 0 || i===data.length-1)) return;
    const tp = mk('text', {x: pPrev[i].x-6, y: pPrev[i].y-6, 'text-anchor':'end', style:`font-size:${pctLblSize}px;font-weight:700;fill:var(--previsto);`});
    tp.textContent = pctFor(d.cumPrevisto).toFixed(0)+'%';
    svg.appendChild(tp);
    const tr = mk('text', {x: pCum[i].x+6, y: pCum[i].y-6, 'text-anchor':'start', style:`font-size:${pctLblSize}px;font-weight:700;fill:var(--executado);`});
    tr.textContent = pctFor(d.cum).toFixed(0)+'%';
    svg.appendChild(tr);
  });

  if(ids.legend){
    document.getElementById(ids.legend).innerHTML = `<span><span class="dot" style="background:var(--previsto)"></span>Previsto (histograma)</span><span><span class="dot" style="background:var(--executado)"></span>Realizado / Projeção</span><span><span class="dot" style="background:var(--executado);opacity:.4;"></span>Acima do previsto</span><span><span class="dot" style="background:var(--previsto);opacity:.4;"></span>Abaixo do previsto</span>` + (hasAvanco ? `<span><span class="dot" style="background:var(--amber)"></span>Avanço Físico Planejado (linha fina pontilhada)</span><span><span class="dot" style="background:var(--notTot)"></span>Avanço Físico Realizado (linha fina)</span>` : '');
  }
  if(ids.footnote){
    const ultimo = data[data.length-1];
    const cor = ultimo.diferenca>0 ? 'var(--red)' : 'var(--green-dark)';
    const pct = ultimo.cumPrevisto>0 ? (ultimo.cum/ultimo.cumPrevisto*100).toFixed(1)+'%' : '—';
    document.getElementById(ids.footnote).innerHTML = `No fim do projeto: desvio de <strong style="color:${cor}">${fmtR(ultimo.diferenca)}</strong> (${pct} do previsto).`;
  }
}

// Gráfico "Avanço Físico" — só as duas linhas de Avanço Físico (Planejado x Realizado,
// âmbar/violeta), sem a Curva S de custo/horas por trás (pedido do cliente pra simplificar:
// primeiro na Visão BI, depois na Início também). Régua própria em % do Avanço Físico — não
// compartilha escala com custo/horas. ids.height presente = modo compacto (card da Visão BI);
// ausente = tamanho cheio (card da Início, mesmo padrão visual dos outros gráficos da página).
function renderBIAvancoFisicoChart(ids){
  ids = ids || {wrap:'biCurvaWrap', svg:'biCurvaChart', legend:'biCurvaLegend', height:142, minWidth:240, marginL:34};
  const wrap = document.getElementById(ids.wrap);
  if(!wrap || wrap.clientWidth===0) return;
  // Só até o último mês com hora lançada (monthsPresent) — não o cronograma inteiro do projeto
  // (PROJECT_MONTHS/projTimeline). O Avanço Físico externo tem dado planejado/realizado pra bem
  // mais meses do que isso, mas mostrar além do que já foi apontado não faz sentido aqui: esse
  // gráfico não tem mais Previsto/Projeção pra comparar contra (removido a pedido do cliente),
  // então o horizonte dele é sempre "até onde já temos horas", nunca o futuro do cronograma.
  const data = monthsPresent.map(mk=>({mesKey:mk, mes:mesLabel(mk)}));
  const svg = document.getElementById(ids.svg);
  const compact = !!ids.height;
  const hasAvanco = data.some(d=>AVANCO_FISICO_PLAN.has(d.mesKey) || AVANCO_FISICO_REAL.has(d.mesKey));
  const W = Math.max(wrap.clientWidth, ids.minWidth||900), H = ids.height||300;
  const marginL = ids.marginL||(compact?34:84), marginR = compact?14:52, marginT = compact?10:20, marginB = compact?22:40;
  const plotW = W - marginL - marginR, plotH = H - marginT - marginB;
  svg.setAttribute('viewBox', `0 0 ${W} ${H}`);
  svg.setAttribute('width', W); svg.setAttribute('height', H);
  svg.innerHTML = '';
  if(ids.legend) document.getElementById(ids.legend).innerHTML = '';
  if(!hasAvanco || data.length===0) return;
  const ns = 'http://www.w3.org/2000/svg';
  const mk = (tag, attrs) => { const el = document.createElementNS(ns, tag); for(const k in attrs) el.setAttribute(k, attrs[k]); return el; };
  const tooltip = document.getElementById('tooltip');
  const gridFontSize = compact?9:11, labelFontSize = compact?8.5:10.5, dotR = compact?2.5:4, strokeW = compact?2:2.5;

  // Escala do eixo Y calculada só com os meses realmente exibidos (data), não o Avanço Físico
  // inteiro do arquivo externo (que cobre muito mais meses que isso) — senão o eixo ficaria com
  // um espaço vazio enorme no topo, escalado pra um % que só aparece bem lá na frente.
  const pctsTodos = data.flatMap(d=>[AVANCO_FISICO_PLAN.get(d.mesKey), AVANCO_FISICO_REAL.get(d.mesKey)]).filter(v=>v!=null);
  const maxPct = Math.max(100, ...pctsTodos) * 1.08;
  const yForPct = pct => marginT + plotH - (pct/maxPct)*plotH;
  const groupW = plotW/data.length;
  const xFor = i => marginL + groupW*i + groupW/2;

  const ticks = compact?4:5;
  for(let i=0;i<=ticks;i++){
    const pctTick = maxPct/ticks*i;
    const y = yForPct(pctTick);
    svg.appendChild(mk('line', {x1:marginL, x2:W-marginR, y1:y, y2:y, class:'gridline'}));
    const t = mk('text', {x:marginL-6, y:y+4, 'text-anchor':'end', class:'axis', style:`font-size:${gridFontSize}px;`});
    t.textContent = pctTick.toFixed(0)+'%';
    svg.appendChild(t);
  }
  svg.appendChild(mk('line', {x1:marginL, x2:marginL, y1:marginT, y2:marginT+plotH, class:'axis'}));

  const estLabelW = compact?22:32;
  const maxLabels = Math.max(1, Math.floor(plotW/estLabelW));
  const labelSkip = Math.max(1, Math.ceil(data.length/maxLabels));
  data.forEach((d,i)=>{
    if(i % labelSkip === 0 || i===data.length-1){
      const lbl = mk('text', {x:xFor(i), y:marginT+plotH+(compact?15:18), 'text-anchor':'middle', class:'axis', style:`font-size:${labelFontSize}px;`});
      lbl.textContent = d.mes.replace('/20','/');
      svg.appendChild(lbl);
    }
  });

  const corFisPlan = 'var(--amber)', corFisReal = 'var(--notTot)';
  const pFisPlan = data.map((d,i)=>({x:xFor(i), y:yForPct(AVANCO_FISICO_PLAN.get(d.mesKey)), d, pct:AVANCO_FISICO_PLAN.get(d.mesKey)})).filter(p=>p.pct!=null);
  const pFisReal = data.map((d,i)=>({x:xFor(i), y:yForPct(AVANCO_FISICO_REAL.get(d.mesKey)), d, pct:AVANCO_FISICO_REAL.get(d.mesKey)})).filter(p=>p.pct!=null);
  if(pFisPlan.length>1) svg.appendChild(mk('polyline', {points: pFisPlan.map(p=>`${p.x},${p.y}`).join(' '), fill:'none', stroke:corFisPlan, 'stroke-width':strokeW, 'stroke-dasharray':'4,3'}));
  if(pFisReal.length>1) svg.appendChild(mk('polyline', {points: pFisReal.map(p=>`${p.x},${p.y}`).join(' '), fill:'none', stroke:corFisReal, 'stroke-width':strokeW}));
  const addFisDots = (pontos, tipo, cor) => pontos.forEach(p=>{
    const dot = mk('circle', {cx:p.x, cy:p.y, r:dotR, fill:cor});
    dot.addEventListener('mousemove', e=>{
      tooltip.style.opacity=1;
      tooltip.innerHTML = `<strong>${p.d.mes} — Avanço Físico ${tipo}</strong><br>${p.pct.toFixed(1)}%`;
      tooltip.style.left=(e.clientX+14)+'px'; tooltip.style.top=(e.clientY+10)+'px';
    });
    dot.addEventListener('mouseleave', ()=>tooltip.style.opacity=0);
    svg.appendChild(dot);
  });
  addFisDots(pFisPlan, 'Planejado', corFisPlan);
  addFisDots(pFisReal, 'Realizado', corFisReal);

  if(ids.legend){
    document.getElementById(ids.legend).innerHTML = `<span><span class="dot" style="background:${corFisPlan}"></span>Planejado</span><span><span class="dot" style="background:${corFisReal}"></span>Realizado</span>`;
  }
}

function renderProjChart(ids){
  ids = ids || {wrap:'projChartWrap', svg:'projChart', legend:'projChartLegend', footnote:'projChartFootnote'};
  const wrap = document.getElementById(ids.wrap);
  if(!wrap || wrap.clientWidth===0) return;
  const data = projTimeline();
  const svg = document.getElementById(ids.svg);
  const W = Math.max(wrap.clientWidth, 900), H = 300;
  const marginL = 74, marginR = 20, marginT = 20, marginB = 40;
  const plotW = W - marginL - marginR, plotH = H - marginT - marginB;
  svg.setAttribute('viewBox', `0 0 ${W} ${H}`);
  svg.setAttribute('width', W); svg.setAttribute('height', H);
  svg.innerHTML = '';
  const ns = 'http://www.w3.org/2000/svg';
  const mk = (tag, attrs) => { const el = document.createElementNS(ns, tag); for(const k in attrs) el.setAttribute(k, attrs[k]); return el; };
  const tooltip = document.getElementById('tooltip');

  const maxVal = Math.max(...data.map(d=>Math.max(d.custo, d.previsto||0)), 10) * 1.15;
  const groupW = plotW/data.length;
  const barW = Math.min(40, groupW*0.6);

  const ticks = 5;
  for(let i=0;i<=ticks;i++){
    const val = maxVal/ticks*i;
    const y = marginT + plotH - (val/maxVal)*plotH;
    svg.appendChild(mk('line', {x1:marginL, x2:W-marginR, y1:y, y2:y, class:'gridline'}));
    const t = mk('text', {x:marginL-8, y:y+4, 'text-anchor':'end', class:'axis'});
    t.textContent = val.toLocaleString('pt-BR', {notation:'compact', maximumFractionDigits:1});
    svg.appendChild(t);
  }
  svg.appendChild(mk('line', {x1:marginL, x2:marginL, y1:marginT, y2:marginT+plotH, class:'axis'}));

  const estLabelW = 32;
  const maxLabels = Math.max(1, Math.floor(plotW/estLabelW));
  const labelSkip = Math.max(1, Math.ceil(data.length/maxLabels));
  data.forEach((d,i)=>{
    const gx = marginL + groupW*i + groupW/2 - barW/2;
    if(i % labelSkip === 0 || i===data.length-1){
      const lbl = mk('text', {x:gx+barW/2, y:marginT+plotH+18, 'text-anchor':'middle', class:'axis'});
      lbl.textContent = d.mes.replace('/20','/');
      svg.appendChild(lbl);
    }
    const h = (d.custo/maxVal)*plotH;
    const rect = mk('rect', {x:gx, y:marginT+plotH-h, width:barW, height:h, rx:3,
      fill: 'var(--executado)', 'fill-opacity': d.realizado ? '1' : '0.32',
      stroke: d.realizado ? 'none' : 'var(--executado)', 'stroke-width': d.realizado?0:1.5, 'stroke-dasharray': d.realizado?'none':'3,2'});
    rect.addEventListener('mousemove', e=>{ tooltip.style.opacity=1; tooltip.innerHTML = `<strong>${d.mes} — ${d.realizado?'Realizado':'Projetado'}</strong><br>${fmtR(d.custo)}<br>Acumulado: ${fmtR(d.cum)}`; tooltip.style.left=(e.clientX+14)+'px'; tooltip.style.top=(e.clientY+10)+'px'; });
    rect.addEventListener('mouseleave', ()=>tooltip.style.opacity=0);
    svg.appendChild(rect);
    // Rótulo de valor só nos meses que já têm rótulo de mês visível (mesmo espaçamento) — com 17
    // meses num timeline só, rótulo em toda barra ia embolar tudo.
    if(d.custo>0 && (i % labelSkip === 0 || i===data.length-1)){
      const tv = mk('text', {x:gx+barW/2, y:marginT+plotH-h-4, 'text-anchor':'middle', style:'font-size:9px;font-weight:700;fill:var(--text);'});
      tv.textContent = d.custo.toLocaleString('pt-BR',{notation:'compact',maximumFractionDigits:1});
      svg.appendChild(tv);
    }
  });

  // Curva de referência: Previsto (histograma) mês a mês, por cima das barras — mesma escala,
  // pra comparar visualmente cada mês de Realizado/Projetado contra o que estava planejado.
  const pontos = data.map((d,i)=>{
    const gx = marginL + groupW*i + groupW/2;
    const y = marginT + plotH - (d.previsto/maxVal)*plotH;
    return {x:gx, y, d};
  });
  const polyline = mk('polyline', {points: pontos.map(p=>`${p.x},${p.y}`).join(' '), fill:'none', stroke:'var(--previsto)', 'stroke-width':2, 'stroke-dasharray':'5,3'});
  svg.appendChild(polyline);
  pontos.forEach(p=>{
    const dot = mk('circle', {cx:p.x, cy:p.y, r:4, fill:'var(--previsto)'});
    dot.addEventListener('mousemove', e=>{ tooltip.style.opacity=1; tooltip.innerHTML = `<strong>${p.d.mes} — Previsto (histograma)</strong><br>${fmtR(p.d.previsto)}`; tooltip.style.left=(e.clientX+14)+'px'; tooltip.style.top=(e.clientY+10)+'px'; });
    dot.addEventListener('mouseleave', ()=>tooltip.style.opacity=0);
    svg.appendChild(dot);
  });

  document.getElementById(ids.legend).innerHTML = `<span><span class="dot" style="background:var(--executado)"></span>Realizado</span><span><span class="dot" style="background:var(--executado);opacity:.4;"></span>Projetado</span><span><span class="dot" style="background:var(--previsto)"></span>Previsto (histograma, linha)</span>`;
  if(ids.footnote){
    const totalProjetado = data.length>0 ? data[data.length-1].cum : 0;
    const previstoProjeto = custoPrevistoOrcadoTotal();
    const desvioProjecao = totalProjetado - previstoProjeto;
    const pctProjecao = previstoProjeto>0 ? totalProjetado/previstoProjeto*100 : null;
    document.getElementById(ids.footnote).innerHTML = `<strong>Projetado: ${fmtR(totalProjetado)} &middot; Previsto até o fim: ${fmtR(previstoProjeto)} &middot; Diferença: <span style="color:${desvioProjecao>0?'var(--red)':'var(--green-dark)'}">${fmtR(desvioProjecao)}</span> &middot; ${pctProjecao==null?'—':pctProjecao.toFixed(1)+'%'} do previsto</strong>`;
  }
}

function renderProjTable(){
  const data = projTimeline();
  document.querySelector('#projTable tbody').innerHTML = data.map(d=>`
    <tr>
      <td>${d.mes}</td>
      <td>${d.realizado
        ? `<span class="badge" style="background:color-mix(in srgb, var(--green-dark) 16%, transparent);color:var(--green-dark);">Realizado</span>`
        : `<span class="badge" style="background:color-mix(in srgb, var(--accent) 15%, transparent);color:var(--accent);">Projetado</span>`}</td>
      <td class="num">${fmtR(d.custo)}</td>
      <td class="num">${fmtR(d.cum)}</td>
    </tr>
  `).join('');
}

function finSetorAgg(){
  const rows = filteredRows();
  const map = new Map();
  rows.forEach(r=>{
    if(!map.has(r.setor)) map.set(r.setor, {setor:r.setor, pessoasSet:new Set(), pessoaMeses:0, custoEfetivo:0, custoExtra:0, horas:0});
    const a = map.get(r.setor);
    a.pessoasSet.add(r.nome); a.pessoaMeses++; a.custoEfetivo+=r.salario; a.custoExtra+=r.custoHoraExtra;
    a.horas += r.horasMes+r.extraTotal;
  });
  const custoPrevistoMap = custoPrevistoAgrupado(r=>r.setor);
  // Setor com Previsto (histograma) mas ZERO gente real no filtro atual — sem isso a soma da
  // tabela não bate com o KPI "Custo Previsto" do topo.
  custoPrevistoMap.forEach((valor, setor)=>{
    if(!map.has(setor) && valor>0) map.set(setor, {setor, pessoasSet:new Set(), pessoaMeses:0, custoEfetivo:0, custoExtra:0, horas:0});
  });
  const arr = [...map.values()].map(a=>{
    const custoTotal = a.custoEfetivo+a.custoExtra;
    return {
      setor:a.setor, pessoas:a.pessoasSet.size, custoPrevisto: custoPrevistoMap.get(a.setor)||0,
      custoEfetivo:a.custoEfetivo, custoExtra:a.custoExtra, custoTotal,
      pessoaMeses:a.pessoaMeses, horas:a.horas,
      // A8: custo médio por pessoa é ponderado por pessoa-mês (contagem de linhas), não por
      // nome distinto do período — senão quem trabalhou 1 mês pesa igual a quem trabalhou 5.
      custoPP: a.pessoaMeses>0 ? custoTotal/a.pessoaMeses : 0,
      custoPorHora: a.horas>0 ? custoTotal/a.horas : 0,
    };
  });
  const grand = arr.reduce((s,a)=>s+a.custoTotal,0);
  arr.forEach(a=>a.pctTotal = grand>0 ? a.custoTotal/grand*100 : 0);
  return arr;
}

function renderFinSetorTable(){
  const arr = sortArr(finSetorAgg(), state.setorSort.key, state.setorSort.dir);
  const rowsHtml = arr.map(a=>`
    <tr>
      <td>${a.setor}</td>
      <td class="num">${fmtN(a.pessoas)}</td>
      <td class="num">${fmtR(a.custoPrevisto)}</td>
      <td class="num">${fmtR(a.custoEfetivo)}</td>
      <td class="num">${fmtR(a.custoExtra)}</td>
      <td class="num">${fmtR(a.custoTotal)}</td>
      <td class="num">${fmtR(a.custoPP)}</td>
      <td class="num">${fmtR(a.custoPorHora)}</td>
      <td class="num">${a.pctTotal.toFixed(1)}%</td>
    </tr>
  `).join('');
  const totalHtml = arr.length===0 ? '' : (()=>{
    // custoPP/custoPorHora do Total NÃO são média das linhas (cada setor tem denominador
    // diferente — pessoa-mês e horas distintos, então média simples das razões inventaria um
    // número) — é o Custo Total geral ÷ soma de pessoa-mês/horas de todos os setores, mesma
    // conta ponderada de cada linha, só que no agregado inteiro.
    // pctTotal soma direto (mesmo denominador "grand" em toda linha) e fecha em 100%.
    const t = sumCols(arr, ['pessoas','custoPrevisto','custoEfetivo','custoExtra','custoTotal','pctTotal','pessoaMeses','horas']);
    const totalCustoPP = t.pessoaMeses>0 ? t.custoTotal/t.pessoaMeses : 0;
    const totalCustoPorHora = t.horas>0 ? t.custoTotal/t.horas : 0;
    return `<tr class="row-total">
      <td>Total</td>
      <td class="num">${fmtN(t.pessoas)}</td>
      <td class="num">${fmtR(t.custoPrevisto)}</td>
      <td class="num">${fmtR(t.custoEfetivo)}</td>
      <td class="num">${fmtR(t.custoExtra)}</td>
      <td class="num">${fmtR(t.custoTotal)}</td>
      <td class="num">${fmtR(totalCustoPP)}</td>
      <td class="num">${fmtR(totalCustoPorHora)}</td>
      <td class="num">${t.pctTotal.toFixed(1)}%</td>
    </tr>`;
  })();
  document.querySelector('#finSetorTable tbody').innerHTML = (rowsHtml + totalHtml) || '<tr><td colspan="9" style="text-align:center;color:var(--muted);padding:20px;">Nenhum resultado</td></tr>';
  updateSortArrows('finSetorTable', state.setorSort);
}
wireSort('finSetorTable', 'setorSort', renderFinSetorTable);


function extraTipoAgg(){
  const rows = filteredRows();
  const totalCustoExtra = rows.reduce((s,r)=>s+r.custoHoraExtra,0);
  const arr = PCT_EXTRA_KEYS.map(k=>{
    const horas = rows.reduce((s,r)=>s+r[k],0);
    const valor = rows.reduce((s,r)=>s+r['v'+k],0);
    return {tipo:'Ex'+EXTRA_LABELS[k], horas, valor, pct: totalCustoExtra>0 ? valor/totalCustoExtra*100 : 0};
  });
  const horasNot = rows.reduce((s,r)=>s+r.notTot,0);
  const valorNot = rows.reduce((s,r)=>s+r.vexNoturno,0);
  if(horasNot>0) arr.push({tipo:'Noturno (Not.Tot.)', horas:horasNot, valor:valorNot, pct: totalCustoExtra>0 ? valorNot/totalCustoExtra*100 : 0});
  return arr;
}

function renderExtraTipoTable(){
  const arr = sortArr(extraTipoAgg(), state.extraTipoSort.key, state.extraTipoSort.dir).filter(a=>a.horas>0);
  const rowsHtml = arr.map(a=>`
    <tr>
      <td>${a.tipo}</td>
      <td class="num">${fmtH(a.horas)}</td>
      <td class="num">${a.valor>0 ? fmtR(a.valor) : '—'}</td>
      <td class="num">${a.valor>0 ? a.pct.toFixed(1)+'%' : '—'}</td>
    </tr>
  `).join('');
  const totalHtml = arr.length===0 ? '' : (()=>{
    // pct soma direto: cada linha já é valor/totalCustoExtra (mesmo denominador), então a soma
    // fecha certo em ~100% sozinha, sem recalcular nada.
    const t = sumCols(arr, ['horas','valor','pct']);
    return `<tr class="row-total">
      <td>Total</td>
      <td class="num">${fmtH(t.horas)}</td>
      <td class="num">${t.valor>0 ? fmtR(t.valor) : '—'}</td>
      <td class="num">${t.valor>0 ? t.pct.toFixed(1)+'%' : '—'}</td>
    </tr>`;
  })();
  document.querySelector('#extraTipoTable tbody').innerHTML = (rowsHtml + totalHtml) || '<tr><td colspan="4" style="text-align:center;color:var(--muted);padding:20px;">Nenhum resultado</td></tr>';
}

// ---------- Sienge: previsto x executado, saldo calculado ----------
// Previsto = SOMENTE o relatório do Sienge (aba "Relatório": Preço total reajustado, em custo;
// Quantidade, em horas) — confirmado pelo cliente: "sienge é somente planilha relatorio". Sem
// reserva/fallback pelo histograma aqui (isso é usado no resto do dashboard, não nesta tabela).
// Funções sem linha no relatório do Sienge (ver SIENGE_EXTRAS, ex: "Encarregado de
// Terraplanagem") ficam sempre "Sem previsto" nesta tabela.
// Executado = custo de folha realizado (salário + hora extra) da função, somando
// TODAS as pessoas dessa função em TODOS os meses lançados até agora — não filtra por Mês/Local/
// Setor do topo de propósito: é um valor corrente, não faz sentido zerar ao trocar de filtro.

// Monta as "linhas finais" da comparação: aplica SIENGE_GRUPOS (soma linhas do Sienge que são a
// mesma função na prática), acrescenta SIENGE_EXTRAS (funções sem linha no relatório do Sienge
// ainda) e deixa passar direto as demais funções do relatório sem alteração.
function siengeLinhas(){
  const usados = new Set();
  const linhas = [];
  SIENGE_GRUPOS.forEach(g=>{
    const membrosOrcado = SIENGE_ORCADO.filter(s=>g.membros.includes(s.funcao));
    membrosOrcado.forEach(m=>usados.add(m.funcao));
    linhas.push({
      label: g.label, membros: g.membros,
      orcadoSienge: membrosOrcado.reduce((s,m)=>s+m.orcado,0),
      orcadoHorasSienge: membrosOrcado.reduce((s,m)=>s+(m.orcadoHoras||0),0),
    });
  });
  SIENGE_ORCADO.forEach(s=>{
    if(!usados.has(s.funcao)) linhas.push({label:s.funcao, membros:[s.funcao], orcadoSienge:s.orcado, orcadoHorasSienge:s.orcadoHoras||0});
  });
  SIENGE_EXTRAS.forEach(e=>linhas.push({label:e.label, membros:e.membros, orcadoSienge:null, orcadoHorasSienge:null, semPrevisto:true}));
  // Qualquer "Função x SIENGE" real da base que ainda não caiu em nenhuma linha acima (nem
  // orçada no relatório, nem SIENGE_EXTRAS) também entra aqui, sem orçamento — a pedido do
  // cliente, pra tabela mostrar TODAS as funções de verdade, não só as que já têm linha no
  // relatório do Sienge ou entrada manual. Casam pela própria "Função x SIENGE" (não por Cargo
  // Agrupado — por isso NÃO leva semPrevisto:true, que faria siengeRowsFor() casar errado por
  // Cargo Agrupado). "Não Previsto" fica de fora de propósito, mesma regra de sempre (não é uma
  // função, é a ausência de uma — ver SEM_FUNCAO_SIENGE).
  const membrosJaCobertos = new Set(linhas.flatMap(l=>l.membros));
  const funcoesSemLinha = [...new Set(RAW.map(r=>r.funcaoSienge))]
    .filter(f=>f && !SEM_FUNCAO_SIENGE.has(f) && !membrosJaCobertos.has(f))
    .sort();
  funcoesSemLinha.forEach(f=>linhas.push({label:f, membros:[f], orcadoSienge:null, orcadoHorasSienge:null}));
  // Cargo Agrupado cujas pessoas NUNCA têm "Função x SIENGE" real (100% "Não Previsto" pra esse
  // cargo) também entra aqui, casando por Cargo Agrupado — mesmo tratamento que já existia só
  // pro "Encarregado de Terraplanagem" (SIENGE_EXTRAS), agora automático pra qualquer cargo nessa
  // situação (ex: os "Operador de X" de equipamento, vários Motoristas/Analistas etc. que
  // ficavam 100% invisíveis nesta tabela, mesmo com gente de verdade apontando hora). Usa
  // CARGO_CONTA_COMO pra juntar grafias/cargos que já são a mesma função na prática (ex: "Op de
  // Munck"/"Operador de Munck"), mesma régua já usada na aba Horas.
  const cargosComFuncaoReal = new Set(RAW.filter(r=>r.funcaoSienge && !SEM_FUNCAO_SIENGE.has(r.funcaoSienge)).map(r=>r.funcaoBase));
  const membrosJaCobertos2 = new Set(linhas.flatMap(l=>l.membros));
  const labelsJaUsados = new Set(linhas.map(l=>l.label));
  const gruposCargoSemFuncao = new Map(); // label canônico -> Set de Cargos Agrupados brutos
  [...new Set(RAW.map(r=>r.funcaoBase))].forEach(cargo=>{
    if(cargosComFuncaoReal.has(cargo) || membrosJaCobertos2.has(cargo)) return;
    const label = CARGO_CONTA_COMO[cargo] || cargo;
    if(labelsJaUsados.has(label)) return;
    if(!gruposCargoSemFuncao.has(label)) gruposCargoSemFuncao.set(label, new Set());
    gruposCargoSemFuncao.get(label).add(cargo);
  });
  [...gruposCargoSemFuncao.entries()].sort((a,b)=>a[0].localeCompare(b[0]))
    .forEach(([label, cargosSet])=>linhas.push({label, membros:[...cargosSet], orcadoSienge:null, orcadoHorasSienge:null, semPrevisto:true}));
  return linhas;
}
// Linhas de RAW que contam pra essa "linha final" do Sienge. Linhas SIENGE_EXTRAS (sem orçamento
// no relatório do Sienge, ex: Terraplanagem) casam por Cargo Agrupado, do jeito de sempre. Todas
// as outras casam pela coluna "Função x SIENGE" da planilha (fonte oficial, pessoa por pessoa)
// quando ela tem um valor real. Quando não tem (coluna vazia OU "Não Previsto" — pessoa ainda não
// confirmada função por função) cai pro fallback por Cargo Agrupado: se o nome do Cargo Agrupado
// bate com o nome desta própria linha (ou com o SIENGE_TO_FUNCAO dela), conta aqui mesmo assim —
// sem isso, uma pessoa com o cargo exato da função orçada (ex: Cargo Agrupado "Encarregado de
// Elétrica" batendo com a linha orçada "Encarregado de Elétrica" no relatório do Sienge) ficava de
// fora só porque a coluna "Função x SIENGE" dela especificamente ainda não foi preenchida — mesmo
// com a vaga clara pelo próprio cargo (achado real: uma pessoa em JUL/2026 —
// R$5.354,78 que sumiam do Realizado do Sienge por causa disso). Respeita o filtro Local
// (Obra/Matriz) do topo — Mês e Setor continuam de fora de propósito (Orçamento do Sienge é um
// valor fixo do contrato inteiro, não fatiado por mês/setor). O Orçamento em si (coluna previsto)
// não muda com esse filtro — só o Executado, que aí sim é sempre de gente real com Local definido.
function siengeRowsFor(linha){
  const porLocal = r => state.local==='all' || r.local===state.local;
  if(linha.semPrevisto) return RAW.filter(r => linha.membros.includes(r.funcaoBase) && porLocal(r));
  return RAW.filter(r=>{
    if(!porLocal(r)) return false;
    if(r.funcaoSienge && !SEM_FUNCAO_SIENGE.has(r.funcaoSienge)) return linha.membros.includes(r.funcaoSienge);
    return linha.membros.some(m => (SIENGE_TO_FUNCAO[m]||m) === r.funcaoBase);
  });
}
// Previsto final = só o relatório do Sienge, sem reserva.
function siengePrevistoCustoParaLinha(linha){
  return linha.orcadoSienge;
}
function siengePrevistoHorasParaLinha(linha){
  return (linha.orcadoHorasSienge!=null && linha.orcadoHorasSienge>0) ? linha.orcadoHorasSienge : null;
}
function siengeComparativo(){
  return siengeLinhas().map(linha=>{
    const rows = siengeRowsFor(linha);
    const valorNormal = rows.reduce((s,r)=>s+r.salario,0);
    const valorExtra = rows.reduce((s,r)=>s+r.custoHoraExtra,0);
    const totalApontado = valorNormal + valorExtra;
    const previsto = siengePrevistoCustoParaLinha(linha);
    const diferencaComExtra = previsto==null ? null : previsto-totalApontado;
    const diferencaSemExtra = previsto==null ? null : previsto-valorNormal;
    // % da diferença em relação ao Orçamento — positivo = ainda sobra orçamento; negativo = já
    // passou do orçamento naquele tanto por cento.
    const pctComExtra = (previsto!=null && previsto>0) ? diferencaComExtra/previsto*100 : null;
    const pctSemExtra = (previsto!=null && previsto>0) ? diferencaSemExtra/previsto*100 : null;
    return {funcao:linha.label, orcadoSienge:linha.orcadoSienge, previsto, valorNormal, valorExtra, totalApontado, diferencaComExtra, diferencaSemExtra, pctComExtra, pctSemExtra};
  });
}
// Mesma comparação, em horas: Previsto = Quantidade do relatório do Sienge. Horas normais/extras
// realmente lançadas pelas mesmas pessoas/funções de cima.
function siengeComparativoHoras(){
  return siengeLinhas().map(linha=>{
    const rows = siengeRowsFor(linha);
    const horasNormais = rows.reduce((s,r)=>s+r.horasMes,0);
    const horasExtras = rows.reduce((s,r)=>s+r.extraTotal,0);
    const totalApontado = horasNormais + horasExtras;
    const previsto = siengePrevistoHorasParaLinha(linha);
    const diferencaComExtra = previsto==null ? null : previsto-totalApontado;
    const diferencaSemExtra = previsto==null ? null : previsto-horasNormais;
    const pctComExtra = (previsto!=null && previsto>0) ? diferencaComExtra/previsto*100 : null;
    const pctSemExtra = (previsto!=null && previsto>0) ? diferencaSemExtra/previsto*100 : null;
    return {funcao:linha.label, previsto, horasNormais, horasExtras, totalApontado, diferencaComExtra, diferencaSemExtra, pctComExtra, pctSemExtra};
  });
}
// Renderiza a tabela Sienge (custo) recebendo o id da tabela alvo.
function renderSiengeTable(tableId){
  const tbody = document.querySelector('#'+tableId+' tbody');
  if(!tbody) return;
  const arr = siengeComparativo();
  if(arr.length===0){
    tbody.innerHTML = '<tr><td colspan="7" style="text-align:center;color:var(--muted);padding:20px;">Nenhuma função configurada para este comparativo.</td></tr>';
    return;
  }
  const corDif = d => d==null?'inherit':d>=0?'var(--green-dark)':'var(--red)';
  const cel = (val, pct, fmt) => {
    if(val==null) return '—';
    const pctTxt = pct==null?'—':pct.toFixed(1)+'%';
    return `${fmt(val)} <span class="tag">(${pctTxt})</span>`;
  };
  const rowsHtml = arr.map(a=>`
    <tr>
      <td>${a.funcao}</td>
      <td class="num">${a.previsto==null?'—':fmtR(a.previsto)}</td>
      <td class="num">${fmtR(a.valorNormal)}</td>
      <td class="num">${fmtR(a.valorExtra)}</td>
      <td class="num">${fmtR(a.totalApontado)}</td>
      <td class="num" style="color:${corDif(a.diferencaComExtra)}">${cel(a.diferencaComExtra, a.pctComExtra, fmtR)}</td>
      <td class="num" style="color:${corDif(a.diferencaSemExtra)}">${cel(a.diferencaSemExtra, a.pctSemExtra, fmtR)}</td>
    </tr>
  `).join('');
  // Linha de Total: Diferença = Previsto somado − Total Apontado somado (não soma as diferenças
  // linha a linha) — funções "Sem previsto" (SIENGE_EXTRAS) entram no Total Apontado mas não no
  // Previsto, então somar as diferenças de cada linha bateria diferente da conta feita "de olho"
  // (Previsto − Total Apontado) em cima dos dois totais já mostrados nas colunas ao lado.
  const t = sumCols(arr, ['previsto','valorNormal','valorExtra','totalApontado']);
  const difComExtraT = t.previsto>0 ? t.previsto-t.totalApontado : null;
  const difSemExtraT = t.previsto>0 ? t.previsto-t.valorNormal : null;
  const pctComExtraT = t.previsto>0 ? difComExtraT/t.previsto*100 : null;
  const pctSemExtraT = t.previsto>0 ? difSemExtraT/t.previsto*100 : null;
  const totalHtml = `<tr class="row-total">
    <td>Total</td>
    <td class="num">${t.previsto>0?fmtR(t.previsto):'—'}</td>
    <td class="num">${fmtR(t.valorNormal)}</td>
    <td class="num">${fmtR(t.valorExtra)}</td>
    <td class="num">${fmtR(t.totalApontado)}</td>
    <td class="num" style="color:${corDif(difComExtraT)}">${cel(difComExtraT, pctComExtraT, fmtR)}</td>
    <td class="num" style="color:${corDif(difSemExtraT)}">${cel(difSemExtraT, pctSemExtraT, fmtR)}</td>
  </tr>`;
  tbody.innerHTML = rowsHtml + totalHtml;
}
// Mesma tabela, versão horas.
function renderSiengeHorasTable(tableId){
  const tbody = document.querySelector('#'+tableId+' tbody');
  if(!tbody) return;
  const arr = siengeComparativoHoras();
  if(arr.length===0){
    tbody.innerHTML = '<tr><td colspan="7" style="text-align:center;color:var(--muted);padding:20px;">Nenhuma função configurada para este comparativo.</td></tr>';
    return;
  }
  const corDif = d => d==null?'inherit':d>=0?'var(--green-dark)':'var(--red)';
  const cel = (val, pct, fmt) => {
    if(val==null) return '—';
    const pctTxt = pct==null?'—':pct.toFixed(1)+'%';
    return `${fmt(val)} <span class="tag">(${pctTxt})</span>`;
  };
  const rowsHtml = arr.map(a=>`
    <tr>
      <td>${a.funcao}</td>
      <td class="num">${a.previsto==null?'—':fmtH(a.previsto)}</td>
      <td class="num">${fmtH(a.horasNormais)}</td>
      <td class="num">${fmtH(a.horasExtras)}</td>
      <td class="num">${fmtH(a.totalApontado)}</td>
      <td class="num" style="color:${corDif(a.diferencaComExtra)}">${cel(a.diferencaComExtra, a.pctComExtra, fmtH)}</td>
      <td class="num" style="color:${corDif(a.diferencaSemExtra)}">${cel(a.diferencaSemExtra, a.pctSemExtra, fmtH)}</td>
    </tr>
  `).join('');
  // Mesmo raciocínio do renderSiengeTable: Diferença = Previsto somado − Total Apontado somado,
  // não a soma das diferenças linha a linha (ver comentário lá).
  const t = sumCols(arr, ['previsto','horasNormais','horasExtras','totalApontado']);
  const difComExtraT = t.previsto>0 ? t.previsto-t.totalApontado : null;
  const difSemExtraT = t.previsto>0 ? t.previsto-t.horasNormais : null;
  const pctComExtraT = t.previsto>0 ? difComExtraT/t.previsto*100 : null;
  const pctSemExtraT = t.previsto>0 ? difSemExtraT/t.previsto*100 : null;
  const totalHtml = `<tr class="row-total">
    <td>Total</td>
    <td class="num">${t.previsto>0?fmtH(t.previsto):'—'}</td>
    <td class="num">${fmtH(t.horasNormais)}</td>
    <td class="num">${fmtH(t.horasExtras)}</td>
    <td class="num">${fmtH(t.totalApontado)}</td>
    <td class="num" style="color:${corDif(difComExtraT)}">${cel(difComExtraT, pctComExtraT, fmtH)}</td>
    <td class="num" style="color:${corDif(difSemExtraT)}">${cel(difSemExtraT, pctSemExtraT, fmtH)}</td>
  </tr>`;
  tbody.innerHTML = rowsHtml + totalHtml;
}

// ================= PÁGINA: FUNCIONÁRIOS =================
// ================= PÁGINA: AVANÇO DA OBRA =================
function sumHeadcountBy(list, mk){
  const m = new Map();
  list.filter(h=>h.mesKey===mk).forEach(h=>{ m.set(h.setor, (m.get(h.setor)||0)+h.headcount); });
  return m;
}

// ---------- Orçamento (teto do projeto inteiro): Custo Previsto = histograma (headcount ×
// salário médio da função) + valor fixo da Matriz/sede (média mensal atual). Cobre todo o
// horizonte do projeto (19 meses). O salário médio por função vem da coluna "Salario medio por
// função" que o cliente já calcula na própria aba HIST-MO (PAYLOAD.salarioMedioHist) — não é
// mais derivado da folha real aqui, é a referência que o cliente definiu direto na planilha.
// (Exceção: a tabela do Sienge não usa isso — ela só olha o relatório do Sienge, sem reserva.)
// Diferente do Custo Previsto de Início/Financeiro (custoPrevistoPorFuncaoMes), que só olha os
// meses já lançados e usa o salário vigente mês a mês da folha real (B1) — HIST_PREVISTO/avgMap
// aqui cobre inclusive meses futuros, que não têm folha real pra tirar salário nenhum.
function avgSalarioPorHistFunc(){
  return new Map(Object.entries(PAYLOAD.salarioMedioHist));
}
function custoMatrizFixoMensal(){
  if(state.local==='Obra') return 0; // filtro pediu só Obra — Matriz não entra
  if(monthsPresent.length===0) return 0;
  const total = monthsPresent.reduce((s,mk)=>
    s + RAW.filter(r=>r.mesKey===mk && r.local==='Matriz' && (state.setor.length===0||state.setor.includes(r.setor)))
      .reduce((s2,r)=>s2+r.salario,0), 0);
  return total/monthsPresent.length;
}
// monthsFilter opcional: quando informado, só soma o headcount do histograma NESSES meses (em
// vez do projeto inteiro).
function custoPrevistoObraTotal(monthsFilter){
  if(state.local==='Matriz') return 0; // filtro pediu só Matriz — Obra não entra
  const avgMap = avgSalarioPorHistFunc();
  return HIST_PREVISTO.filter(h=>(state.setor.length===0||state.setor.includes(h.setor)) && (!monthsFilter || monthsFilter.includes(h.mesKey)))
    .reduce((s,h)=>s+h.headcount*(avgMap.get(h.funcaoBase)||0), 0);
}
// Previsto Até o Fim (confirmado pelo cliente): o ORÇAMENTO do histograma pro projeto inteiro —
// Obra (todo o cronograma HIST-MO) + Matriz (fixo mensal × todos os meses). É um valor fixo da
// planilha, não recalculado com o realizado — usada aqui como o "orçado" pra comparar com a
// Projeção (ritmo atual).
function custoPrevistoOrcadoTotal(){
  const previstoMatriz = state.local!=='Obra' ? custoMatrizFixoMensal()*PROJECT_MONTHS.length : 0;
  return custoPrevistoObraTotal() + previstoMatriz;
}
// ================= ORQUESTRAÇÃO =================
// ================= PÁGINA: VISÃO BI =================
// Reaproveita os mesmos cálculos já usados em Horas/Financeiro (previstoCombinado,
// custoPrevistoTotalFiltro, monthlyAgg, custoPrevistoAgrupado) — não recalcula nada novo,
// só monta uma versão só-gráfico da mesma comparação Previsto x Realizado, sem tabela nenhuma.
//
// Filtros: Mês/Local/Setor são os mesmos do topo (state.month/state.local/state.setor, já
// aplicados por filteredRows()). O Realizado de TODO gráfico/KPI respeita os três. Os dois
// gráficos reaproveitados de outras páginas (Custo por Mês, Curva S) continuam só com
// Mês/Local/Setor — são funções compartilhadas com Financeiro.
function renderKPIsBI(){
  const rows = filteredRows();
  const totalNormais = rows.reduce((s,r)=>s+r.horasMes,0);
  const totalExtra = rows.reduce((s,r)=>s+r.extraTotal,0);
  const realizadoHoras = totalNormais+totalExtra;
  const months = state.month.length===0 ? monthsPresent : state.month;
  const previstoHoras = previstoCombinado(months);
  const saldoHoras = previstoHoras - realizadoHoras;
  const pctExecHoras = previstoHoras>0 ? realizadoHoras/previstoHoras*100 : null;

  const custoSalarioTotal = rows.reduce((s,r)=>s+r.salario,0);
  const custoExtraTotal = rows.reduce((s,r)=>s+r.custoHoraExtra,0);
  const custoRealizado = custoSalarioTotal+custoExtraTotal;
  const custoPrevisto = custoPrevistoTotalFiltro();
  const pctExecCusto = custoPrevisto>0 ? custoRealizado/custoPrevisto*100 : null;
  const pctExtras = realizadoHoras>0 ? totalExtra/realizadoHoras*100 : 0;
  const saldoHorasNormal = previstoHoras - totalNormais;
  const saldoHorasExtraParte = -totalExtra;

  const kpis = [
    {label:'Horas Previstas', value: fmtH(previstoHoras), color:'var(--accent)'},
    {label:'Horas Realizadas', value: fmtH(realizadoHoras), breakdown: [{label:'Normal', value:fmtH(totalNormais), color:'var(--green)'}, {label:'Extra', value:fmtH(totalExtra), color:'var(--red)'}], color: pctExecHoras==null?'var(--accent)':pctExecHoras<=100?'var(--green-dark)':'var(--red)'},
    {label:'Diferença de Horas', value: fmtH(Math.abs(saldoHoras)) + (saldoHoras>=0?' de folga':' estourado'), breakdown: [{label:'Normal', value:(saldoHorasNormal>=0?'+':'')+fmtH(saldoHorasNormal), color: saldoHorasNormal>=0?'var(--green)':'var(--red)'}, {label:'Extra', value:(saldoHorasExtraParte>=0?'+':'')+fmtH(saldoHorasExtraParte), color:'var(--red)'}], color: saldoHoras>=0?'var(--green-dark)':'var(--red)'},
    {label:'Custo Previsto', value: fmtR(custoPrevisto), color:'var(--accent)'},
    {label:'Custo Realizado', value: fmtR(custoRealizado), breakdown: [{label:'Salário', value:fmtR(custoSalarioTotal), color:'var(--green)'}, {label:'Extra', value:fmtR(custoExtraTotal), color:'var(--red)'}], color: pctExecCusto==null?'var(--accent)':pctExecCusto<=100?'var(--green-dark)':'var(--red)'},
    // Sem "ideal" de %Extra definido em nenhum lugar da base/planilha — cor neutra de propósito,
    // pra não inventar uma meta que o cliente nunca confirmou.
    {label:'% Extras', value: pctExtras.toFixed(1)+'%', color:'var(--accent)'},
  ];
  document.getElementById('kpisBI').innerHTML = kpis.map(k=>`
    <div class="kpi"><div class="label">${k.label}</div><div class="value" style="color:${k.color}">${k.value}</div>${k.breakdown ? `<div class="kpi-breakdown">${k.breakdown.map(b=>`<span class="kpi-bd-item"><span class="dot" style="background:${b.color}"></span>${b.label} <b>${b.value}</b></span>`).join('')}</div>` : (k.sub ? `<div class="sub">${k.sub}</div>` : '')}</div>
  `).join('');
}

// Mesma ideia de monthlyAgg() (Horas), mas respeitando Mês/Local/Setor da Visão BI. Usada
// pelos gráficos (a) Horas por Mês e (e) Horas Extras por Tipo.
function biMonthlyAgg(){
  return monthsPresent.map(mk=>{
    let rows = RAW.filter(r=>r.mesKey===mk && (state.local==='all'||r.local===state.local) && (state.setor.length===0||state.setor.includes(r.setor)));
    const agg = {mesKey:mk, mes:(rows[0]||{mes:mesLabel(mk)}).mes, normais:0, extraTotal:0};
    EXTRA_KEYS.forEach(k=>agg[k]=0);
    rows.forEach(r=>{ agg.normais+=r.horasMes; agg.extraTotal+=r.extraTotal; EXTRA_KEYS.forEach(k=>agg[k]+=r[k]); });
    const hasPrevisto = state.local==='Obra' ? overlapMonths.includes(mk) : true;
    agg.previsto = hasPrevisto ? previstoCombinado([mk]) : 0;
    return agg;
  });
}

// (a) Horas por Mês — Previsto (azul) x Realizado empilhado (Normais + Extras), mesmo padrão
// visual do gráfico "Custo por Mês" desta página.
function renderBIHorasChart(){
  const wrap = document.getElementById('biHorasWrap');
  if(!wrap || wrap.clientWidth===0) return;
  const data = biMonthlyAgg();
  const svg = document.getElementById('biHorasChart');
  const W = Math.max(wrap.clientWidth, 240), H = 142;
  const marginL = 42, marginR = 10, marginT = 10, marginB = 22;
  const plotW = W - marginL - marginR, plotH = H - marginT - marginB;
  svg.setAttribute('viewBox', `0 0 ${W} ${H}`); svg.setAttribute('width', W); svg.setAttribute('height', H);
  svg.innerHTML = '';
  const ns = 'http://www.w3.org/2000/svg';
  const mk = (tag, attrs) => { const el = document.createElementNS(ns, tag); for(const k in attrs) el.setAttribute(k, attrs[k]); return el; };
  const tooltip = document.getElementById('tooltip');
  const totals = data.map(d=>d.normais+d.extraTotal);
  const maxVal = Math.max(...totals, ...data.map(d=>d.previsto), 10) * 1.15;
  const groupW = plotW/Math.max(data.length,1);
  const barW = Math.min(22, groupW*0.28);
  const ticks = 3;
  for(let i=0;i<=ticks;i++){
    const val = maxVal/ticks*i;
    const y = marginT + plotH - (val/maxVal)*plotH;
    svg.appendChild(mk('line', {x1:marginL, x2:W-marginR, y1:y, y2:y, class:'gridline'}));
    const t = mk('text', {x:marginL-6, y:y+3, 'text-anchor':'end', class:'axis', style:'font-size:9px;'});
    t.textContent = val.toLocaleString('pt-BR', {notation:'compact', maximumFractionDigits:1});
    svg.appendChild(t);
  }
  svg.appendChild(mk('line', {x1:marginL, x2:marginL, y1:marginT, y2:marginT+plotH, class:'axis'}));
  function addTip(el, html){
    el.addEventListener('mousemove', e=>{ tooltip.style.opacity=1; tooltip.innerHTML = html; tooltip.style.left=(e.clientX+14)+'px'; tooltip.style.top=(e.clientY+10)+'px'; });
    el.addEventListener('mouseleave', ()=>tooltip.style.opacity=0);
  }
  data.forEach((d,i)=>{
    const cx = marginL + groupW*i + groupW/2;
    const total = d.normais+d.extraTotal;
    const hP = (d.previsto/maxVal)*plotH;
    const rP = mk('rect', {x:cx-barW-2, y:marginT+plotH-hP, width:barW, height:hP, rx:2, fill:'var(--previsto)', class:'bar'});
    addTip(rP, `<strong>${d.mes} &mdash; Previsto</strong><br>${fmtH(d.previsto)}`);
    svg.appendChild(rP);
    // Realizado empilhado (Normais + Extras) — mesmo padrão visual do gráfico de Custo por Mês.
    const hNorm = (d.normais/maxVal)*plotH;
    const rNorm = mk('rect', {x:cx+2, y:marginT+plotH-hNorm, width:barW, height:hNorm, fill:'var(--green)', class:'bar'});
    addTip(rNorm, `<strong>${d.mes} &mdash; Normais</strong><br>${fmtH(d.normais)}`);
    svg.appendChild(rNorm);
    const hEx = (d.extraTotal/maxVal)*plotH;
    const rEx = mk('rect', {x:cx+2, y:marginT+plotH-hNorm-hEx, width:barW, height:hEx, fill:'var(--red)', class:'bar'});
    addTip(rEx, `<strong>${d.mes} &mdash; Extras</strong><br>${fmtH(d.extraTotal)}<br><span style="opacity:.8;">Realizado total: ${fmtH(total)} (${d.previsto>0?(total/d.previsto*100).toFixed(1):'—'}% do previsto)</span>`);
    svg.appendChild(rEx);
    const t = mk('text', {x:cx, y:marginT+plotH+14, 'text-anchor':'middle', class:'axis', style:'font-size:9px;'});
    t.textContent = (d.mes||'').split('/')[0];
    svg.appendChild(t);
  });
  document.getElementById('biHorasLegend').innerHTML = `<span><span class="dot" style="background:var(--previsto)"></span>Previsto</span><span><span class="dot" style="background:var(--green)"></span>Normais</span><span><span class="dot" style="background:var(--red)"></span>Extras</span>`;
}


// (e) Horas Extras por Tipo (Ex50%–Ex110%) por mês — barras empilhadas, paleta --ex50…--ex110.
function renderBIExtrasChart(){
  const wrap = document.getElementById('biExtrasWrap');
  if(!wrap || wrap.clientWidth===0) return;
  const data = biMonthlyAgg();
  const svg = document.getElementById('biExtrasChart');
  const W = Math.max(wrap.clientWidth, 240), H = 142;
  const marginL = 38, marginR = 10, marginT = 10, marginB = 22;
  const plotW = W - marginL - marginR, plotH = H - marginT - marginB;
  svg.setAttribute('viewBox', `0 0 ${W} ${H}`); svg.setAttribute('width', W); svg.setAttribute('height', H);
  svg.innerHTML = '';
  const ns = 'http://www.w3.org/2000/svg';
  const mk = (tag, attrs) => { const el = document.createElementNS(ns, tag); for(const k in attrs) el.setAttribute(k, attrs[k]); return el; };
  const tooltip = document.getElementById('tooltip');
  const maxVal = Math.max(...data.map(d=>PCT_EXTRA_KEYS.reduce((s,k)=>s+d[k],0)), 10) * 1.15;
  const groupW = plotW/Math.max(data.length,1);
  const barW = Math.min(30, groupW*0.5);
  const ticks = 3;
  for(let i=0;i<=ticks;i++){
    const val = maxVal/ticks*i;
    const y = marginT + plotH - (val/maxVal)*plotH;
    svg.appendChild(mk('line', {x1:marginL, x2:W-marginR, y1:y, y2:y, class:'gridline'}));
    const t = mk('text', {x:marginL-6, y:y+3, 'text-anchor':'end', class:'axis', style:'font-size:9px;'});
    t.textContent = val.toLocaleString('pt-BR', {notation:'compact', maximumFractionDigits:1});
    svg.appendChild(t);
  }
  svg.appendChild(mk('line', {x1:marginL, x2:marginL, y1:marginT, y2:marginT+plotH, class:'axis'}));
  function addTip(el, html){
    el.addEventListener('mousemove', e=>{ tooltip.style.opacity=1; tooltip.innerHTML = html; tooltip.style.left=(e.clientX+14)+'px'; tooltip.style.top=(e.clientY+10)+'px'; });
    el.addEventListener('mouseleave', ()=>tooltip.style.opacity=0);
  }
  data.forEach((d,i)=>{
    const cx = marginL + groupW*i + groupW/2;
    let acc = 0;
    PCT_EXTRA_KEYS.forEach(k=>{
      if(d[k]<=0) return;
      const h = (d[k]/maxVal)*plotH;
      const rect = mk('rect', {x:cx-barW/2, y:marginT+plotH-acc-h, width:barW, height:h, fill:EXTRA_COLORS[k], class:'bar'});
      addTip(rect, `<strong>${d.mes} &mdash; Ex${EXTRA_LABELS[k]}</strong><br>${fmtH(d[k])}`);
      svg.appendChild(rect);
      acc += h;
    });
    // Total do mês acima da pilha inteira (não por segmento — 6 tipos empilhados ficaria poluído).
    if(acc>0){
      const totalMes = PCT_EXTRA_KEYS.reduce((s,k)=>s+d[k],0);
      const tv = mk('text', {x:cx, y:marginT+plotH-acc-3, 'text-anchor':'middle', style:'font-size:7.5px;font-weight:700;fill:var(--text);'});
      tv.textContent = (totalMes||0).toLocaleString('pt-BR',{notation:'compact',maximumFractionDigits:1})+'h';
      svg.appendChild(tv);
    }
    const t = mk('text', {x:cx, y:marginT+plotH+14, 'text-anchor':'middle', class:'axis', style:'font-size:9px;'});
    t.textContent = (d.mes||'').split('/')[0];
    svg.appendChild(t);
  });
  document.getElementById('biExtrasLegend').innerHTML = PCT_EXTRA_KEYS.map(k=>`<span><span class="dot" style="background:${EXTRA_COLORS[k]}"></span>Ex${EXTRA_LABELS[k]}</span>`).join('');
}

// (f) Top Funções — Horas Extras, ranking horizontal (âmbar). Agrupado por Cargo Agrupado
// (função), somando extraTotal de todo mundo daquela função — não por pessoa.
function renderBIRanking(){
  const el = document.getElementById('biRankingWrap');
  if(!el) return;
  const map = new Map();
  filteredRows().forEach(r=>{ map.set(r.funcaoBase, (map.get(r.funcaoBase)||0)+r.extraTotal); });
  let arr = [...map.entries()].map(([funcao,extra])=>({funcao,extra})).filter(a=>a.extra>0);
  arr = arr.sort((a,b)=>b.extra-a.extra).slice(0,18);
  if(arr.length===0){ el.innerHTML = '<p class="footnote">Sem horas extras no filtro atual.</p>'; return; }
  const maxVal = Math.max(...arr.map(a=>a.extra), 1);
  el.innerHTML = arr.map(a=>`
    <div class="hbar-row">
      <div class="hbar-label" title="${a.funcao}">${a.funcao}</div>
      <div class="hbar-track"><div class="hbar-fill" style="width:${(a.extra/maxVal*100).toFixed(1)}%;background:var(--amber);"></div></div>
      <div class="hbar-val">${fmtH(a.extra)}</div>
    </div>
  `).join('');
}

// Faixa "estilo mercado financeiro" no topo da Visão BI: desvio de cada setor (Realizado vs
// Previsto) rolando em looping infinito — mesma fonte de dados do gráfico "Custo por Setor"
// (biSetorData), só que renderizado como ticker em vez de barra. Lista duplicada 2x + a
// animação CSS anda -50% da largura, pra criar o loop sem "salto" visível no fim.
function renderBIPage(){
  renderKPIsBI();
  renderBIRanking();
  renderBIHorasChart();
  renderBIAvancoFisicoChart({wrap:'biCurvaWrap', svg:'biCurvaChart', legend:'biCurvaLegend', height:142, minWidth:240, marginL:34});
  renderFinChart({wrap:'biCustoWrap', svg:'biCustoChart', legend:'biCustoLegend', height:142, minWidth:240, marginL:46});
  renderBIExtrasChart();
}

// ---------- Visão BI: filtros da barra lateral (Mês/Setor espelham o topo — mesma
// state.month/state.setor, outra instância do mesmo componente) ----------
const monthMsfBI = makeMultiSelectFilter({
  btnId:'biFiltroMesBtn', panelId:'biFiltroMesPanel',
  getOptions: () => monthsPresent.map(m=>({value:m, label:RAW.find(r=>r.mesKey===m).mes})),
  getSelected: () => state.month, setSelected: v=>{ state.month = v; },
  allLabel:'Todos os meses', itemNoun:'meses selecionados'
});
const setorMsfBI = makeMultiSelectFilter({
  btnId:'biFiltroSetorBtn', panelId:'biFiltroSetorPanel',
  getOptions: () => setoresPresentes.map(s=>({value:s, label:s})),
  getSelected: () => state.setor, setSelected: v=>{ state.setor = v; },
  allLabel:'Todos os setores', itemNoun:'setores selecionados'
});
refreshAllMsf(); // primeira pintura de todas as instâncias (topo + BI)

function renderAll(){
  renderHome();
  renderHomeHistograma();
  renderBIAvancoFisicoChart({wrap:'homeCurvaSWrap', svg:'homeCurvaSChart', legend:'homeCurvaSLegend'});
  renderKPIsHoras();
  renderChart();
  renderFuncaoTable();
  renderSetorHorasTable();
  renderKPIsFin();
  renderFinChart();
  renderProjChart();
  renderCurvaSChart();
  renderProjTable();
  renderFinSetorTable();
  renderExtraTipoTable();
  renderSiengeTable('siengeTable');
  renderSiengeHorasTable('siengeHorasTableHoras');
  renderBIPage();
}
renderAll();
window.addEventListener('resize', ()=>{ renderChart(); renderHomeHistograma(); renderBIAvancoFisicoChart({wrap:'homeCurvaSWrap', svg:'homeCurvaSChart', legend:'homeCurvaSLegend'}); renderFinChart(); renderProjChart(); renderCurvaSChart(); renderBIPage(); });

// ================= SHELL: título da página, drawer (mobile), tema =================
// Título/subtítulo do topbar mudam por página (a sidebar não tem mais o cabeçalho
// fixo "Dashboard de Projetos" que existia no menu horizontal) — sem precisar duplicar
// texto dentro de cada page div, que já não tem mais o <p class="section-title"> próprio.
const PAGE_META = {
  inicio:     {t:'Início',      s:'Panorama de efetivo, custo e avanço de obra.'},
  bi:         {t:'Visão BI',    s:'Previsto x realizado em gráficos comparativos.'},
  horas:      {t:'Horas',       s:'Horas previstas, normais e extras por setor e função.'},
  financeiro: {t:'Financeiro',  s:'Custo de folha: salário-base e hora extra.'},
  glossario:  {t:'Explicações', s:'Como cada número é calculado e de onde ele vem.'}
};
function syncShell(page){
  const m = PAGE_META[page]; if(!m) return;
  document.getElementById('pageTitle').textContent = m.t;
  document.getElementById('pageSub').textContent = m.s;
  document.title = m.t + ' — Dashboard ' + OBRA_NOME;
  document.body.classList.remove('nav-open');
  window.scrollTo({top:0, behavior:'smooth'});
}
const _goToPage = goToPage;
goToPage = function(page){ _goToPage(page); syncShell(page); };

// Drawer da sidebar no mobile (<1024px, ver media query do CSS)
const menuBtn = document.getElementById('menuBtn');
const sidebarScrim = document.getElementById('sidebarScrim');
if(menuBtn && sidebarScrim){
  menuBtn.addEventListener('click', ()=>document.body.classList.toggle('nav-open'));
  sidebarScrim.addEventListener('click', ()=>document.body.classList.remove('nav-open'));
  document.addEventListener('keydown', e=>{ if(e.key==='Escape') document.body.classList.remove('nav-open'); });
}

// Alternância de tema claro/escuro — antes só existia via prefers-color-scheme do SO,
// sem controle manual nenhum. Persiste a escolha no localStorage do navegador.
const themeToggle = document.getElementById('themeToggle');
function currentTheme(){
  const set = document.documentElement.getAttribute('data-theme');
  if(set) return set;
  return window.matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light';
}
try{ const saved = localStorage.getItem('medeiros-theme'); if(saved) document.documentElement.setAttribute('data-theme', saved); }catch(e){}
if(themeToggle){
  themeToggle.addEventListener('click', ()=>{
    const next = currentTheme()==='dark' ? 'light' : 'dark';
    document.documentElement.setAttribute('data-theme', next);
    try{ localStorage.setItem('medeiros-theme', next); }catch(e){}
  });
}
"""

html = HTML_TEMPLATE.replace('__DATA_JSON__', data_json).replace('__JS_BODY__', JS_BODY).replace('__OBRA_NOME__', OBRA_NOME)
with open(OUT, 'w', encoding='utf-8') as f:
    f.write(html)
print('Written to', OUT)
