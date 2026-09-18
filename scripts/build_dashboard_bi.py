# -*- coding: utf-8 -*-
"""
Gera um HTML enxuto contendo SÓ a página "Visão BI" do dashboard (sem Início, Horas,
Financeiro, Funcionários, Suporte à Decisão, Explicações).

Como funciona (pra não duplicar as ~430 linhas de leitura do Excel nem os cálculos de
Previsto/Custo, que moram só em build_dashboard.py):
1. Importa build_dashboard.py como módulo. Isso roda o pipeline inteiro (Excel -> payload)
   e, como efeito colateral, RE-ESCREVE "Dashboard Horas - Medeiros.html" de novo (mesmo
   conteúdo de sempre) -- garantia extra de que a Visão BI nunca fica com dado desatualizado
   em relação ao dashboard completo, ao custo de regravar os dois arquivos sempre que
   qualquer um dos dois scripts roda.
2. Reaproveita bd.HTML_TEMPLATE (cabeçalho <head>/CSS e o bloco de HTML da própria página
   "Visão BI") e bd.data_json BYTE A BYTE, direto da string-fonte -- sem recopiar/redigitar,
   pra não divergir do dashboard completo se o layout mudar.
3. O JS da Visão BI reaproveita várias funções do dashboard inteiro (renderFinChart,
   renderCurvaSChart, custoPrevistoAgrupado, filteredRows, etc. -- ver comentário
   "PÁGINA: VISÃO BI" em build_dashboard.py). Em vez de copiar módulo JS na mão (arriscado:
   qualquer ajuste futuro nessas funções compartilhadas ficaria fora de sincronia), este
   script fatia o JS_BODY original em blocos de topo (cada function/const/let) e calcula o
   fecho transitivo de dependências a partir da própria Visão BI -- ou seja, sempre que
   build_dashboard.py mudar, rodar este script de novo já recalcula o que precisa ser
   incluído, sem curadoria manual de nomes.

Saída: "Dashboard BI - Medeiros.html" (arquivo separado; não mexe no dashboard completo além
de regravá-lo idêntico, como já roda hoje).
"""
import sys, os, re

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, SCRIPT_DIR)
import build_dashboard as bd

OUT_BI = r"C:\Users\User\OneDrive - ENGETECNICA\Documentos\Homem Hora\Dashboard BI - Medeiros.html"

# ============================================================
# 1. Fatiar o JS_BODY original em blocos de topo (function NOME(...) / const NOME = / let NOME =)
# ============================================================
js_lines = bd.JS_BODY.split('\n')
ANCHOR_RE = re.compile(r'^(function\s+([A-Za-z_$][\w$]*)\s*\(|(?:const|let)\s+([A-Za-z_$][\w$]*)\s*=)')
anchors = []
for i, line in enumerate(js_lines):
    m = ANCHOR_RE.match(line)
    if m:
        anchors.append((i, m.group(2) or m.group(3)))

blocks = []
for idx, (line_i, name) in enumerate(anchors):
    end_i = anchors[idx + 1][0] if idx + 1 < len(anchors) else len(js_lines)
    blocks.append({'name': name, 'text': '\n'.join(js_lines[line_i:end_i])})

name_to_block = {b['name']: b for b in blocks}
all_names = set(name_to_block)
dupes = [n for n in all_names if sum(1 for b in blocks if b['name'] == n) > 1]
if dupes:
    raise RuntimeError(f'build_dashboard.py tem declarações de topo com nome repetido: {dupes} '
                        f'-- a fatia por âncora ficou ambígua, revise build_dashboard_bi.py.')

# ============================================================
# 2. Fecho transitivo de dependências, partindo do que a Visão BI usa de verdade
# ============================================================
IDENT_RE = re.compile(r'[A-Za-z_$][\w$]*')
JS_KEYWORDS = {
    'const', 'let', 'var', 'function', 'return', 'if', 'else', 'for', 'while', 'do', 'switch',
    'case', 'break', 'continue', 'new', 'this', 'typeof', 'instanceof', 'in', 'of', 'null',
    'undefined', 'true', 'false', 'void', 'delete', 'try', 'catch', 'finally', 'throw', 'class',
    'extends', 'super', 'yield', 'async', 'await', 'static', 'get', 'set', 'import', 'export',
    'default', 'from', 'as',
}

def deps_of(text, self_name):
    found = set()
    for m in IDENT_RE.finditer(text):
        tok = m.group(0)
        if tok == self_name or tok in JS_KEYWORDS:
            continue
        if tok in all_names:
            found.add(tok)
    return found

# Tudo que o código da própria "PÁGINA: VISÃO BI" em build_dashboard.py chama/lê diretamente,
# mais o setup base (payload/state/constantes/formatação) que todo o resto depende.
ROOTS = [
    'renderKPIsBI', 'biMonthlyAgg',
    'renderBIHorasChart', 'renderBIAvancoFisicoChart', 'renderBIExtrasChart',
    'renderBIRanking', 'renderBIPage',
    'monthMsfBI', 'setorMsfBI',
    'PAYLOAD', 'RAW', 'HIST_PREVISTO', 'HIST_CONTRATADO', 'HIST_LIBERADO', 'PROJECT_MONTHS',
    'HOURS_PP', 'TODAY_KEY', 'SETOR_ORDER', 'OBRA_NOME', 'CARGO_TO_HISTFUNC',
    'SIENGE_ORCADO', 'AVANCO_FISICO', 'AVANCO_FISICO_PLAN', 'AVANCO_FISICO_REAL',
    'MONTH_ORDER', 'EXTRA_KEYS', 'PCT_EXTRA_KEYS', 'EXTRA_LABELS', 'EXTRA_COLORS',
    'mesLabel', 'fmtH', 'fmtN', 'fmtR', 'fmtPct', 'kpiSignal', 'state', 'monthsPresent',
    'msfInstances', 'openMsfPanels', 'closeAllMsfPanels', 'refreshAllMsf', 'makeMultiSelectFilter',
    'setoresPresentes', 'localSel',
    # 'cargoEntries' não é lido por nome em nenhum outro lugar, mas seu forEach de cauda POPULA
    # HISTFUNC_TO_CARGOS (declarado como {} vazio no bloco anterior) -- custoPrevistoAgrupado lê
    # HISTFUNC_TO_CARGOS, então essa população precisa entrar mesmo sem ninguém referenciar o
    # identificador "cargoEntries" (achado pela auditoria de "declaração oca" abaixo).
    'cargoEntries',
    # Só pro texto do subtítulo do cabeçalho (Período + nº de pessoas) -- ver DEP_OVERRIDES/
    # CUSTOM_TEXT abaixo, que cortam o resto do bloco original (esses dois nomes vêm seguidos,
    # na sequência, do código que liga os botões de navegação entre páginas).
    'firstMes', 'lastMes',
]
missing_roots = [r for r in ROOTS if r not in all_names]
if missing_roots:
    raise RuntimeError(f'build_dashboard.py mudou e estes nomes que a Visão BI precisa não '
                        f'existem mais no JS: {missing_roots}. Reveja a lista ROOTS em '
                        f'build_dashboard_bi.py (grep pelo nome antigo em build_dashboard.py '
                        f'pra achar o que ele virou).')

# 'renderAll' original chama as 20 páginas do dashboard inteiro (é o corpo que o botão
# "Selecionar todos"/"Limpar" de cada filtro e o <select> de Local disparam por nome) -- só
# queremos o NOME (pra essas ligações continuarem funcionando), não o corpo original. O mesmo
# bloco de 'renderAll' carrega, colado depois da função, a chamada de bootstrap e o listener de
# resize -- por isso os dois entram junto na substituição de texto (CUSTOM_TEXT) abaixo.
# 'lastMes': o texto real, depois de montar o subtítulo, liga os botões de navegação entre
# páginas (que não existem mais aqui) -- cortamos essa parte também.
DEP_OVERRIDES = {
    'renderAll': {'renderBIPage'},
    'lastMes': {'OBRA_NOME', 'RAW', 'firstMes'},
}

included = set()
queue = list(dict.fromkeys(ROOTS))
while queue:
    n = queue.pop()
    if n in included:
        continue
    included.add(n)
    dep_set = DEP_OVERRIDES.get(n)
    if dep_set is None:
        dep_set = deps_of(name_to_block[n]['text'], n)
    for d in dep_set:
        if d not in included:
            queue.append(d)

# ---- Auditoria de segurança: barra a geração se o fecho ficou incoerente ----
# (a) nenhum bloco incluído deveria citar um nome excluído -- exceto os que a gente
#     DELIBERADAMENTE reescreve em CUSTOM_TEXT (esses citam o excluído só no texto ORIGINAL,
#     que a gente descarta na montagem final).
CUSTOM_TEXT_NAMES = {'renderAll', 'lastMes'}
excluded = all_names - included
bad_refs = {}
for b in blocks:
    if b['name'] not in included or b['name'] in CUSTOM_TEXT_NAMES:
        continue
    hits = deps_of(b['text'], b['name']) & excluded
    if hits:
        bad_refs[b['name']] = hits
if bad_refs:
    raise RuntimeError(f'Fecho de dependências inconsistente -- estes blocos incluídos citam '
                        f'nomes que ficaram de fora: {bad_refs}. Ou adiciona o(s) nome(s) '
                        f'faltante(s) em ROOTS, ou (se for referência morta, tipo handler de '
                        f'botão que não existe mais nesta página) trata como CUSTOM_TEXT.')

# (b) declarações "ocas" (const NOME = {} / [] / new Map() / new Set()) cujo POVOAMENTO de
#     verdade mora em outro bloco -- se esse outro bloco não estiver incluído, o objeto fica
#     sempre vazio e algo vai silenciosamente dar errado (dado sumido, não erro).
HOLLOW_RE = re.compile(r'=\s*(\{\}|\[\]|new Map\(\)|new Set\(\))\s*;?\s*$')
for name in included:
    text = name_to_block[name]['text']
    nonempty = [l for l in text.split('\n') if l.strip() and not l.strip().startswith('//')]
    if not nonempty or len(nonempty) > 2 or not HOLLOW_RE.search(nonempty[0]):
        continue
    mutators = [b['name'] for b in blocks
                if b['name'] != name and re.search(re.escape(name) + r'(\[|\.push\(|\.set\(|\.add\()', b['text'])]
    missing_mutators = [m for m in mutators if m not in included]
    # Só é problema se TODOS os povoadores reais ficaram de fora (se pelo menos um está
    # incluído, o objeto é povoado igual -- os outros podem ser consumidores só de leitura,
    # tipo "obj[chave]" pra checar existência, que o regex acima também casa).
    if missing_mutators and not (set(mutators) - set(missing_mutators)):
        raise RuntimeError(f'"{name}" é declarado vazio e só é populado em {mutators}, mas '
                            f'nenhum desses está incluído -- adicione um deles em ROOTS.')

# ============================================================
# 3. Montar o JS enxuto, preservando a ordem original (import por const depende da ordem)
# ============================================================
CUSTOM_TEXT = {
    'lastMes': (
        "const lastMes = RAW.find(r=>r.mesKey===monthsPresent[monthsPresent.length-1]).mes;\n"
        "document.getElementById('subtitleRange').textContent = "
        "`${OBRA_NOME}  \u2022  ${firstMes} a ${lastMes}  \u2022  ${new Set(RAW.map(r=>r.nome)).size} pessoas`;"
    ),
    'renderAll': (
        "function renderAll(){ renderBIPage(); }\n"
        "renderAll();\n"
        "window.addEventListener('resize', ()=>{ renderBIPage(); });"
    ),
}

js_parts = []
for b in blocks:  # ordem original do arquivo-fonte
    if b['name'] not in included:
        continue
    js_parts.append(CUSTOM_TEXT.get(b['name'], b['text']))
js_bi = '\n\n'.join(js_parts)

print(f'Visão BI: {len(included)}/{len(all_names)} funções/consts do JS original reaproveitadas '
      f'(fecho de dependências a partir da própria página).')

# ============================================================
# 4. Montar o HTML: cabeçalho <head>/CSS e o bloco da página "Visão BI" vêm literalmente do
#    template original (bd.HTML_TEMPLATE) -- zero retransciração manual de CSS/marcação.
# ============================================================
tpl = bd.HTML_TEMPLATE
head_html = tpl[tpl.index('<head>'):tpl.index('</head>') + len('</head>')]
head_html = head_html.replace('<title>Dashboard de Projetos - Medeiros</title>',
                               '<title>Visão BI - Medeiros</title>')

BI_START = '<!-- ================= PÁGINA: VISÃO BI ================= -->'
BI_END = '<!-- ================= PÁGINA: HORAS ================= -->'
bi_page_html = tpl[tpl.index(BI_START):tpl.index(BI_END)].rstrip()
bi_page_html = bi_page_html.replace('<div class="page" id="page-bi">', '<div class="page active" id="page-bi">', 1)

html = f"""<!doctype html>
<html lang="pt-BR">
{head_html}
<body class="bi-active">
<div class="main" style="min-height:100vh;">
  <header class="topbar">
    <div class="topbar-title" style="display:flex;align-items:center;gap:10px;">
      <svg class="brand-mark" viewBox="0 0 36 36" xmlns="http://www.w3.org/2000/svg">
        <circle cx="14" cy="18" r="10.5" fill="none" stroke="var(--text)" stroke-width="2.4"/>
        <circle cx="23" cy="18" r="10.5" fill="none" stroke="#E8622C" stroke-width="2.4"/>
      </svg>
      <div>
        <h1>Vis&atilde;o BI &mdash; __OBRA_NOME__</h1>
        <p class="subtitle" id="subtitleRange">Per&iacute;odo: --</p>
      </div>
    </div>
    <div class="controls">
      <select id="localFilter" title="Local">
        <option value="all">Obra + Matriz</option>
        <option value="Obra">S&oacute; Obra</option>
        <option value="Matriz">S&oacute; Matriz</option>
      </select>
    </div>
  </header>

  <div class="wrap">
{bi_page_html}
  </div>
</div>
<div class="tooltip" id="tooltip"></div>

<script id="raw-data" type="application/json">__DATA_JSON__</script>
<script>
{js_bi}
</script>
</body>
</html>
"""
html = html.replace('__DATA_JSON__', bd.data_json).replace('__OBRA_NOME__', bd.OBRA_NOME)

with open(OUT_BI, 'w', encoding='utf-8') as f:
    f.write(html)
print(f'Gerado: {OUT_BI}')
