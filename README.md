# Dashboard Homem-Hora — SE Medeiros Neto II

Gera, a partir da planilha de folha (Excel), um dashboard HTML de **horas** e **custo de mão de obra**
(Previsto x Realizado), com comparativo contra o orçamento do Sienge e o avanço físico da obra.

> **Dados pessoais fora do repositório.** As planilhas (`.xlsx`) e os HTMLs gerados contêm nome e salário
> de funcionários e ficam de fora do Git (ver `.gitignore`). Aqui só vão código e documentação.
> O arquivo [`STATUS.md`](STATUS.md) traz apenas **totais agregados** e é atualizado sozinho a cada execução.

## Como atualizar

1. Feche o Excel (o arquivo não pode estar aberto).
2. Dê dois cliques em **`Atualizar Dashboard.bat`**.

O `.bat` faz, em ordem:

1. `scripts/build_dashboard.py` → gera `Dashboard Horas - Medeiros.html`
2. `scripts/build_dashboard_bi.py` → gera `Dashboard BI - Medeiros.html` (só a página Visão BI)
3. `scripts/update_status.py` → gera o `STATUS.md` (resumo agregado)
4. `git commit` + `git push` do `STATUS.md` (e do código, se tiver mudado) para o GitHub

Se o passo do Git falhar (sem internet, sem login), o dashboard já foi gerado normalmente — só o envio ao GitHub é pulado.

## Estrutura

| Arquivo | Função |
|---|---|
| `Atualizar Dashboard.bat` | Atalho que roda tudo acima |
| `scripts/build_dashboard.py` | Lê a planilha, calcula e gera o dashboard completo (HTML + JS autocontidos) |
| `scripts/build_dashboard_bi.py` | Gera a versão enxuta só com a Visão BI, reaproveitando o JS do dashboard completo |
| `scripts/funcao_map.py` | Mapas de cargo → setor / função do histograma / função do Sienge |
| `scripts/update_status.py` | Gera o `STATUS.md` a partir do HTML gerado |
| `STATUS.md` | Resumo agregado (período, pessoas, horas, custo por mês) — automático |

## Fontes de dados esperadas (fora do Git)

- `HHT - Medeiros - Base de dados.xlsx` — aba **Base de dados** (Power Query sobre as abas mensais), **HIST-MO** (histograma de mão de obra) e **Relatório** (orçamento do Sienge)
- `Base de dados/HHT - 5005 MEDEIROS .xlsx` — abas mensais brutas de horas extras
- Planilha de curvas S (avanço físico), caminho definido em `build_dashboard.py`

## Regras de cálculo importantes

- **Jornada padrão**: 220 h/mês; quando "Normais" < 130 h no mês, considera 110 h.
- **Hora extra**: `Salário/hora × multiplicador × horas` (Ex50% ×1,5 … Ex110% ×2,1); adicional noturno = `salário/hora × 1,2 × horas`.
- Linhas com cargo **"Demissão"** (marcador de desligamento) ficam fora do dashboard.
- **Diferença** = Previsto − Realizado, decomposta em parte Normal/Salário e parte Hora Extra (que não tem previsto próprio).
- **% Executado** = Realizado ÷ Previsto, com o desvio (pctExec − 100) decomposto em pontos percentuais de Normal/Salário e de Hora Extra.

## Requisitos

- Python 3.10+ com `openpyxl`
- Git e GitHub CLI (`gh`) autenticados, só para o envio automático do `STATUS.md`
