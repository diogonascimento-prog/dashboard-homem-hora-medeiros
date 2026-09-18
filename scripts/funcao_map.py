# -*- coding: utf-8 -*-
# Setor (departamento) por função. A partir de agora a "função consolidada" em si
# vem pronta da coluna "Cargo Agrupado" da planilha (feita pelo usuário) — aqui só
# mapeamos essa função (ou, para o histograma, a função crua do HIST-MO) para um setor.

SETOR_ORDER = ['Gestão/Administrativo', 'SSMA', 'Qualidade', 'Topografia', 'Civil',
               'Terraplanagem', 'Equipamentos', 'Montagem', 'Elétrica', 'Transporte']

# ---- Cargo Agrupado (Base de dados) -> Setor ----
CARGO_TO_SETOR = {
    'Administrativo de Obras': 'Gestão/Administrativo',
    'Ajudante de Obra': 'Civil',
    'Almoxarife': 'Gestão/Administrativo',
    'Analista SR - Planejamento': 'Gestão/Administrativo',
    'Analista de Planejamento': 'Gestão/Administrativo',
    'Analista de Qualidade': 'Qualidade',
    'Armador': 'Civil',
    'Auxiliar Administrativo': 'Gestão/Administrativo',
    'Auxiliar de Almoxarifado': 'Gestão/Administrativo',
    'Auxiliar de Eletricista': 'Elétrica',
    'Auxiliar de Escritorio': 'Gestão/Administrativo',
    'Auxiliar de Serviços Gerais': 'Gestão/Administrativo',
    'Carpinteiro': 'Civil',
    'Demissão': 'Desligamentos',
    'Eletricista FC': 'Elétrica',
    'Encarregado Civil': 'Civil',
    'Encarregado de Montagem': 'Montagem',
    'Encarregado de terraplanagem': 'Terraplanagem/Equipamentos',
    'Enfermeiro Socorrista': 'SSMA',
    'Engenheiro Seguranca Trabalho': 'SSMA',
    'Estagiario (a)': 'Gestão/Administrativo',
    'Gerente de Contrato': 'Gestão/Administrativo',
    'Gerente de Obras': 'Gestão/Administrativo',
    'Montador': 'Montagem',
    'Motorista Socorrista': 'SSMA',
    'Motorista Truck': 'Transporte',
    'Motorista de Caminhao': 'Transporte',
    'Motorista de Caminhão Basculante': 'Transporte',
    'Motorista de caminhão comboio': 'Transporte',
    'Motorista de Ônibus': 'Transporte',
    'Op de Munck': 'Montagem',
    'Operador Betoneira': 'Terraplanagem/Equipamentos',
    'Operador de Guindaste': 'Montagem',
    'Operador de Maquinas': 'Terraplanagem/Equipamentos',
    'Operador de Motoniveladora': 'Terraplanagem/Equipamentos',
    'Operador de Motoserra': 'Terraplanagem/Equipamentos',
    'Operador de Munck': 'Montagem',
    'Operador de Retroescavadeira': 'Terraplanagem/Equipamentos',
    'Operador de Trator de Esteira': 'Terraplanagem/Equipamentos',
    'Pedreiro': 'Civil',
    'Projetista PL - Projeto Civil': 'Gestão/Administrativo',
    'Supervisor Segurança do Trabalho': 'SSMA',
    'Supervisor de Obras': 'Gestão/Administrativo',
    'Supervisor de Planejamento': 'Gestão/Administrativo',
    'Tecnico Seguranca do Trabalho': 'SSMA',
    'Tecnico de Meio Ambiente': 'SSMA',
    'Tecnico de Qualidade': 'Qualidade',
    'Topografo': 'Topografia',
    'Técnico de Qualidade': 'Qualidade',
}

# ---- Obra x Matriz: regra real, copiada da própria Power Query "Base de dados" da planilha
# (etapa "Local Adicionado" em Formulas/Section1.m). A query classifica pelo texto da FUNÇÃO:
# contém "estagi", "planejamento", "gerente de contrato" ou "projetista" -> Matriz; senão -> Obra.
# Fica assim 100% igual ao que a própria planilha do usuário já calcula.
CARGO_IS_OBRA = {
    'Administrativo de Obras': True,
    'Ajudante de Obra': True,
    'Almoxarife': True,
    'Analista SR - Planejamento': False,   # contém "planejamento"
    'Analista de Planejamento': False,     # contém "planejamento"
    'Analista de Qualidade': True,
    'Armador': True,
    'Auxiliar Administrativo': True,
    'Auxiliar de Almoxarifado': True,
    'Auxiliar de Eletricista': True,
    'Auxiliar de Escritorio': True,
    'Auxiliar de Serviços Gerais': True,
    'Carpinteiro': True,
    'Eletricista FC': True,                # = "Eletricista Força Controle"
    'Encarregado Civil': True,
    'Encarregado de Montagem': True,       # = "Encarregado Montagem"
    'Encarregado de terraplanagem': True,  # = "Encarregado Terraplanagem"
    'Enfermeiro Socorrista': True,          # confirmado pelo usuário: é cargo de obra
    'Engenheiro Seguranca Trabalho': True, # = "Engenheiro de Segurança"
    'Estagiario (a)': False,
    'Gerente de Contrato': False,          # aparece no histograma (alocação de custo), mas confirmado pelo usuário: é Matriz
    'Gerente de Obras': True,              # confirmado pelo usuário: é cargo de obra
    'Montador': True,
    'Motorista Socorrista': True,          # = "Motorista de Ambulância"
    'Motorista Truck': True,               # = "Operador de Caminhão Truck"
    'Motorista de Caminhao': True,
    'Motorista de Caminhão Basculante': True,  # = "Motorista de caçamba"
    'Motorista de caminhão comboio': True, # = "Operador de Caminhão Comboio"
    'Motorista de Ônibus': True,           # confirmado pelo usuário: é cargo de obra
    'Op de Munck': True,
    'Operador Betoneira': True,            # = "Operador de Betoneira"
    'Operador de Guindaste': True,
    'Operador de Maquinas': True,          # confirmado pelo usuário: é cargo de obra
    'Operador de Motoniveladora': True,
    'Operador de Motoserra': True,         # confirmado pelo usuário: é cargo de obra
    'Operador de Munck': True,
    'Operador de Retroescavadeira': True,
    'Operador de Trator de Esteira': True, # = "Operador de Trator de Esteiras"
    'Pedreiro': True,
    'Projetista PL - Projeto Civil': False,
    'Supervisor Segurança do Trabalho': True,  # confirmado pelo usuário: é cargo de obra
    'Supervisor de Obras': True,           # confirmado pelo usuário: é cargo de obra
    'Supervisor de Planejamento': False,
    'Tecnico Seguranca do Trabalho': True, # = "Tec. De Segurança"
    'Tecnico de Meio Ambiente': True,      # = "Tec. Meio Ambiente"
    'Tecnico de Qualidade': True,          # = "Tec. De Qualidade"
    'Topografo': True,
    'Técnico de Qualidade': True,
}

# ---- Cargo Agrupado -> nome exato da função no histograma HIST-MO (quando existe) ----
# None = esse cargo não tem linha própria no histograma (planejamento não detalhou por essa
# função). Usado para calcular "Horas Previstas" por função/pessoa nas tabelas do dashboard.
CARGO_TO_HISTFUNC = {
    'Administrativo de Obras': 'Administrativo de Obra',
    'Ajudante de Obra': 'Ajudante civil',
    'Almoxarife': 'Almoxarife',
    'Analista SR - Planejamento': None,
    'Analista de Planejamento': None,
    'Analista de Qualidade': None,
    'Armador': 'Armador',
    'Auxiliar Administrativo': None,
    'Auxiliar de Almoxarifado': 'Almoxarife',
    'Auxiliar de Eletricista': 'Ajudante elétrica',
    'Auxiliar de Escritorio': None,
    # Confirmado pelo cliente: quem está lançado como "Auxiliar de Serviços Gerais" na prática é
    # a vaga de Zeladora do histograma (mesmo padrão do Clenilson/Engenheiro Residente).
    'Auxiliar de Serviços Gerais': 'Zeladora',
    'Carpinteiro': 'Carpinteiro',
    'Demissão': None,
    'Eletricista FC': 'Eletricista Força Controle',
    'Encarregado Civil': 'Encarregado Civil',
    'Encarregado de Montagem': 'Encarregado Montagem',
    'Encarregado de terraplanagem': 'Encarregado Terraplanagem',
    # Confirmado pelo cliente: "Enfermeiro Socorrista" é a mesma vaga que "Tec. De Enfermagem" no
    # histograma.
    'Enfermeiro Socorrista': 'Tec. De Enfermagem',
    'Engenheiro Seguranca Trabalho': 'Engenheiro de Segurança',
    'Estagiario (a)': None,
    'Gerente de Contrato': 'Gerente de Contrato',
    # Confirmado pelo cliente: quem está lançado como "Gerente de Obras" (Clenilson) na prática é
    # o engenheiro residente da obra — liga com a vaga do histograma em vez de deixar as duas
    # linhas separadas (uma sem previsto, outra sem apontamento).
    'Gerente de Obras': 'Engenheiro Residente',
    'Montador': 'Montador',
    'Motorista Socorrista': 'Motorista de Ambulância',
    'Motorista Truck': 'Operador de Caminhão Truck',
    # Confirmado pelo cliente (2026-08-26): todo "Motorista de Caminhao" é motorista de caçamba,
    # exceto quem já tem cargo próprio (ex: Augusto Cesar Martins é "Motorista de caminhão
    # comboio", mapeado separado logo abaixo — não entra aqui). "Motorista de Caminhão
    # Basculante" já apontava pra essa mesma função (é a mesma coisa — caçamba = basculante), só
    # que com o nome antigo "Operador de caçamba"; a função no histograma mudou de nome pra
    # "Motorista de caçamba" e essa entrada tinha ficado pra trás (achado ao investigar por que o
    # Previsto de Basculante dava 0 — corrigido junto).
    'Motorista de Caminhao': 'Motorista de caçamba',
    'Motorista de Caminhão Basculante': 'Motorista de caçamba',
    'Motorista de caminhão comboio': 'Operador de Caminhão Comboio',
    'Motorista de Ônibus': None,
    'Op de Munck': 'Operado de Munk',  # mesma função de "Operador de Munck" — nome cru do histograma tem esse erro de digitação mesmo
    'Operador Betoneira': 'Operador de Betoneira',
    'Operador de Guindaste': 'Operador de Guindaste',
    # "Operador de Maquinas" é um Cargo Agrupado genérico que cobre gente que opera máquinas
    # diferentes (confirmado pessoa a pessoa com o arquivo do Gustavo: Escavadeira,
    # Retroescavadeira, Bob Cat, Perfuratriz, Trator de Pneus — os 3 últimos aqui ainda sem
    # ninguém real lançado, adicionados por consistência com o padrão dos outros). Por isso é uma
    # LISTA — o Previsto desse cargo soma o histograma de todas essas funções juntas.
    'Operador de Maquinas': ['Operador de Escavadeira', 'Operador de Retroescavadeira', 'Operador de Bob Cat',
                             'Operador de Perfuratriz', 'Operador de Trator de Pneus',
                             'Operador de Rolo Compactador', 'Operador Miniescavadeira', 'Operador Valetadeira'],
    'Operador de Motoniveladora': 'Operador de Motoniveladora',
    'Operador de Motoserra': None,
    'Operador de Munck': 'Operado de Munk',
    'Operador de Retroescavadeira': 'Operador de Retroescavadeira',
    'Operador de Trator de Esteira': 'Operador de Trator de Esteiras',
    'Pedreiro': 'Pedreiro',
    'Projetista PL - Projeto Civil': None,
    'Supervisor Segurança do Trabalho': None,
    # Mesmo caso: "Supervisor de Obras" cobre 2 pessoas com função de verdade diferente
    # (confirmado com o arquivo do Gustavo: Supervisor Civil e Supervisor de Montagem).
    'Supervisor de Obras': ['Supervisor Civil', 'Supervisor de Montagem'],
    'Supervisor de Planejamento': None,
    'Tecnico Seguranca do Trabalho': 'Tec. De Segurança',
    'Tecnico de Meio Ambiente': 'Tec. Meio Ambiente',
    'Tecnico de Qualidade': 'Tec. De Qualidade',
    'Topografo': 'Topógrafo',
    'Técnico de Qualidade': 'Tec. De Qualidade',
}

# ---- Função crua do histograma HIST-MO -> Setor (mesma taxonomia acima) ----
HIST_FUNC_TO_SETOR = {
    'Administrativo de Obra': 'Gestão/Administrativo',
    'Almoxarife': 'Gestão/Administrativo',
    'Encarregado Civil': 'Civil',
    'Engenheiro de Segurança': 'SSMA',
    'Encarregado Eletrica': 'Elétrica',
    'Supervisor Eletrica': 'Elétrica',
    'Encarregado Montagem': 'Montagem',
    'Supervisor de Montagem': 'Montagem',
    'Engenheiro Residente': 'Gestão/Administrativo',
    'Engenheiro': 'Gestão/Administrativo',
    'Coodenador Obras': 'Gestão/Administrativo',
    'Coodenador Administrativo': 'Gestão/Administrativo',
    'Coodenador de Projetos': 'Gestão/Administrativo',
    'Tec. Meio Ambiente': 'SSMA',
    'Tec. De Qualidade': 'Qualidade',
    'Tec. De Enfermagem': 'SSMA',
    'Tec. De Segurança': 'SSMA',
    'Gerente de Contrato': 'Gestão/Administrativo',
    'Assistente Social': 'Gestão/Administrativo',
    'Motorista de Ambulância': 'SSMA',
    'Segurança Patrimonial': 'Gestão/Administrativo',
    'Zeladora': 'Gestão/Administrativo',
    'Topógrafo': 'Topografia',
    'Auxiliares de Topografia': 'Topografia',
    'Médico do Trabalho': 'SSMA',
    'Operador Guindaste': 'Equipamentos',
    'Operador de Guindaste': 'Equipamentos',
    'Ajudante civil': 'Civil',
    'Supervisor Civil': 'Civil',
    'Pedreiro': 'Civil',
    'Poceiro': 'Civil',
    'Armador': 'Civil',
    'Pintor': 'Civil',
    'Técnico de Comissionamento': 'Qualidade',
    'Carpinteiro': 'Civil',
    'Encarregado Terraplanagem': 'Terraplanagem',
    'Operador de Escavadeira': 'Terraplanagem',
    'Operador de Trator de Esteira': 'Terraplanagem',
    'Operador de Trator de Esteiras': 'Terraplanagem',
    # Sem gente real lançada ainda nesses 3 — não tem como confirmar Terraplanagem x Equipamentos
    # com o arquivo do Gustavo (só tinha pessoa pros outros). Mantido em Equipamentos por
    # enquanto (mesmo padrão da maioria dos "Operador de X"); ajustar se/quando confirmado.
    'Operador de Rolo Compactador': 'Equipamentos',
    'Operador Miniescavadeira': 'Equipamentos',
    'Operador Valetadeira': 'Equipamentos',
    'Operador de Retroescavadeira': 'Equipamentos',
    'Operador de Trator de Pneus': 'Terraplanagem',
    'Operador de Caminhão Truck': 'Equipamentos',
    'Operador de Caminhão Comboio': 'Terraplanagem',
    'Operador de Caminhão Pipa': 'Transporte',
    'Motorista de Caminhão Pipa': 'Transporte',
    'Operador de Motoniveladora': 'Terraplanagem',
    'Operador de Betoneira': 'Equipamentos',
    'Operador de Bob Cat': 'Equipamentos',
    'Operador de Perfuratriz': 'Equipamentos',
    # Estava em "Transporte", mas as pessoas reais com esse cargo ("Motorista de
    # Caminhao"/"Motorista de Caminhão Basculante") estão classificadas como Terraplanagem na
    # coluna Setor da planilha — alinhado aqui pra Previsto e Realizado baterem no mesmo setor
    # (pedido do cliente: usar a classificação da planilha como fonte de verdade). Nome da chave
    # corrigido de "Operador de caçamba" (antigo, não existe mais no histograma) pra "Motorista de
    # caçamba" (nome atual) — a chave errada fazia esse Previsto cair silenciosamente em "Não
    # classificado" em vez de Terraplanagem (2026-08-26).
    'Motorista de caçamba': 'Terraplanagem',
    'Montador': 'Montagem',
    'Ajudante Montador': 'Montagem',
    'Eletricista Força Controle': 'Elétrica',
    'Eletricista Comum': 'Elétrica',
    'Ajudante elétrica': 'Elétrica',
    'Op de Munck': 'Equipamentos',
    'Operado de Munk': 'Equipamentos',
    'Apontador': 'Gestão/Administrativo',
}

# A aba HIST-MO tem duas linhas separadas pra função que é a mesma coisa (erro de digitação na
# planilha, não corrigido na fonte) — junta as duas nessa aqui na leitura, sem precisar editar o
# Excel. Formato: nome cru como está na planilha -> nome canônico que deve ser usado no lugar dele.
HIST_FUNC_ALIAS = {
    'Operador Guindaste': 'Operador de Guindaste',
}

# Mesma ideia, mas pra coluna "Função x SIENGE" da "Base de dados" x a descrição exata usada na
# aba "Relatório" do Sienge (BOQ) — nomes que são a mesma função na prática, só escritos diferente
# entre as duas planilhas. Formato: como está na coluna "Função x SIENGE" -> como está no Relatório.
SIENGE_FUNC_ALIAS = {
    'Encarregado de Montagem': 'Encarregado Montagem',
}
