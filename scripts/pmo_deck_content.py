# -*- coding: utf-8 -*-
"""Conteudo das paginas 1 e 3 do deck de status das quatro frentes.

A pagina 2 usa `pmo_content.py`, que reproduz o print original.

Status aceitos (mesmas chaves de pmo_status.py):
    'ok'      -> check verde           concluido
    'plano'   -> tres pontos verdes    em andamento, dentro do plano
    'risco'   -> tres pontos amarelos  parado em parceiro / sem prazo
    'atraso'  -> tres pontos vermelhos em andamento, fora do plano

Cada marco e um dict:
    feito  o que precisava ser feito (em negrito na celula)
    qual   complemento do titulo, em cinza e sem negrito (opcional)
    prazo  data ou TBD
    nota   observacao curta em vermelho ao lado do prazo (opcional)
    status chave do semaforo
    prog   progresso recente
    prox   proximos passos
    esc    pontos a escalar (so na pagina 1)
"""

NA = 'NA'
TBD = 'TBD'

# ---------------------------------------------------------------- pagina 1
# Reproduz a pagina HTML: identidade por frente com o color code do plano de
# escalada, responsavel e valor gerado na faixa da esquerda.

P1_TITULO_BOLD = 'Acompanhamento das 4 frentes de F&I:'
P1_TITULO = (' 3 de 10 entregáveis concluídos, 2 fora do plano e 3 travados '
             'em terceiros')
P1_CHIP = 'ACOMPANHAMENTO'
P1_MARCA = '/ PRELIMINAR'
P1_FONTE = ('Fonte: memorando Lead to Sales — Aceleração Seminovos e plano de '
            'escalada de F&I (preliminar) · Ref.: 03/set/26 · '
            'TBD = prazo não acordado; TBC = valor a confirmar')
P1_CLASSE = 'Classificação da Informação: INTERNA'

P1_LEGENDA = (('ok', 'Concluído', 3),
              ('plano', 'Em andamento, dentro do plano', 2),
              ('atraso', 'Em andamento, fora do plano', 2),
              ('risco', 'Bloqueado em terceiros', 3))

P1 = [
    dict(nome='Simulador de pré-análise', resp='Sônia e Ricardo',
         valor='R$ 29–74 M', marcos=[
        dict(feito='Desenhar plano de comunicação', prazo='02/set',
             status='ok',
             prog='Formato, data, participantes e duração definidos para cada conversa',
             prox='Implementar o plano de comunicação a partir de 08/set',
             esc=NA),
        dict(feito='Ajustar tela do simulador',
             qual=' (textos e card de crédito parcial)',
             prazo='02/set', nota='prazo em redefinição', status='atraso',
             prog='Desenvolvimentos preliminares concluídos',
             prox='Implementar os ajustes observados durante a visita em loja',
             esc='Avaliar dificuldade e prazos para implementar todas as melhorias mapeadas'),
        dict(feito='Selecionar lojas do piloto', prazo='02/set', status='ok',
             prog='Lojas selecionadas nas regionais Nordeste, MG, CO, ES e RJ; '
                  '75 lojas escolhidas para o piloto',
             prox=NA, esc=NA),
    ]),
    dict(nome='Modelo de distribuição de fichas', resp='Sandra Reis',
         valor='R$ 24–52 M', marcos=[
        dict(feito='Iniciar modelo escalável em 2 lojas',
             qual=' (VCGRU e VCSJS)', prazo='01/set', status='ok',
             prog='Modelo implementado na VCGRU e na VCSJS',
             prox='Expandir gradualmente para as demais lojas que já fazem parte do piloto',
             esc=NA),
        dict(feito='Implantar sistema de reanálise nos modelos',
             prazo='25/set', status='plano', prog=NA, prox=NA, esc=NA),
    ]),
    dict(nome='Ecossistema de F&I para lojistas', resp='Thiago Lopes',
         valor='TBC', marcos=[
        dict(feito='Integrar o PAN no FAS', prazo=TBD, status='risco',
             prog=NA, prox=NA,
             esc='Definir quem fará o desenvolvimento: internamente ou do lado do PAN'),
        dict(feito='Finalizar os desenvolvimentos de tecnologia necessários do VW',
             prazo=TBD, status='risco', prog=NA,
             prox='Garantir que os nossos desenvolvimentos sejam priorizados junto ao VW',
             esc=NA),
        dict(feito='Resolver as inconsistências nas ofertas do Bradesco',
             prazo=TBD, status='risco', prog=NA,
             prox='Garantir que os problemas de instabilidade sejam resolvidos',
             esc=NA),
    ]),
    dict(nome='Gestão integrada do ciclo de crédito',
         resp='Bárbara Gabrielle e Lucas Ávila', valor='TBC', marcos=[
        dict(feito='Levantar os dados de Bradesco e VW', prazo='02/set',
             status='plano', prog=NA, prox=NA, esc=NA),
        dict(feito='Criar score de crédito Localiza', prazo=TBD,
             status='atraso',
             prog='Acessos concedidos e bases entendidas',
             prox='Estruturar as bases, treinar o modelo e fazer backtest com '
                  'dados históricos para avaliar dispersão e acurácia',
             esc=NA),
    ]),
]

P1_BOLD = {
    'Lojas selecionadas nas regionais Nordeste, MG, CO, ES e RJ; '
    '75 lojas escolhidas para o piloto': ['75 lojas'],
    'Modelo implementado na VCGRU e na VCSJS': ['VCGRU', 'VCSJS'],
    'Acessos concedidos e bases entendidas': ['Acessos concedidos'],
    'Definir quem fará o desenvolvimento: internamente ou do lado do PAN':
        ['internamente ou do lado do PAN'],
}

# ---------------------------------------------------------------- pagina 3
# Combina as duas: color code e identidade da frente da pagina 1, titulo
# construido a partir das contagens e contagem por frente da pagina 2.
# A coluna "Pontos a escalar" sai da tabela (7 das 10 celulas eram NA) e vira
# a faixa de decisoes pedidas no pe da pagina.

P3_TITULO_BOLD = '3 dos 10 marcos concluídos;'
P3_TITULO = (' 2 estão fora do plano por falta de prazo e 3 parados em PAN, '
             'VW e Bradesco')
P3_SUB = 'Status das quatro frentes de F&I  |  posição de 03/set/26'
P3_FONTE = ('Fonte: memorando Lead to Sales — Aceleração Seminovos e plano de '
            'escalada de F&I (preliminar) · TBD = prazo não acordado; '
            'TBC = valor a confirmar')
P3_CLASSE = 'Classificação da Informação: INTERNA'

P3_LEGENDA = (('ok', 'Concluído'), ('plano', 'No plano'),
              ('risco', 'Parado em parceiro'), ('atraso', 'Fora do plano'))
P3_RESUMO = {'ok': 'concluído', 'plano': 'no plano',
             'risco': 'parado em parceiro', 'atraso': 'fora do plano'}

P3 = [
    dict(nome='Simulador de pré-análise', resp='Sônia e Ricardo',
         valor='R$ 29–74 M', marcos=[
        dict(feito='Desenhar plano de comunicação', prazo='02/set',
             status='ok',
             prog='Formato, data, participantes e duração definidos para cada conversa',
             prox='Implementar o plano a partir de 08/set'),
        dict(feito='Ajustar tela do simulador, textos e cards de crédito parcial',
             prazo='02/set', nota='vencido', status='atraso',
             prog='Desenvolvimentos preliminares concluídos, mas a visita em '
                  'loja mapeou melhorias adicionais',
             prox='Dimensionar o esforço das melhorias e repactuar o prazo'),
        dict(feito='Selecionar lojas do piloto', prazo='02/set', status='ok',
             prog='75 lojas escolhidas nas regionais Nordeste, MG, CO, ES e RJ',
             prox='Iniciar a escalada da expansão em 08/set'),
    ]),
    dict(nome='Modelo de distribuição de fichas', resp='Sandra Reis',
         valor='R$ 24–52 M', marcos=[
        dict(feito='Iniciar modelo escalável em 2 lojas', prazo='01/set',
             status='ok',
             prog='Modelo rodando na VCGRU e na VCSJS',
             prox='Expandir para as demais lojas do piloto e medir o resultado '
                  'do novo sistema'),
        dict(feito='Implantar sistema de reanálise nos modelos',
             prazo='25/set', status='plano',
             prog='Não iniciado, dentro da janela do prazo',
             prox='Detalhar o escopo até 11/set para preservar a data de 25/set'),
    ]),
    dict(nome='Ecossistema de F&I para lojistas', resp='Thiago Lopes',
         valor='TBC', marcos=[
        dict(feito='Integrar o PAN no FAS', prazo=TBD, status='risco',
             prog='Sem avanço: escopo de desenvolvimento não definido',
             prox='Definir o responsável pelo desenvolvimento'),
        dict(feito='Finalizar desenvolvimentos de tecnologia do VW',
             prazo=TBD, status='risco',
             prog='Sem avanço: demanda na fila de priorização do parceiro',
             prox='Priorizar a nossa demanda junto ao VW'),
        dict(feito='Resolver inconsistências nas ofertas do Bradesco',
             prazo=TBD, status='risco',
             prog='Instabilidade recorrente nas ofertas apresentadas em loja',
             prox='Cobrar plano de correção do Bradesco'),
    ]),
    dict(nome='Gestão integrada do ciclo de crédito',
         resp='Bárbara Gabrielle e Lucas Ávila', valor='TBC', marcos=[
        dict(feito='Levantar os dados de Bradesco e VW', prazo='02/set',
             status='plano',
             prog='Coleta em andamento com os dois parceiros',
             prox='Fechar a base consolidada'),
        dict(feito='Criar score de crédito Localiza', prazo=TBD,
             nota='sem data', status='atraso',
             prog='Acessos concedidos e bases entendidas',
             prox='Estruturar as bases, treinar o modelo e fazer backtest para '
                  'medir a acurácia'),
    ]),
]

P3_BOLD = {
    '75 lojas escolhidas nas regionais Nordeste, MG, CO, ES e RJ': ['75 lojas'],
    'Modelo rodando na VCGRU e na VCSJS': ['VCGRU', 'VCSJS'],
    'Acessos concedidos e bases entendidas': ['Acessos concedidos'],
    'Desenvolvimentos preliminares concluídos, mas a visita em loja mapeou '
    'melhorias adicionais': ['melhorias adicionais'],
    'Sem avanço: escopo de desenvolvimento não definido': ['Sem avanço'],
    'Sem avanço: demanda na fila de priorização do parceiro': ['Sem avanço'],
}

# faixa de decisoes pedidas: (frente, decisao, responsavel)
P3_DECISOES_TITULO = 'Decisões pedidas nesta reunião'
P3_DECISOES = [
    ('Simulador de pré-análise',
     'Aprovar o escopo das melhorias mapeadas em loja e repactuar a data do '
     'ajuste de tela',
     'Sônia e Ricardo'),
    ('Ecossistema de F&I para lojistas',
     'Definir se o desenvolvimento do PAN é interno ou do parceiro e escalar a '
     'priorização com VW e Bradesco',
     'Thiago Lopes'),
    ('Gestão integrada do ciclo de crédito',
     'Acordar a data de entrega do score de crédito Localiza',
     'Bárbara Gabrielle e Lucas Ávila'),
]
