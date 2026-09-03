# -*- coding: utf-8 -*-
"""Conteudo da pagina de status das quatro frentes (PMO).

Para editar a pagina, altere apenas este arquivo e rode `python3 pmo_status.py`.

Status aceitos:
    'ok'       -> check verde          (concluido)
    'plano'    -> tres pontos verdes   (em andamento, dentro do plano)
    'risco'    -> tres pontos amarelos (em andamento, com riscos)
    'atraso'   -> tres pontos vermelhos(em andamento, fora do plano)

Cada marco e uma tupla:
    (o que precisava ser feito, prazo, status, progresso recente,
     proximos passos, pontos a escalar)
"""

NA = 'NA'
TBD = 'TBD'

# os contadores sao preenchidos a partir dos status, para nao desalinhar
TITULO = ('{ok} dos {tot} marcos concluídos; simulador e score de crédito fora do '
          'plano e {risco} integrações com parceiros sem prazo')

SUBTITULO = 'Status das quatro frentes  |  posição de 03/set/26'
RODAPE = ('Nota: prazos em TBD dependem de definição conjunta com os parceiros (PAN, VOX e Bradesco). '
          'NA = não aplicável no ciclo atual.')

FRENTES = [
    ('Simulador de pré-análise', [
        ('Desenhar plano de comunicação',
         '02/set/26', 'ok',
         'Plano de comunicação desenhado, com formato, data, participantes e '
         'duração definidos para cada conversa',
         NA, NA),
        ('Ajustar tela do simulador, textos e cards de crédito parcial',
         '02/set/26', 'atraso',
         'Desenvolvimentos preliminares concluídos',
         'Implementar ajustes observados durante visita em loja',
         'Avaliar dificuldade e prazos para implementar todas as melhorias mapeadas'),
        ('Selecionar lojas do piloto',
         '02/set/26', 'ok',
         'Lojas selecionadas nas regionais Nordeste, MG, CO, ES e RJ; 75 lojas '
         'escolhidas para o piloto',
         NA, NA),
    ]),
    ('Modelo de distribuição de fichas', [
        ('Iniciar modelo escalável em duas lojas',
         '01/set/26', 'ok',
         'Modelo implementado na VCGRU e na VCSCJS',
         'Expandir gradualmente para as outras lojas que já fazem parte do piloto',
         NA),
        ('Implementar sistema de reanálise nos modelos',
         '25/set/26', 'plano',
         NA, NA, NA),
    ]),
    ('Ecossistema de financiamento para lojistas', [
        ('Integrar o PAN no FAS',
         TBD, 'risco',
         NA, NA,
         'Definir quem fará o desenvolvimento: internamente ou do lado do PAN'),
        ('Finalizar desenvolvimentos de tecnologia necessários do VOX',
         TBD, 'risco',
         NA,
         'Garantir que nossos desenvolvimentos de tecnologia sejam priorizados com o VOX',
         NA),
        ('Resolver inconsistência nas ofertas do Bradesco',
         TBD, 'risco',
         NA,
         'Garantir que os problemas de instabilidade sejam resolvidos',
         NA),
    ]),
    ('Gestão integrada do ciclo de crédito', [
        ('Obter dados de Bradesco e VOX',
         TBD, 'plano',
         NA, NA, NA),
        ('Criar score de crédito Localiza',
         TBD, 'atraso',
         'Acessos concedidos e bases entendidas',
         'Estruturação e consolidação das bases, treinamento do modelo e backtest '
         'com dados históricos para avaliar a acurácia',
         NA),
    ]),
]

# trechos em negrito seletivo (mesmo padrao do deck de frentes)
BOLD = {
    'Plano de comunicação desenhado, com formato, data, participantes e '
    'duração definidos para cada conversa': ['Plano de comunicação desenhado'],
    'Lojas selecionadas nas regionais Nordeste, MG, CO, ES e RJ; 75 lojas '
    'escolhidas para o piloto': ['75 lojas'],
    'Modelo implementado na VCGRU e na VCSCJS': ['VCGRU', 'VCSCJS'],
    'Acessos concedidos e bases entendidas': ['Acessos concedidos'],
    'Definir quem fará o desenvolvimento: internamente ou do lado do PAN':
        ['internamente ou do lado do PAN'],
}
