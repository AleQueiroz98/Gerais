# -*- coding: utf-8 -*-
"""Conteudo do acompanhamento de agosto/26.

Frente 2: merge da pagina atual com o plano de qualificacao e alocacao de
leads (cronograma do PI Planning). Frentes 4 e 5: status e progresso do mes.
Cada milestone e uma tupla:
    (id, descricao, [trechos em negrito], responsavel, prazo, status,
     progresso recente, proximos passos, pontos a escalar)
"""
TBD = 'TBD'

FRENTE_2A = [   # pagina 1: qualificacao de leads
    ('2.1.1',
     'Informações mínimas do cliente coletadas no site: validação dos dados coletados e 2ª página com a loja de preferência',
     ['Informações mínimas do cliente coletadas no site'],
     'Thais Pimenta', 'S2 set/26', 'amarelo',
     'PI Planning realizado e solicitação do CEP obrigatório encaminhada',
     'Validar os dados coletados (S1 set) e publicar a 2ª página com a loja de preferência (S2 set)',
     'Flexibilização do período de congelamento do desenvolvimento a ser discutida'),

    ('2.1.2',
     'Informações mínimas do cliente coletadas no Facebook e no WhatsApp: CEP capturado para todos os leads e 2ª página com a loja de preferência',
     ['Informações mínimas do cliente coletadas no Facebook e no WhatsApp'],
     'Thais Pimenta', 'S4 set/26', 'amarelo',
     'Captura de CEP para todos os leads em implementação nos canais Meta',
     'Publicar a 2ª página com a loja de preferência (S3–S4 set)', TBD),

    ('2.1.3',
     'Informações mínimas do cliente coletadas via canais classificados: CEP capturado para todos os leads',
     ['Informações mínimas do cliente coletadas via canais classificados'],
     'Thais Pimenta', 'S3 set/26', 'amarelo',
     'Captura de CEP em implementação nos canais classificados',
     'Concluir a captura em todos os classificados (S3 set)', TBD),

    ('2.2.1',
     'Informações capturadas no site (ex.: localização por IP) integradas e disponíveis para todos os canais de atendimento',
     ['Informações capturadas no site', 'integradas e disponíveis'],
     'Luma Gomes', 'S4 set/26', 'vermelho', TBD, TBD, TBD),

    ('2.3.1',
     'Lead score evoluído com novas variáveis (score B3 de propensão a financiamento, renda presumida etc.)',
     ['Lead score evoluído'],
     'João Azevedo', 'S3 out/26', 'vermelho', TBD, TBD, TBD),

]

FRENTE_2B = [   # pagina 2: alocacao de leads
    ('2.4.1',
     'Modelo simplificado de alocação (v1.0) desenvolvido e piloto lançado em 30 lojas',
     ['Modelo simplificado de alocação (v1.0) desenvolvido e piloto lançado em 30 lojas'],
     'Luma Gomes e Yure Cangussu', 'S4 ago/26', 'verde',
     'Piloto lançado em 30 lojas; definição do grupo controle em andamento',
     'Concluir a análise preliminar do piloto (S1 set)', TBD),

    ('2.4.2',
     'Piloto do modelo v1.0 escalado para 100 lojas (11/09), com análise dos resultados',
     ['Piloto do modelo v1.0 escalado para 100 lojas'],
     'Luma Gomes e Yure Cangussu', 'S2 set/26', 'amarelo',
     'Análise preliminar do piloto em 30 lojas em andamento',
     'Escalar para 100 lojas em 11/09 e analisar os resultados (S3–S4 set)', TBD),

    ('2.4.3',
     'Modelo v1.0 escalado para as lojas restantes e resultados acompanhados até a desativação (S2 nov/26)',
     ['Modelo v1.0 escalado para as lojas restantes'],
     'Luma Gomes', 'S1 out/26', 'vermelho', TBD, TBD, TBD),

    ('2.4.4',
     '~5 pilotos do modelo robusto desenhados, combinando loja escolhida, distância, canal de origem e dados do carro (modelo, ano, valor e faixa de km)',
     ['~5 pilotos do modelo robusto desenhados'],
     'Yure Cangussu e João Azevedo', 'S1 set/26', 'amarelo',
     'Desenho dos pilotos em andamento, com as variáveis relevantes mapeadas',
     'Concluir o desenho dos ~5 pilotos e iniciar o desenvolvimento (S2 set)', TBD),

    ('2.4.5',
     '~5 novos pilotos desenvolvidos e lançados, com 20 lojas em cada',
     ['~5 novos pilotos desenvolvidos e lançados'],
     'Yure Cangussu e João Azevedo', 'S4 set/26', 'vermelho', TBD, TBD, TBD),

    ('2.4.6',
     'Top 3 modelos testados em 33 lojas cada e modelo vencedor escalado para todas as lojas',
     ['Top 3 modelos testados em 33 lojas cada', 'modelo vencedor escalado para todas as lojas'],
     'Yure Cangussu e João Azevedo', 'S2 nov/26', 'vermelho', TBD, TBD, TBD),

    ('2.4.7',
     'Piloto de alocação total dos leads na Liza, com transbordo para a Central, lançado',
     ['Piloto de alocação total dos leads na Liza'],
     'Luma Gomes e Yure Cangussu', 'S4 out/26', 'vermelho', TBD, TBD, TBD),

    ('2.4.8',
     'Modelo de clusterização revisado com base nas distâncias entre lojas e leads alocados',
     ['Modelo de clusterização revisado'],
     'João Azevedo', 'S2 set/26', 'vermelho', TBD, TBD, TBD),
]

FRENTE_4 = [
    ('4.1.1',
     'Melhores práticas de atendimento na Central mapeadas',
     ['Melhores práticas de atendimento na Central mapeadas'],
     'Luis Dutra', 'S4 ago/26', 'amarelo',
     'Templates e disparos manuais revisitados; jornada de persistência revisitada, em desenvolvimento pelo Labs',
     'Consolidar as melhores práticas mapeadas e difundi-las para os RDVs', TBD),

    ('4.2.1',
     'Modelo de treinamento e nova rotina de gestão de performance dos RDVs evoluídos',
     ['Modelo de treinamento e nova rotina de gestão de performance'],
     'Rômulo Rodrigo', 'S3 set/26', 'amarelo',
     'Jornada de passagem de bastão entre RDV e CV desenhada e refinada; treinamento estruturado para o piloto do Varejo',
     'Desenvolver a jornada de passagem de bastão nas sprints 1 e 2', TBD),

    ('4.3.1',
     'Teste com contratação de atendentes mais qualificados (classe 13, nível de EVs, vs. classe 10, nível de CVs) realizado',
     ['Teste com contratação de atendentes mais qualificados'],
     'Eloah Aguiar e Rômulo Rodrigo', 'S3 set/26', 'vermelho',
     'Sem avanços reportados em agosto', TBD, TBD),

    ('4.4.1',
     'Rotina de gestão de performance implementada',
     ['Rotina de gestão de performance implementada'],
     'Vitor Sueth', 'S4 ago/26', 'amarelo',
     'Atuação em leads muito frios excluída, com envio para a Liza Scout: ganho de produtividade por RDV, com baixa queda na conversão em agendamento e em comparecimento',
     'Acompanhar o efeito da exclusão dos leads frios sobre as conversões', TBD),

    ('4.5.1',
     'Piloto de atendimento da Central em lojas do varejo lançado',
     ['Piloto de atendimento da Central em lojas do varejo lançado'],
     'Rômulo Rodrigo', 'S2 set/26', 'amarelo',
     'Estruturação e treinamento concluídos para o início em 08/09',
     'Iniciar o piloto em 08/09 nas lojas VCLEX e VCSBH', TBD),

    ('4.5.2',
     'Piloto de atendimento da Central em lojas do varejo expandido (condicionado a resultados favoráveis)',
     ['Piloto de atendimento da Central em lojas do varejo expandido'],
     'Rômulo Rodrigo', 'S4 dez/26', 'vermelho', TBD, TBD, TBD),
]

FRENTE_5 = [
    ('5.1.1',
     'Solução de atendimento via WhatsApp Business integrado ao Slack expandida para todas as lojas do varejo, em ondas mensais (1 loja jul., 28 ago., 80 set., 156 out., 235 nov., 256 dez.)',
     ['Solução de atendimento via WhatsApp Business integrado ao Slack expandida',
      'todas as lojas do varejo'],
     'Isadora Savaget e Denys Damm', 'dez/26 (TBC)', 'amarelo',
     'Melhorias operacionais na Slackbot (consumo de estoque e coaching). Visita da Diretoria às 4 lojas piloto para mapear dores e oportunidades. Resultados iniciais: ganho relevante de TME, retorno no D0 de ~54%, conectividade geral +70% e alto volume de mensagens, com baixa ou nenhuma instabilidade da plataforma',
     'Endereçar as dores mapeadas nas 4 lojas piloto, com menor robotização dos templates, e seguir com as ondas de expansão',
     'Ligações via plataforma ainda não permitidas pela Meta'),

    ('5.2.1',
     'Inteligência de gestão da performance dos CVs implementada, combinando indicadores do funil e dados qualitativos via Alethe.IA',
     ['Inteligência de gestão da performance dos CVs implementada'],
     'Isadora Savaget', 'S4 set/26', 'amarelo',
     'Estruturação e refinamento dos dados e dos novos indicadores viabilizadores da Alethe.IA em andamento',
     'Concluir a base de indicadores e conectar a Alethe.IA à gestão de performance', TBD),

    ('5.3.1',
     'Ritos e rotina de gestão das equipes de loja revisados e implantados',
     ['Ritos e rotina de gestão das equipes de loja revisados e implantados'],
     'Pablo Chaves', 'S4 set/26', 'vermelho',
     'Sem avanços reportados em agosto; coaching endereçado parcialmente pelas melhorias na Slackbot',
     TBD, TBD),
]
