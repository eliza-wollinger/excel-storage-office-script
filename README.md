# Controle Automatizado de Estoque de EPIs (Botas e CAs)

## Visão Geral
Este projeto implementa um sistema automatizado de controle de estoque de EPIs, com foco em botas de segurança e seus respectivos Certificados de Aprovação (CA). 📦

O processo é totalmente desenvolvido em Excel Online, utilizando Office Script para garantir que as atualizações de entradas, retiradas, histórico e monitoramento sejam feitas de forma automática e padronizada.


## Objetivo
Garantir o controle eficiente e rastreável das movimentações de estoque, principalmente botas de segurança e uniformes EPI, assegurando a conformidade com as normas de Segurança do Trabalho e o acompanhamento das validades de CA.

## Estrutura do Arquivo
O arquivo principal contém as seguintes planilhas:

- Entradas – registro de novos lotes de botas, com informações de modelo, tamanho, quantidade, fornecedor, data de recebimento e número/validade do CA.
- Retiradas – registro de EPIs entregues aos colaboradores, incluindo nome, setor, modelo, data e motivo da retirada.
- Histórico – consolidado automático de todas as movimentações (entradas e retiradas), com rastreabilidade de cada item.
- Estoque Atual – resumo dinâmico com o saldo de botas disponíveis por modelo, tamanho e CA.
- Monitoramento – painel de acompanhamento com indicadores como:
    - Quantidade total em estoque;
    - Quantidade em uso;
    - Percentual de EPIs com CA próximo ao vencimento;
    - Histórico mensal de movimentações.


## Funcionamento da Automação
O Office Script é responsável por:

1. Ler as tabelas “Entradas” e “Retiradas”
    - Identifica novas movimentações não processadas.

2. Atualizar o Histórico
    - Registra automaticamente as novas entradas e retiradas com data, tipo e detalhes do EPI.

3. Atualizar o Estoque Atual
    - Recalcula o saldo por modelo, tamanho e CA.
    - Verifica e sinaliza CAs vencidos ou próximos do vencimento.

4. Atualizar o Monitoramento
    - Atualiza indicadores de controle e gráficos automáticos.

5. Prevenir Duplicidades
    - Marca as movimentações já processadas, evitando recontagem.


## Tecnologias Utilizadas
- Microsoft Excel Online
- Office Script (TypeScript)


## Como Executar

1. Abra o arquivo no Excel Online.
2. Acesse a guia Automatizar > Meus scripts.
3. Selecione o script Atualizar_Estoque_EPIs.
4. Clique em Executar.
5. Aguarde a atualização das planilhas “Histórico”, “Estoque Atual” e “Monitoramento”.

💡 É possível agendar a execução automática via Power Automate, integrando o fluxo para rodar periodicamente (ex.: diariamente ou semanalmente).


## Logs e Alertas
- O script exibe mensagens no console de execução informando o andamento e eventuais inconsistências (ex.: CA vencido, campos vazios, duplicidade).
- Linhas com erro permanecem pendentes até correção.


## Manutenção Preventiva
- Verificar mensalmente se os CAs estão dentro da validade.
- Atualizar o script conforme novas regras de controle de EPI ou formato de planilha.
- Revisar permissões de acesso no SharePoint.