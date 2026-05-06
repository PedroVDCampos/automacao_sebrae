## [v1.2.8]

### Melhorado
- Ajustado o cálculo de tempo médio dos atendimentos.
- O RAE Turbo agora mede o tempo real de lançamento de cada atendimento, iniciando no momento em que o CNPJ começa a ser digitado no RAE e encerrando no clique em “Finalizar atendimento”.
- Atendimentos com erro, CNPJ não encontrado, cadastro pendente, pessoa jurídica inativa, serviço não encontrado ou falha definitiva não entram na média real.
- O resumo de execução agora armazena tempo total real, média real, menor tempo e maior tempo de atendimentos lançados com sucesso.

### Técnico
- Padronizado o retorno da função `registrar_no_rae`.
- Adicionada normalização de retorno no `orquestrador.py` para manter compatibilidade com retornos antigos.
- Adicionadas métricas específicas para tempos reais de atendimentos concluídos.