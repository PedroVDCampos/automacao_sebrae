## [v1.2.0]

### Melhorado
- PDFs agora só são enviados para a pasta final após confirmação de lançamento no RAE.
- PDFs com erro são movidos para `_ERROS_RAE_TURBO` dentro da pasta de origem.
- Separação de erros por motivo: CNPJ não encontrado, falha definitiva, duplicidade e dados insuficientes.

### Corrigido
- Evita que PDFs com atendimento não registrado sejam arquivados como se tivessem sido concluídos.