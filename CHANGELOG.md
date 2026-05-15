## [v1.3.1] - 2026-05-15

### Adicionado
- Adicionada leitura da **Data de Abertura** em arquivos CCMEI.
- Adicionada função específica para interpretar arquivos CCMEI e retornar nome, CNPJ e data de abertura.
- Adicionada classificação automática entre **Formalização** e **Alteração** para arquivos CCMEI.
- Adicionada pasta de erro específica para CCMEI quando a Data de Abertura não puder ser identificada.

### Melhorado
- Melhorada a precisão na identificação do tipo de atendimento dos arquivos CCMEI.
- O RAE Turbo agora diferencia formalização e alteração comparando a Data de Abertura do MEI com a data do arquivo.
- Arquivos CCMEI não são mais classificados apenas pelo nome do arquivo.
- Reduzido o risco de registrar atendimentos CCMEI com o serviço incorreto no RAE.
- Melhorado o cálculo do tempo médio real dos atendimentos lançados.
- O tempo médio real agora considera apenas atendimentos concluídos com sucesso no RAE.

### Corrigido
- Corrigido problema em que arquivos CCMEI de alteração poderiam ser tratados incorretamente como formalização.
- Corrigido problema em que arquivos CCMEI com o mesmo padrão de nome eram classificados de forma imprecisa.
- Corrigido cálculo de tempo médio que considerava o tempo total da execução, incluindo pausas, login, reinicializações e demoras externas ao lançamento real.

### Técnico
- Criadas funções auxiliares no `extrator_pdf.py` para extração da Data de Abertura do CCMEI.
- Atualizada a lógica do `orquestrador.py` para classificar CCMEI com base no conteúdo do PDF.
- Adicionada proteção para impedir lançamento automático quando a Data de Abertura não for encontrada.
- Mantida compatibilidade com o fluxo de envio de PDFs com erro para `_ERROS_RAE_TURBO`.

### Observações
- Para arquivos CCMEI, a classificação agora segue a regra:
  - Data de Abertura igual à data do arquivo: **Formalização**.
  - Data de Abertura diferente da data do arquivo: **Alteração**.
- Recomenda-se validar se o serviço exato de alteração no RAE está configurado corretamente como `MEI - Alteração do MEI`.