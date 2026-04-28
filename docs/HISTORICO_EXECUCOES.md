# Histórico de Execuções — RAE Turbo

A partir desta versão, o RAE Turbo salva automaticamente um histórico de execuções em CSV, pensado para abrir direto no Excel ou LibreOffice.

## Local dos arquivos

Os arquivos ficam em:

```txt
AppData/Local/RAETurbo/relatorios/
```

Arquivos gerados:

```txt
historico_execucoes.csv
relatorio_YYYYMMDD_HHMMSS.txt
```

Além disso, o último resumo continua salvo em:

```txt
AppData/Local/RAETurbo/ultimo_relatorio_execucao.txt
```

## Métricas salvas

O CSV contém dados como:

- versão do RAE Turbo;
- status da execução;
- início e fim;
- tempo total em segundos;
- PDFs encontrados;
- PDFs processados;
- RAEs lançados com sucesso;
- PDFs movidos para destino;
- PDFs enviados para pasta de erros;
- CNPJs não encontrados;
- duplicidades barradas;
- falhas definitivas;
- tempo médio por RAE lançado;
- tempo médio por PDF processado.

## Uso recomendado

Use o arquivo `historico_execucoes.csv` para montar uma planilha de impacto com:

```txt
quantidade de execuções
quantidade de RAEs lançados
quantidade de PDFs processados
tempo total economizado
tempo médio por atendimento
taxa de erro
```

Esses dados ajudam a demonstrar produtividade e justificar piloto, expansão ou reconhecimento interno.
