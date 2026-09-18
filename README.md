# Turbo

## Organizador automatizado de documentos de atendimento

O Turbo identifica documentos PDF, extrai as informações necessárias e organiza os arquivos automaticamente por ano, mês, tipo de atendimento e cliente.

### Funcionalidades
- Identificação automática de PDFs.
- Extração de nome e CNPJ.
- Classificação de documentos.
- Diferenciação de CCMEI entre Formalização e Alteração pela Data de Abertura.
- Organização por ano, mês, tipo e cliente.
- Tratamento de documentos com erro.
- Histórico de execuções em CSV.
- Logs locais.
- Atualização automática via GitHub Releases.

### Execução
```bash
pip install -r requirements.txt
python main.pyw
```

A versão 2.0.0 não depende do Chrome, ChromeDriver ou do sistema Sebrae.
