# Política de Atualização — RAE Turbo

Última atualização: 27/04/2026

## Processo

1. Alterar o código.
2. Testar localmente.
3. Atualizar `version.py`.
4. Atualizar `CHANGELOG.md`.
5. Commitar as alterações.
6. Criar tag `vX.X.X`.
7. Enviar a tag para o GitHub.
8. O GitHub Actions gera e publica `RAE_Turbo.exe`.

## Checklist antes de publicar

- Programa abre corretamente.
- Configuração da unidade salva e carrega.
- Bases em `data/` são encontradas.
- ChromeDriver é encontrado.
- Atualizador reconhece a versão.
- Logs não expõem CNPJ completo.
