## [v1.2.4]

### Corrigido
- Corrigido o fluxo de atualização automática.
- O updater agora baixa o executável da release antes de aplicar a atualização.
- Corrigida chamada incorreta da função `aplicar_atualizacao`.
- Adicionada validação de tamanho do arquivo baixado.
- Adicionada validação opcional de SHA-256 quando disponível.
- Mantido reinício com `PYINSTALLER_RESET_ENVIRONMENT=1` para evitar erro de DLL temporária do PyInstaller.