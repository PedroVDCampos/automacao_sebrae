## [v1.2.3]

### Corrigido
- Corrigido erro `Failed to load Python DLL` após atualização automática.
- O reinício após atualização agora usa `PYINSTALLER_RESET_ENVIRONMENT=1` para evitar reutilização incorreta dos arquivos temporários do PyInstaller.
- Corrigido nome exato do serviço de parcelamento no RAE.
- Ajustada seleção de serviço para evitar múltiplas marcações quando o RAE exibe opções duplicadas com o mesmo nome.