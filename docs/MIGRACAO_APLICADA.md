# Migração aplicada neste pacote

## Alterações principais

- `base_dados_acoes.json` e `base_dados_projetos.json` foram movidos para `data/`.
- `mapeador_acoes.py` e `mapeador_sebrae.py` foram movidos para `scripts/`.
- `config_unidade.json` foi movido para `docs/config_unidade.exemplo.json`.
- A configuração real do usuário agora é salva em `AppData/Local/RAETurbo/config_unidade.json`.
- O log é salvo em `AppData/Local/RAETurbo/rae_turbo_execucao.log`.
- O executável foi padronizado como `RAE_Turbo.exe`.
- A versão foi centralizada em `version.py`.

## Atenção

Este pacote foi refatorado sem execução real no ambiente do RAE. Teste primeiro em ambiente controlado antes de usar em produção.
