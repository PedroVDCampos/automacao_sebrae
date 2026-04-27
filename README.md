# RAE Turbo

**RAE Turbo** é uma ferramenta de automação criada para otimizar a rotina de atendimento MEI, organizando PDFs de atendimentos e automatizando o preenchimento do RAE — Registro de Ação Empreendedora.

> Versão deste pacote: `v1.1.7`  
> Última atualização da documentação: 27/04/2026

## Funcionalidades principais

- Interface gráfica para usuários não técnicos.
- Organização automática dos PDFs.
- Extração de CNPJ e nome do cliente.
- Automação do fluxo completo no RAE.
- Configuração por unidade, projeto, ação e ano.
- Base mapeada de unidades/projetos/ações do Sebrae-SP.
- Logs em `AppData/Local/RAETurbo`.
- Atualização automática via GitHub Releases.
- Build automatizado por GitHub Actions.

## Melhorias aplicadas

- Versão centralizada em `version.py`.
- Executável padronizado como `RAE_Turbo.exe`.
- Bases JSON movidas para `data/`.
- Scripts de mapeamento movidos para `scripts/`.
- Configuração salva em `AppData/Local/RAETurbo/config_unidade.json`.
- Pausa de login agora bloqueia a execução até o usuário confirmar.
- `erro_fatal` tratado na interface.
- CNPJ mascarado em logs e mensagens de erro.
- Atualizador mais robusto.

## Desenvolvimento

```bash
pip install -r requirements.txt
python main.pyw
```

## Nova versão

```bash
git add .
git commit -m "prepara versão v1.1.7"
git tag v1.1.7
git push
git push origin v1.1.7
```
