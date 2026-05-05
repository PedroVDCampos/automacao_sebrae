# RAE Turbo

**RAE Turbo** é uma ferramenta de automação criada para otimizar a rotina de atendimento MEI, organizando PDFs de atendimentos e automatizando o preenchimento do RAE — Registro de Ação Empreendedora.

> Versão deste pacote: `v1.2.5`  
> Última atualização da documentação: 05/05/2026

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
