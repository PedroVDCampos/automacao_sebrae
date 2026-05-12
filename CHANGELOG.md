## [v1.3.0]

### Adicionado
- Download automático do ChromeDriver compatível com a versão instalada do Google Chrome.
- Verificação automática de compatibilidade entre Chrome e ChromeDriver antes da execução.
- Fallback para o ChromeDriver embutido caso ele ainda seja compatível.

### Melhorado
- Reduzida a necessidade de lançar nova versão do RAE Turbo apenas por atualização do Google Chrome.
- Mensagens de erro mais claras quando não for possível preparar o ChromeDriver automaticamente.

### Técnico
- Criado `.github/workflows/atualizar_driver.yml`.
- Integração com os endpoints JSON oficiais do Chrome for Testing.