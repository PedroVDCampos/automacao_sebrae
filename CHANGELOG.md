## RAE Turbo v1.2.5

Versão focada em estabilidade e diagnóstico inteligente do cliente no RAE.

### Principais mudanças
- O robô agora verifica o semáforo cadastral e a situação da pessoa jurídica antes de prosseguir com o atendimento.
- O Chrome só é reiniciado quando há travamento real do RAE.
- CNPJ não encontrado, cadastro pendente/desatualizado e pessoa jurídica inativa agora são tratados como erros conhecidos, sem reiniciar o navegador.
- Corrigido o serviço `MEI - Parcelamentos de Débitos`.
- Ajustada a seleção de serviços para evitar marcação duplicada quando o RAE exibe opções repetidas.
- Melhorada a abertura e validação do plano orçamentário usando IDs reais da página.
- PDFs com erro continuam sendo enviados para `_ERROS_RAE_TURBO`.

### Arquivos
- `RAE_Turbo.exe`
- `RAE_Turbo.exe.sha256`