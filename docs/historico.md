# Histórico e Aprendizados

## O que foi construído

- Dashboard de agenda e status de cargas em Cloudflare.
- Worker protegido para sincronização da VM.
- Coletor Ubuntu com perfis Chromium persistentes e CDP separados.
- Processamento de Grade OTM, ZLES002, Geflux, MB51 e MB52.
- Excel operacional e publicação compatível com Excel Online.
- Base de integração com WMS sem misturar saldo SAP e estoque operacional.
- Alertas, logs, healthcheck, autorecover e retenção de arquivos.

## Lições importantes

- Sessões gráficas e serviços externos são pontos de falha; health real precisa
  medir resultado, não só processo ativo.
- Uma automação que abre nova aba a cada tentativa consome RAM e mascara erros.
- IA local deve apoiar diagnóstico/revisão, mas ações devem ser limitadas e
  verificáveis.
- O arquivo Excel é uma saída operacional, não a única fonte de verdade.

## Problemas conhecidos a acompanhar

- Sessões corporativas podem expirar e exigem intervenção autorizada.
- Citrix HTML5 pode abrir janelas duplicadas; preservar a janela SAP ativa e
  fechar apenas portais StoreWeb repetidos.
- TryCloudflare não fornece URL permanente.
- Modelos F32 grandes podem esgotar RAM/swap da VM; usar somente sob demanda.
