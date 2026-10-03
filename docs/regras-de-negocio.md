# Regras de Negócio

## DT e cruzamento

- DT é comparada como texto numérico normalizado: remover espaços, pontuação e
  sufixo `.0` antes de cruzar Grade, ZLES002 e Geflux.
- DT presente somente na Grade deve continuar aparecendo no dashboard.
- DT da Grade sem ZLES002: `status_consolidado=PLANEJADO` e
  `status_dados=AGUARDANDO_ZLES002`.

## Prioridade de status

- OTM representa agenda planejada.
- ZLES002 representa dados de expedição/faturamento.
- Geflux representa estado operacional atual.
- `NF EMITIDA` no Geflux equivale a carga expedida.
- `NF SOLICITADA` representa etapa de faturamento, não expedição.
- Status manual não deve ser apagado sem regra explícita e auditável.

## Peso

Peso precisa ser normalizado por unidade. Quilogramas não podem ser exibidos
como toneladas: `600 kg = 0,6 t`; toneladas com separador brasileiro devem ser
tratadas como valor decimal, por exemplo `4.921,586 = 4.921586 t` quando a fonte
já informa toneladas. A unidade da fonte deve sempre prevalecer sobre heurística.

## Segurança operacional

- Automação pode reutilizar sessão salva e clicar em opção SSO conhecida.
- Não pode digitar senha, preencher MFA ou burlar autenticação.
- Popup desconhecido, login ou tela de logoff deve parar a ação e registrar erro.
