# Automacao Oracle VM

Esta pasta alimenta o dashboard existente. Ela nao possui frontend e nao substitui
os uploads manuais de `public/cco.html`.

## Fluxo

1. Obtem a grade OTM ou usa um arquivo local autorizado.
2. Filtra LOCAL 1110/1111 e normaliza a grade no formato de `reporte_carga`.
3. Valida a quantidade de DTs antes de enviar qualquer alteracao.
4. Obtem a ZLES002 ou usa um arquivo local.
5. Cruza a ZLES002 somente com as DTs da grade processada.
6. Envia grade, campos operacionais derivados e materiais ao Worker.
7. Registra o status e gera uma imagem PNG da execucao.

O Worker preserva `status`, `hora_chegada`, `n_portaria` e os demais campos
operacionais durante a atualizacao da grade. Campos reconhecidos da ZLES002 podem
atualizar status e dados SAP.

## Preparacao da VM

```bash
cd /opt/dash-carga/automacoes/oracle
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
cp .env.example .env
```

Preencha no minimo:

```text
WORKER_BASE_URL
ORACLE_JOB_SECRET
OTM_GRADE_FILE
SAP_ZLES002_FILE
```

Use o mesmo `ORACLE_JOB_SECRET` configurado como secret no Cloudflare Worker.

## Teste com arquivos locais

```bash
python main.py --grade /caminho/grade.xlsx --zles002 /caminho/zles002.xlsx
```

Arquivos normalizados e a imagem ficam em `WORK_DIR/saida`. Logs ficam em
`WORK_DIR/logs/oracle-automation.log`.

## Agendamento a cada 20 minutos

Execucao direta:

```bash
python scheduler.py
```

Exemplo de servico systemd:

```ini
[Unit]
Description=Dashboard de Carga - Oracle VM
After=network-online.target

[Service]
Type=simple
WorkingDirectory=/opt/dash-carga/automacoes/oracle
EnvironmentFile=/opt/dash-carga/automacoes/oracle/.env
ExecStart=/opt/dash-carga/automacoes/oracle/.venv/bin/python scheduler.py
Restart=always
RestartSec=15

[Install]
WantedBy=multi-user.target
```

Ativacao:

```bash
sudo systemctl daemon-reload
sudo systemctl enable --now dashboard-carga-oracle
sudo journalctl -u dashboard-carga-oracle -f
```

## OTM e SAP

`OTM_MODE=manual` e `SAP_MODE=manual` sao os modos funcionais iniciais. Os
modulos recusam modo automatico enquanto a integracao autorizada nao for
implementada. Nao armazene usuario ou senha no repositorio.

## Protecoes

- grade vazia e rejeitada;
- quantidade abaixo de `MIN_GRADE_DTS` ou acima de `MAX_GRADE_DTS` e rejeitada;
- ZLES002 sem DT correspondente e rejeitada;
- materiais de uma DT so sao substituidos quando existe um conjunto valido;
- apenas uma execucao do scheduler pode rodar por vez;
- falhas sao registradas em `/api/oracle/job-status`.
