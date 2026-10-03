# Infraestrutura e Recuperação

## VM

- Sistema: Ubuntu na Oracle Cloud.
- Projeto: `/opt/dash-carga/automacoes/oracle`.
- Dados de entrada: `/opt/dash-carga/entrada`.
- Logs e screenshots: `/opt/dash-carga/logs`.
- Perfis Chromium: `/opt/dash-carga/chromium-profiles`.

## Perfis e CDP

| Serviço | Perfil | Porta CDP |
|---|---|---|
| OTM | `otm` | 9223 |
| WhatsApp | `whatsapp` | 9224 |
| Citrix/SAP | `sap` | 9225 |
| Excel Online | `excel-online` | 9226 |
| Geflux | `geflux` | 9227 |

Uma porta CDP indisponível significa que o Chromium daquele perfil não está
aberto ou morreu. Antes de reabrir, verificar páginas existentes: não iniciar
novo Chromium com o mesmo perfil em uso.

## Recuperação mínima

```bash
cd /opt/dash-carga/automacoes/oracle
dash-carga-healthcheck
systemctl list-timers | grep dash
free -h
curl -s http://127.0.0.1:9223/json/version
```

Verifique também `journalctl` e os logs locais antes de reiniciar serviços. A
VM pode ficar lenta se muitos perfis Chromium e modelos Ollama estiverem abertos.
Modelos locais devem ser descarregados depois de uso (`keep_alive=0`).

## Acesso remoto

RDP deve continuar via túnel SSH, sem expor a porta 3389. Quick Tunnel
TryCloudflare é temporário: a URL muda ao reiniciar o processo, mesmo que a VM
continue ligada. O notificador precisa registrar e enviar a URL renovada.
