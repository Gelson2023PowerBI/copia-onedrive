"""
Rotina: Copia ADM_2026.xlsm do OneDrive Pessoal para o SharePoint Profissional
Execução: GitHub Actions — todo dia às 23:00 (horário de Brasília)
"""

import os
import sys
import requests
import logging
from datetime import datetime

# ─────────────────────────────────────────────
#  CREDENCIAIS — lidas dos Secrets do GitHub
# ─────────────────────────────────────────────
CLIENT_ID     = os.environ["CLIENT_ID"]
CLIENT_SECRET = os.environ["CLIENT_SECRET"]
TENANT_ID     = os.environ["TENANT_ID"]

# OneDrive Pessoal — link de compartilhamento direto
PERSONAL_SHARE_URL = "https://1drv.ms/x/c/ffc5a76f08188416/IQAgvxyXLx__Q6xX4aY2VkPFAVzIC7p7M-1cBMxJTvX0wx0?e=wouYQi"

# SharePoint Profissional — destino
SHAREPOINT_HOSTNAME  = "planilhasmagicas-my.sharepoint.com"
SHAREPOINT_SITE_PATH = "/personal/gelson_planilhasmagicas_onmicrosoft_com"
DEST_FOLDER          = "Documents/Gelson_SharePoint/TranferenciaArquivosClaude"
DEST_FILE_NAME       = "ADM_2026.xlsm"

# ─────────────────────────────────────────────
#  LOGGING
# ─────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s"
)
log = logging.getLogger(__name__)


def get_token() -> str:
    url = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"
    resp = requests.post(url, data={
        "grant_type":    "client_credentials",
        "client_id":     CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "scope":         "https://graph.microsoft.com/.default",
    })
    resp.raise_for_status()
    log.info("✅ Token obtido com sucesso.")
    return resp.json()["access_token"]


def get_site_drive_id(token: str) -> str:
    """Obtém o drive ID do site do SharePoint."""
    headers = {"Authorization": f"Bearer {token}"}

    # Busca o site pelo hostname e caminho
    site_url = (
        f"https://graph.microsoft.com/v1.0/sites/"
        f"{SHAREPOINT_HOSTNAME}:{SHAREPOINT_SITE_PATH}"
    )
    log.info("🔍 Buscando site do SharePoint...")
    resp = requests.get(site_url, headers=headers)
    resp.raise_for_status()
    site_id = resp.json()["id"]
    log.info(f"✅ Site ID: {site_id}")

    # Busca o drive do site
    drives_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives"
    resp = requests.get(drives_url, headers=headers)
    resp.raise_for_status()
    drives = resp.json().get("value", [])

    # Pega o drive "Documents"
    drive_id = None
    for drive in drives:
        log.info(f"   Drive encontrado: {drive['name']} — {drive['id']}")
        if drive.get("name") in ("Documents", "Documentos"):
            drive_id = drive["id"]
            break

    if not drive_id:
        drive_id = drives[0]["id"]

    log.info(f"✅ Drive ID selecionado: {drive_id}")
    return drive_id


def download_arquivo() -> bytes:
    """Baixa o arquivo do OneDrive Pessoal via link de compartilhamento."""
    log.info("⬇️  Baixando arquivo do OneDrive Pessoal...")
    session = requests.Session()
    download_url = PERSONAL_SHARE_URL + "&download=1"
    resp = session.get(download_url, allow_redirects=True)
    resp.raise_for_status()
    log.info(f"✅ Download OK — {len(resp.content):,} bytes")
    return resp.content


def upload_arquivo(token: str, drive_id: str, conteudo: bytes) -> None:
    """Envia o arquivo para o SharePoint, substituindo se existir."""
    dest_path = f"{DEST_FOLDER}/{DEST_FILE_NAME}"
    url = (
        f"https://graph.microsoft.com/v1.0/drives/{drive_id}"
        f"/root:/{dest_path}:/content"
    )
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type":  "application/vnd.ms-excel.sheet.macroEnabled.12",
    }
    log.info(f"⬆️  Enviando para: {dest_path}")
    resp = requests.put(url, headers=headers, data=conteudo)
    resp.raise_for_status()
    log.info("✅ Arquivo substituído com sucesso no SharePoint!")


def main():
    log.info("=" * 55)
    log.info(f"🚀 Início: {datetime.utcnow():%d/%m/%Y %H:%M:%S} UTC")
    try:
        conteudo = download_arquivo()
        token    = get_token()
        drive_id = get_site_drive_id(token)
        upload_arquivo(token, drive_id, conteudo)
        log.info("🎉 Rotina concluída com sucesso!")
    except requests.HTTPError as e:
        log.error(f"❌ Erro HTTP {e.response.status_code}: {e.response.text}")
        sys.exit(1)
    except Exception as e:
        log.error(f"❌ Erro inesperado: {e}")
        sys.exit(1)
    log.info("=" * 55)


if __name__ == "__main__":
    main()
