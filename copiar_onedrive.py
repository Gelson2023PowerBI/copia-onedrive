"""
Rotina: Copia ADM_2026.xlsm do OneDrive Pessoal para o OneDrive Profissional
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

# OneDrive Pessoal — origem
PERSONAL_USER_ID   = "ffc5a76f08188416"
PERSONAL_FILE_PATH = "Documents/Gelson/ADM_2026.xlsm"

# OneDrive Profissional — destino
WORK_USER_EMAIL  = "gelson@planilhasmagicas.onmicrosoft.com"
WORK_DEST_FOLDER = "Documents/Gelson_SharePoint"
WORK_FILE_NAME   = "ADM_2026.xlsm"

# ─────────────────────────────────────────────
#  LOGGING
# ─────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s"
)
log = logging.getLogger(__name__)


def get_token() -> str:
    """Obtém token OAuth 2.0 via client_credentials."""
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


def download_arquivo(token: str) -> bytes:
    """Baixa o arquivo do OneDrive Pessoal."""
    url = (
        f"https://graph.microsoft.com/v1.0/users/{PERSONAL_USER_ID}"
        f"/drive/root:/{PERSONAL_FILE_PATH}:/content"
    )
    log.info(f"⬇️  Baixando: {PERSONAL_FILE_PATH}")
    resp = requests.get(url, headers={"Authorization": f"Bearer {token}"})
    resp.raise_for_status()
    log.info(f"✅ Download OK — {len(resp.content):,} bytes")
    return resp.content


def upload_arquivo(token: str, conteudo: bytes) -> None:
    """Envia o arquivo para o OneDrive Profissional, substituindo se existir."""
    dest_path = f"{WORK_DEST_FOLDER}/{WORK_FILE_NAME}"
    url = (
        f"https://graph.microsoft.com/v1.0/users/{WORK_USER_EMAIL}"
        f"/drive/root:/{dest_path}:/content"
    )
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type":  "application/vnd.ms-excel.sheet.macroEnabled.12",
    }
    log.info(f"⬆️  Enviando para: {dest_path}")
    resp = requests.put(url, headers=headers, data=conteudo)
    resp.raise_for_status()
    log.info("✅ Arquivo substituído com sucesso no OneDrive Profissional!")


def main():
    log.info("=" * 55)
    log.info(f"🚀 Início: {datetime.utcnow():%d/%m/%Y %H:%M:%S} UTC")
    try:
        token    = get_token()
        conteudo = download_arquivo(token)
        upload_arquivo(token, conteudo)
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
