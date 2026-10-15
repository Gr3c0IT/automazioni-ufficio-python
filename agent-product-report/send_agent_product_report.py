#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Report Prodotti - Script unificato multi-cliente
Uso: python send_agent_product_report.py [CLIENTE_A|CLIENTE_D]
     (se non specificato, esegue entrambi i clienti)
"""

import argparse
import pandas as pd
import urllib.parse
from sqlalchemy import create_engine
import os
from datetime import datetime
import base64
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from googleapiclient.discovery import build
import warnings
import logging
from logging.handlers import RotatingFileHandler
import requests
from PIL import Image as PILImage
import io
import hashlib
import shutil

warnings.filterwarnings('ignore')

try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_DIR = os.getenv('REPORT_OUTPUT_DIR', SCRIPT_DIR)

# ─────────────────────────────────────────────
# CONFIGURAZIONE CLIENTI
# ─────────────────────────────────────────────
CLIENTS = {
    'CLIENTE_A': {
        'label':            'CLIENTE_A',
        'log_suffix':       '',                         # es. report_prodotti_YYYYMMDD.log
        'id_ls_revisione':  10,                         # Id_LSRevisione del listino del cliente
        'output_file':      os.path.join(OUTPUT_DIR, 'report_prodotti_CLIENTE_A.xlsx'),
        'email_subject':    'Report Prodotti CLIENTE_GDO',
        'recipients': [
            'collega_a@azienda.example.com',
            'collega_b@azienda.example.com',
            'agente_1@example.com',
            'agente_2@example.com',
            'agente_3@example.com',
            'agente_4@example.com',
            'agente_5@example.com',
            'collega_c@azienda.example.com',
        ],
    },
    'CLIENTE_D': {
        'label':            'CLIENTE_D',
        'log_suffix':       '_CLIENTE_D',                 # es. report_prodotti_CLIENTE_D_YYYYMMDD.log
        'id_ls_revisione':  20,                         # Id_LSRevisione del listino del cliente
        'output_file':      os.path.join(OUTPUT_DIR, 'report_prodotti_CLIENTE_D.xlsx'),
        'email_subject':    'Report Prodotti CLIENTE_GDO',
        'recipients': [
            'collega_a@azienda.example.com',
            'collega_b@azienda.example.com',
            'agente_6@agenzia-a.example.com',
            'agente_7@example.com',
            'collega_c@azienda.example.com',
        ],
    },
}

# ─────────────────────────────────────────────
# DATABASE
# ─────────────────────────────────────────────
def get_connection_string():
    """Stringa di connessione ODBC ad ARCA, letta dalle variabili d'ambiente (.env)."""
    conn_str = (
        "Driver={SQL Server};"
        f"Server={os.environ['SERVER_DB']};"
        f"Database={os.environ['DB_ARCA']};"
    )
    if os.getenv("DB_USER"):
        return conn_str + f"UID={os.environ['DB_USER']};PWD={os.getenv('DB_PASSWORD', '')};"
    return conn_str + "Trusted_Connection=yes;"

SENDER_EMAIL = os.getenv('GMAIL_SENDER', 'collega_a@azienda.example.com')
LOGO_FILE    = os.path.join(SCRIPT_DIR, 'logo.png')
PICKLE_FILE  = os.getenv('GMAIL_TOKEN_FILE', os.path.join(SCRIPT_DIR, 'token.pickle'))


# ─────────────────────────────────────────────
# LOGGER
# ─────────────────────────────────────────────
def setup_logger(log_suffix: str) -> logging.Logger:
    """Configura un logger dedicato per il cliente corrente."""
    log_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'logs')
    os.makedirs(log_dir, exist_ok=True)

    log_file = os.path.join(
        log_dir,
        f'report_prodotti{log_suffix}_{datetime.now().strftime("%Y%m%d")}.log'
    )

    logger = logging.getLogger(f'ReportProdotti{log_suffix}')
    logger.setLevel(logging.INFO)

    # Evita di aggiungere handler duplicati se il logger esiste già
    if logger.handlers:
        logger.handlers.clear()

    fmt = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')

    fh = RotatingFileHandler(log_file, maxBytes=5*1024*1024, backupCount=5, encoding='utf-8')
    fh.setLevel(logging.INFO)
    fh.setFormatter(fmt)

    ch = logging.StreamHandler()
    ch.setLevel(logging.INFO)
    ch.setFormatter(fmt)

    logger.addHandler(fh)
    logger.addHandler(ch)
    logger.propagate = False

    return logger


# ─────────────────────────────────────────────
# IMMAGINI
# ─────────────────────────────────────────────
def download_and_resize_image(
    image_url: str,
    logger: logging.Logger,
    max_width: int = 80,
    max_height: int = 60,
) -> str | None:
    """
    Scarica un'immagine da un URL e la ridimensiona per l'inserimento in Excel.
    Restituisce il percorso locale del file oppure None in caso di errore.
    """
    if not image_url or pd.isna(image_url) or str(image_url).strip() == '':
        return None

    try:
        temp_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'temp_images')
        os.makedirs(temp_dir, exist_ok=True)

        url_hash     = hashlib.md5(str(image_url).encode()).hexdigest()
        temp_filename = os.path.join(temp_dir, f'img_{url_hash}.png')

        if os.path.exists(temp_filename):
            logger.debug(f"Immagine già in cache: {temp_filename}")
            return temp_filename

        headers  = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'}
        response = requests.get(str(image_url).strip(), headers=headers, timeout=10)
        response.raise_for_status()

        image = PILImage.open(io.BytesIO(response.content))

        # Gestione trasparenza
        if image.mode in ('RGBA', 'LA', 'P'):
            background = PILImage.new('RGB', image.size, (255, 255, 255))
            if image.mode == 'P':
                image = image.convert('RGBA')
            background.paste(image, mask=image.split()[-1] if image.mode == 'RGBA' else None)
            image = background
        elif image.mode != 'RGB':
            image = image.convert('RGB')

        image.thumbnail((max_width, max_height), PILImage.Resampling.LANCZOS)
        image.save(temp_filename, 'PNG', optimize=True)

        logger.debug(f"Immagine scaricata: {image_url} -> {temp_filename}")
        return temp_filename

    except requests.exceptions.RequestException as e:
        logger.warning(f"Errore download immagine {image_url}: {e}")
    except Exception as e:
        logger.warning(f"Errore elaborazione immagine {image_url}: {e}")
    return None


def cleanup_temp_images(logger: logging.Logger) -> None:
    """Elimina la directory delle immagini temporanee."""
    temp_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'temp_images')
    if os.path.exists(temp_dir):
        try:
            shutil.rmtree(temp_dir)
            logger.info("Immagini temporanee eliminate.")
        except Exception as e:
            logger.warning(f"Errore eliminazione immagini temporanee: {e}")


# ─────────────────────────────────────────────
# GMAIL
# ─────────────────────────────────────────────
def gmail_authenticate(logger: logging.Logger):
    """Carica le credenziali dal file pickle e restituisce il servizio Gmail."""
    import pickle
    try:
        with open(PICKLE_FILE, 'rb') as fh:
            creds = pickle.load(fh)
        logger.info(f"Autenticazione Gmail completata (tipo: {type(creds).__name__})")
        if hasattr(creds, 'scopes'):
            logger.info(f"Scope: {', '.join(creds.scopes)}")
        service = build('gmail', 'v1', credentials=creds)
        logger.info("Servizio Gmail creato con successo.")
        return service
    except Exception as e:
        logger.error(f"Errore autenticazione Gmail: {e}")
        logger.error(f"Verifica che '{PICKLE_FILE}' sia nella stessa cartella dello script.")
        return None


def build_email(
    sender: str,
    recipients: list[str],
    subject: str,
    body: str,
    attachment_path: str,
    logger: logging.Logger,
) -> dict:
    """Costruisce il messaggio MIME pronto per l'API Gmail."""
    msg = MIMEMultipart()
    msg['from']    = sender
    msg['to']      = ', '.join(recipients)
    msg['subject'] = subject

    logger.info(f"📧 Destinatari: {msg['to']}")
    msg.attach(MIMEText(body))

    if os.path.exists(attachment_path):
        with open(attachment_path, 'rb') as f:
            att = MIMEApplication(f.read(), _subtype='xlsx')
            att.add_header('Content-Disposition', 'attachment', filename=os.path.basename(attachment_path))
            msg.attach(att)
        logger.info(f"📎 Allegato: {attachment_path}")
    else:
        logger.warning(f"⚠️ File allegato non trovato: {attachment_path}")

    return {'raw': base64.urlsafe_b64encode(msg.as_bytes()).decode('utf-8')}


def send_email(service, message: dict, logger: logging.Logger) -> bool:
    """Invia il messaggio tramite Gmail API. Restituisce True se l'invio ha avuto successo."""
    try:
        logger.info("Invio email in corso...")
        sent = service.users().messages().send(userId='me', body=message).execute()
        logger.info(f"Email inviata! ID Gmail: {sent.get('id')}")
        logger.info("Verifica dettagli invio saltata: il token Gmail ha solo permessi di invio.")
        return True
    except Exception as e:
        logger.error(f"Errore invio email: {e} ({type(e).__name__})")
        return False


# ─────────────────────────────────────────────
# CORE: elaborazione per un singolo cliente
# ─────────────────────────────────────────────
def process_client(client_key: str) -> None:
    cfg    = CLIENTS[client_key]
    logger = setup_logger(cfg['log_suffix'])

    logger.info("=" * 50)
    logger.info(f"Avvio elaborazione report prodotti — {cfg['label']}")

    # ── Connessione DB ──
    try:
        params = urllib.parse.quote_plus(get_connection_string())
        engine = create_engine(f"mssql+pyodbc:///?odbc_connect={params}")
        logger.info("Connessione al database riuscita.")
    except Exception as e:
        logger.error(f"Errore connessione database: {e}")
        return

    # ── Query ──
    query = f"""
    SELECT
        ls.Cd_AR                        AS "Codice Prodotto",
        arm.Descrizione                 AS "Marca",
        arg.Descrizione                 AS "Categoria",
        ar.Descrizione                  AS "Prodotto",
        ara.Alias                       AS "EAN prodotto",
        CAST(ar.PesoNetto   AS INT)     AS "Grammi/ml",
        CAST(arm2.UMFatt    AS INT)     AS "PZ X CF",
        ar.xWcScadenza                  AS "Scadenza",
        CAST(mg.Quantita    AS INT)     AS "Quantita",
        ari.Note_ARImg                  AS "URL_Immagine"
    FROM        LSArticolo  ls
    LEFT JOIN   AR          ar   ON ls.Cd_AR        = ar.Cd_AR
    LEFT JOIN   ARMarca     arm  ON ar.Cd_ARMarca   = arm.Cd_ARMarca
    LEFT JOIN   ARGruppo1   arg  ON ar.Cd_ARGruppo1 = arg.Cd_ARGruppo1
    LEFT JOIN   ARAlias     ara  ON ar.Cd_AR = ara.Cd_AR
                                AND ara.Cd_ARMisura = 'PZ'
                                AND ara.TipoAlias   = 'E'
    LEFT JOIN   ARARMisura  arm2 ON ar.Cd_AR = arm2.Cd_AR
                                AND arm2.Cd_ARMisura = 'CF'
    LEFT JOIN   MGGiac('2026') mg ON ar.Cd_AR = mg.Cd_AR
                                AND mg.Cd_MG = '00001'
    LEFT JOIN   ARImg       ari  ON ar.Cd_AR = ari.Cd_AR
                                AND ari.TipoImg LIKE '%B%'
    WHERE  ls.Id_LSRevisione = {cfg['id_ls_revisione']}
    ORDER BY arm.Descrizione
    """

    try:
        df = pd.read_sql(query, engine)
        logger.info(f"Query eseguita con successo: {len(df)} righe recuperate.")
    except Exception as e:
        logger.error(f"Errore esecuzione query: {e}")
        return

    # ── Elaborazione dati ──
    try:
        df["Scadenza"] = pd.to_datetime(df["Scadenza"], errors='coerce').dt.strftime('%d/%m/%Y')

        quantita      = df["Quantita"].copy()
        url_immagini  = df["URL_Immagine"].copy()

        df = df.drop(columns=["Quantita", "URL_Immagine"])
        df["Ordine"] = ""
        df.insert(0, "Immagine", "")

        logger.info("Elaborazione dati completata.")
    except Exception as e:
        logger.error(f"Errore elaborazione dati: {e}")
        return

    # ── Download immagini ──
    logger.info("Download immagini in corso...")
    percorsi_immagini = []
    for i, url in enumerate(url_immagini):
        if i % 10 == 0:
            logger.info(f"  Immagine {i+1}/{len(url_immagini)}...")
        percorsi_immagini.append(download_and_resize_image(url, logger))

    ok_count = sum(1 for p in percorsi_immagini if p is not None)
    logger.info(f"Download completato: {ok_count}/{len(url_immagini)} immagini scaricate.")

    # ── Creazione Excel ──
    try:
        from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
        from openpyxl.worksheet.table import Table, TableStyleInfo
        from openpyxl.drawing.image import Image as XLImage

        file_name = cfg['output_file']
        writer    = pd.ExcelWriter(file_name, engine='openpyxl')
        df.to_excel(writer, index=False, sheet_name='Report', startrow=3)

        wb = writer.book
        ws = writer.sheets['Report']

        HEADER_ROW = 4

        # Intestazioni grassetto + centrato
        for cell in ws[HEADER_ROW]:
            cell.font      = Font(bold=True)
            cell.alignment = Alignment(horizontal='center')

        # Altezza righe dati
        for row_num in range(HEADER_ROW + 1, HEADER_ROW + 1 + len(df)):
            ws.row_dimensions[row_num].height = 45

        # Larghezza colonna immagini
        ws.column_dimensions['A'].width = 11

        # Inserimento immagini
        logger.info("Inserimento immagini nell'Excel...")
        for i, img_path in enumerate(percorsi_immagini):
            if img_path and os.path.exists(img_path):
                try:
                    img = XLImage(img_path)
                    img.width  = min(img.width,  80)
                    img.height = min(img.height, 60)
                    ws.add_image(img, f'A{HEADER_ROW + 1 + i}')
                except Exception as e:
                    logger.warning(f"Errore inserimento immagine riga {i+1}: {e}")

        # Formattazione condizionale (quantità < 150 → rosso)
        red_fill    = PatternFill(start_color='FF0000', end_color='FF0000', fill_type='solid')
        colored     = 0
        prod_col_idx = df.columns.get_loc("Prodotto") + 1

        for i, q in enumerate(quantita, start=HEADER_ROW + 1):
            if pd.notna(q) and q < 150:
                for cell in ws[i]:
                    cell.fill = red_fill
                prod_cell = ws.cell(row=i, column=prod_col_idx)
                if prod_cell.value and not str(prod_cell.value).startswith("! "):
                    prod_cell.value = f"! {prod_cell.value}"
                colored += 1

        logger.info(f"Formattazione condizionale: {colored} righe evidenziate in rosso.")

        # Logo
        if os.path.exists(LOGO_FILE):
            logo        = XLImage(LOGO_FILE)
            logo.width  = 246
            logo.height = 139
            ws.add_image(logo, 'A1')
        ws.row_dimensions[1].height = 104

        # Intestazione ordine (righe 2-3)
        thick = Side(style='thick', color='000000')
        border_style = Border(top=thick, left=thick, right=thick, bottom=thick)

        ws['A2'] = 'Codice rifatturazione CLIENTE_GDO'
        ws.merge_cells('C2:D2')
        for col in ['C2', 'D2']:
            ws[col].border = border_style

        ws['E2'] = 'Ragione Sociale'
        ws.merge_cells('F2:G2')
        for col in ['F2', 'G2']:
            ws[col].border = border_style

        ws['H2'] = 'Data Ordine'
        ws['I2'].border = border_style

        ws['E3'] = 'Indirizzo di Consegna'
        ws['F3'].border = border_style
        ws['H3'] = 'Data di Consegna'
        ws['I3'].border = border_style

        # Tabella
        last_row = HEADER_ROW + len(df)
        from openpyxl.utils import get_column_letter
        last_col = get_column_letter(len(df.columns))
        tab = Table(displayName="ReportTable", ref=f"A{HEADER_ROW}:{last_col}{last_row}")
        tab.tableStyleInfo = TableStyleInfo(
            name="TableStyleMedium9",
            showFirstColumn=False, showLastColumn=False,
            showRowStripes=True,  showColumnStripes=False,
        )
        ws.add_table(tab)

        # Larghezza colonne automatica (esclusa A)
        from openpyxl.utils import get_column_letter

        for i, column in enumerate(df.columns, start=1):
            if i == 1:
                continue  # salta colonna Immagine

            values = df[column].fillna("").astype(str)
            max_len = values.map(lambda x: len(x) if isinstance(x, str) else len(str(x))).max()

            if pd.isna(max_len):
                max_len = 0

            col_width = max(len(str(column)), int(max_len))
            ws.column_dimensions[get_column_letter(i)].width = min(col_width + 2, 50)

        # Larghezze fisse per colonne Ragione Sociale e Data
        ws.column_dimensions['F'].width = 140 / 7
        ws.column_dimensions['I'].width = 114 / 7

        writer.close()
        logger.info(f"File Excel '{file_name}' creato con successo.")

    except Exception as e:
        logger.error(f"Errore creazione file Excel: {e}")
        cleanup_temp_images(logger)
        return

    # ── Invio email ──
    # try:
    #     service = gmail_authenticate(logger)
    #     if service is None:
    #         logger.error("Invio email saltato: autenticazione Gmail fallita.")
    #     else:
    #         subject = f"{cfg['email_subject']} - {datetime.now().strftime('%d/%m/%Y')}"
    #         body    = (
    #             "Buongiorno,\n\n"
    #             "In allegato il report aggiornato dei prodotti generato automaticamente.\n\n"
    #             "Attenzione: le righe in rosso indicano prodotti con giacenza inferiore a 150 unità.\n"
    #             "Questi prodotti sono anche segnalati con il simbolo ! all'inizio della descrizione.\n\n"
    #             "N.B.: PER VISUALIZZARE LE RIGHE IN ROSSO E LE IMMAGINI, SCARICA IL FILE E APRILO "
    #             "CON MICROSOFT EXCEL, GOOGLE FOGLI O SOFTWARE EQUIVALENTE.\n\n"
    #             "Cordiali saluti,\nCOLLEGA_A"
    #         )
    #         message = build_email(SENDER_EMAIL, cfg['recipients'], subject, body, file_name, logger)
    #         if send_email(service, message, logger):
    #             logger.info(f"✅ Email inviata a {len(cfg['recipients'])} destinatari.")
    #         else:
    #             logger.error("❌ Invio email fallito.")
    # except Exception as e:
    #     logger.error(f"Errore generale invio email: {e}")
    # finally:
    
    cleanup_temp_images(logger)

    logger.info(f"Elaborazione {cfg['label']} completata con successo!")
    logger.info("=" * 50)


# ─────────────────────────────────────────────
# ENTRY POINT
# ─────────────────────────────────────────────
def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Genera e invia il report prodotti per uno o più clienti GDO."
    )
    parser.add_argument(
        'cliente',
        nargs='?',
        choices=list(CLIENTS.keys()),
        default=None,
        help=(
            "Cliente da elaborare. "
            f"Valori validi: {', '.join(CLIENTS.keys())}. "
            "Se omesso, elabora tutti i clienti."
        ),
    )
    return parser.parse_args()


if __name__ == "__main__":
    args = parse_args()

    if args.cliente:
        # Elabora solo il cliente specificato
        process_client(args.cliente)
    else:
        # Elabora tutti i clienti in sequenza
        for key in CLIENTS:
            process_client(key)