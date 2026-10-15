"""
Estrae da ARCA lo scadenzario dei lotti (giacenze con data di scadenza),
lo formatta in Excel e lo invia via email (SMTP Outlook) ai destinatari
configurati; infine elimina la copia dalla posta inviata (IMAP).
"""
import pandas as pd
import xlsxwriter
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import datetime
import sys
import os
import time
import pyodbc
import imaplib
import email
from email.header import decode_header

try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass


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


OUTPUT_DIR = os.getenv('SCADENZARIO_OUTPUT_DIR', os.path.dirname(os.path.abspath(__file__)))
log_file = os.path.join(OUTPUT_DIR, 'output_log.txt')
sys.stdout = open(log_file, 'w')
sys.stderr = open(log_file, 'a')

# INIZIA LO SCRIPT CHE SCARICA LO SCADENZARIO DA ARCA TRAMITE SQL
# Configura la connessione al tuo database SQL Server
conn = pyodbc.connect(get_connection_string())

# Definisci la tua query SQL
query = f"""
SELECT
    MGGiacDisp.Cd_AR,
    AR.Descrizione AS AR_Descrizione,
    ARARMisura.Cd_ARMisura,
    MGGiacDisp.Cd_MG,
    MG.Descrizione AS MG_Descrizione,
    ARARMisura.UMFatt,
    MGGiacDisp.Quantita,
    MGGiacDisp.QuantitaDisp,
    ISNULL(
        (SELECT CONVERT(NUMERIC(18,2), LSArticolo.prezzo)
         FROM LSArticolo
         WHERE CD_AR = AR.CD_AR AND LSArticolo.id_lsrevisione = 1), 0) AS p_b2c,
    ISNULL(
        (SELECT CONVERT(NUMERIC(18,2), LSArticolo.prezzo)
         FROM LSArticolo
         WHERE CD_AR = AR.CD_AR AND LSArticolo.id_lsrevisione = 2), 0) AS p_b2b,
    ISNULL(
        (SELECT CONVERT(INT, UMFatt)
         FROM ARARMisura
         WHERE ARARMISura.CD_AR = AR.Cd_AR AND ARARMisura.cd_ARMisura = 'CF'), 1) AS QxCF,
    ISNULL(
        (SELECT CONVERT(NUMERIC(18,4), ROUND(Costo, 4))
         FROM ARCostoItem
         WHERE TipoCosto = 'U' AND Cd_MGEsercizio = YEAR(GETDATE()) AND CD_AR = AR.CD_AR), 0) AS UltCost,
    ISNULL(
        (SELECT TOP(1) CONVERT(VARCHAR(10), DataScadenza, 105)
         FROM ARLotto
         WHERE DataScadenza >= GETDATE() AND ArLotto.cd_ar = AR.Cd_AR
         ORDER BY DataScadenza), NULL) AS Scadenza
FROM
    MGDisp('{datetime.date.today().year}') MGGiacDisp
    INNER JOIN AR ON MGGiacDisp.Cd_AR = AR.Cd_AR
    INNER JOIN ARARMisura ON AR.Cd_AR = ARARMisura.Cd_AR AND ARARMisura.DefaultMisura = 1
    INNER JOIN MG ON MGGiacDisp.Cd_MG = MG.Cd_MG
WHERE
    1 = 1 AND AR.NoInventario = 0 AND MGGiacDisp.Quantita > 0
    AND AR.CD_AR IN (SELECT DISTINCT cd_ar FROM ARLotto WHERE DataScadenza >= GETDATE())
    AND MGGiacDisp.Cd_MG = '00001'
ORDER BY
    MGGiacDisp.Cd_MG, MGGiacDisp.Cd_AR
"""

# Esegui la query SQL e leggi i dati in un DataFrame di pandas
df = pd.read_sql_query(query, conn)

# Chiudi la connessione al database
conn.close()

# Salva il DataFrame in un file Excel specificando il motore di scrittura 'xlsxwriter'
percorso_file_excel = os.path.join(OUTPUT_DIR, 'Scadenzario.xlsx')
df.to_excel(percorso_file_excel, index=False, engine='xlsxwriter')

# Attesa 5 sec in modo che si crei il file xlsx
time.sleep(5)

# INIZIO SCRIPT CHE CREA LO SCADENZARIO ED INVIA LA MAIL
# Ottieni la data corrente e formattata
data_corrente = datetime.date.today().strftime("%d-%m-%Y")

# Configura le informazioni sull'account email
smtp_server = os.getenv('SMTP_SERVER', 'smtp-mail.outlook.com')
smtp_port = int(os.getenv('SMTP_PORT', '587'))
smtp_username = os.environ['SMTP_USER']
smtp_password = os.environ['SMTP_PASSWORD']

# Definisci una lista di indirizzi email destinatari
# (variabile SCADENZARIO_DESTINATARI, indirizzi separati da virgola)
to_email_list = [e.strip() for e in os.environ['SCADENZARIO_DESTINATARI'].split(',') if e.strip()]

# Configura le informazioni sull'email
from_email = smtp_username
subject = 'Scadenzario del ' + data_corrente
message = f'Ciao,\n\nEcco lo scadenzario del {data_corrente}, ogni 15 giorni riceverai lo scadenzario aggiornato.\n\nSaluti'

for to_email in to_email_list:
    msg = MIMEMultipart()
    msg['From'] = from_email
    msg['To'] = to_email
    msg['Subject'] = subject

    # Aggiungi il corpo del messaggio
    msg.attach(MIMEText(message, 'plain'))

    # Creazione scadenzario
    # Leggi il file Excel
    df = pd.read_excel(percorso_file_excel)
    # Seleziona solo le colonne desiderate
    colonne_desiderate = ['Cd_AR', 'AR_Descrizione', 'Cd_ARMisura', 'Quantita', 'QuantitaDisp', 'p_b2c', 'p_b2b', 'QxCF', 'UltCost', 'Scadenza']
    df = df[colonne_desiderate]
    # Rinomina le colonne
    df.columns = ['Codice Articolo', 'Nome', 'UM', 'Quantita', 'QuantitaDisp', 'Prezzo B2C', 'Prezzo B2B', 'QxCF', 'UltCost', 'Scadenza']
    # Converte la colonna 'Scadenza' in formato datetime
    df['Scadenza'] = pd.to_datetime(df['Scadenza'], format="%d-%m-%Y")
    # Ordina il DataFrame per la colonna 'Scadenza' in ordine crescente (data più recente in cima)
    df = df.sort_values(by='Scadenza', ascending=True)
    # Formatta la colonna 'Scadenza' per mostrare solo la data senza l'orario
    df['Scadenza'] = df['Scadenza'].dt.strftime('%d/%m/%Y')
    # Crea un oggetto ExcelWriter utilizzando il motore xlsxwriter
    scadenzario_filename = os.path.join(OUTPUT_DIR, 'scadenzario.xlsx')
    writer = pd.ExcelWriter(scadenzario_filename, engine='xlsxwriter')
    # Aggiungi bordi neri spessi alle celle con dati
    df = df.style.applymap(lambda x: "border: 4px solid black")
    # Scrivi il DataFrame nel foglio Excel
    df.to_excel(writer, index=False, sheet_name='Sheet1')
    # Congela la riga 1
    worksheet = writer.sheets['Sheet1']
    worksheet.freeze_panes(1, 0)
    # Imposta la formattazione della colonna Scadenza come data in formato dd/mm/yyyy
    date_format = writer.book.add_format({'num_format': 'dd/mm/yyyy'})
    worksheet.set_column('J:J', None, date_format)
    # Salva il file Excel
    writer._save()

    # Allega lo scadenzario
    with open(scadenzario_filename, 'rb') as attachment:
        part = MIMEApplication(attachment.read())
        part.add_header('Content-Disposition', 'attachment', filename=os.path.basename(scadenzario_filename))
        msg.attach(part)

    # Crea una connessione SMTP sicura e invia l'email
    try:
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()
        server.login(smtp_username, smtp_password)
        server.sendmail(from_email, to_email, msg.as_string())
        server.quit()
        print(f'Email inviata con successo alla mail {to_email}')
    except smtplib.SMTPException as e:
        print(f'Errore SMTP: {str(e)}')
    except Exception as e:
        print(f'Errore durante l\'invio dell\'email: {str(e)}')

# ELIMINAZIONE DELLA MAIL INVIATA
# Configura le informazioni sull'account email
imap_server = os.getenv('IMAP_SERVER', 'outlook.office365.com')
imap_port = 993
email_address = smtp_username
email_password = smtp_password

# Connetti al server IMAP
mail = imaplib.IMAP4_SSL(imap_server)
mail.login(email_address, email_password)

# Seleziona la cartella della posta inviata (Sent Mail)
mail.select('Sent')

# Cerca l'email che vuoi eliminare (puoi utilizzare diversi criteri di ricerca)
# In questo esempio, stiamo cercando un'email con un certo oggetto
email_subject_to_delete = 'Scadenzario del ' + data_corrente  # Modifica il soggetto come necessario
# Cerca gli email con il soggetto specificato da eliminare
status, email_ids = mail.search(None, f'SUBJECT "{email_subject_to_delete}"')

# Estrai gli ID delle email trovate
email_id_list = email_ids[0].split()

# Elimina le email trovate
for email_id in email_id_list:
    mail.store(email_id, '+FLAGS', '(\\Deleted)')

# Esegui l'eliminazione effettiva
mail.expunge()
print("La mail inviata è stata eliminata!")

# Chiudi la connessione al server IMAP
mail.logout()