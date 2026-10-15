# automazioni-ufficio-python

Automazioni Python per attività d'ufficio ricorrenti: report di prodotti e
giacenze per gli agenti, scadenzario dei lotti via email, sconti sui prodotti
in scadenza, cataloghi PowerPoint generati da Excel e alcune utilità su Excel,
immagini e PDF.

Richiede: database ARCA Evolution (SQL Server) per `agent-product-report` e
`expiry-report`; gli altri strumenti lavorano su file locali.

## Contesto

Strumenti sviluppati per le esigenze di un'azienda di distribuzione che usa
ARCA Evolution, pubblicati in forma anonimizzata con l'autorizzazione
dell'azienda.

## Problema

Attività settimanali o mensili come estrarre dati dal gestionale, formattarli
in Excel con immagini e colori, inviarli per email e impaginare cataloghi con
centinaia di prodotti venivano svolte a mano.

## Soluzione

Ogni attività è coperta da uno script autonomo, eseguibile manualmente o da
un'attività pianificata:

```
automazioni-ufficio-python/
├── agent-product-report/          report Excel di prodotti e giacenze per gli agenti
├── expiry-report/                 scadenzario dei lotti estratto da ARCA e inviato via email
├── expiry-discounts/              GUI per gli sconti sui prodotti in scadenza
├── pptx-catalog/                  compilazione di un catalogo PowerPoint da Excel
└── utilities/
    ├── excel-price-list-images/   listino Excel con le immagini scaricate dai link
    ├── image-downloader/          download delle immagini articolo da Excel
    └── pdf-price-extractor/       estrazione di codici e prezzi da un catalogo PDF
```

Ogni cartella ha un README con i dettagli d'uso.

## Tecnologie

- Python 3.10+: pandas, openpyxl, XlsxWriter, pyodbc, SQLAlchemy, python-dotenv
- Email: smtplib/imaplib e Gmail API
- GUI: tkinter, PyQt5
- Documenti e immagini: python-pptx, PyMuPDF, Pillow, BeautifulSoup, requests

## Configurazione

1. Installare il driver ODBC per SQL Server e le dipendenze con
   `pip install -r requirements.txt`.
2. Copiare `.env.example` in `.env` e impostare:

   | Variabile                                 | Usata da             | Significato                                     |
   |-------------------------------------------|----------------------|-------------------------------------------------|
   | `SERVER_DB`, `DB_ARCA`                    | report, scadenzario  | server/istanza e database di ARCA               |
   | `DB_USER`, `DB_PASSWORD`                  | report, scadenzario  | credenziali SQL (vuote: autenticazione Windows) |
   | `GMAIL_SENDER`, `GMAIL_TOKEN_FILE`        | agent-product-report | mittente e token OAuth Gmail                    |
   | `REPORT_OUTPUT_DIR`                       | agent-product-report | cartella dei report (default: cartella script)  |
   | `SMTP_SERVER`, `SMTP_PORT`, `IMAP_SERVER` | expiry-report        | server di posta                                 |
   | `SMTP_USER`, `SMTP_PASSWORD`              | expiry-report        | account di invio                                |
   | `SCADENZARIO_DESTINATARI`                 | expiry-report        | indirizzi separati da virgola                   |
   | `SCADENZARIO_OUTPUT_DIR`                  | expiry-report        | cartella dei file generati                      |

Il file `.env` non è versionato.
