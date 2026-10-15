# expiry-report

Scadenzario dei lotti in giacenza, estratto da ARCA Evolution e inviato
periodicamente via email.

Richiede: database ARCA Evolution (SQL Server).

## Problema

Vendite e acquisti devono sapere con anticipo quali prodotti stanno per
scadere, per pianificare promozioni o resi. Le informazioni sono nel gestionale
ma vanno estratte, ordinate e distribuite a più persone.

## Soluzione

`send_expiry_report.py`:

1. interroga `MGDisp('<anno corrente>')` per le giacenze del magazzino `00001`
   con lotti non scaduti, includendo quantità, disponibilità, prezzi B2C/B2B,
   pezzi per cartone, ultimo costo e prima scadenza;
2. salva `Scadenzario.xlsx` (dati grezzi) e `scadenzario.xlsx` (colonne
   rinominate, ordinamento per scadenza, bordi, intestazione bloccata, date
   `gg/mm/aaaa`);
3. invia il file a ogni destinatario tramite SMTP con STARTTLS;
4. elimina via IMAP le copie dalla posta inviata.

L'output di console viene scritto in `output_log.txt`.

## Utilizzo

Impostare nel `.env` alla radice del repository, oltre alla connessione ad
ARCA:

```
SMTP_SERVER=smtp-mail.outlook.com
SMTP_PORT=587
IMAP_SERVER=outlook.office365.com
SMTP_USER=mittente@example.com
SMTP_PASSWORD=...
SCADENZARIO_DESTINATARI=persona1@example.com,persona2@example.com
SCADENZARIO_OUTPUT_DIR=
```

Con la verifica in due passaggi attiva sull'account Microsoft serve una
password per le app. Eseguire con:

```bash
python send_expiry_report.py
```

Lo script è pensato per un'esecuzione pianificata, ad esempio ogni 15 giorni.
