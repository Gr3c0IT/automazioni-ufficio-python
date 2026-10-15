# agent-product-report

Report Excel dei prodotti a listino per gli agenti di vendita, con immagini e
segnalazione delle giacenze basse.

Richiede: database ARCA Evolution (SQL Server).

## Problema

Gli agenti che seguono i clienti GDO hanno bisogno di un elenco aggiornato dei
prodotti a listino per ciascun cliente, con immagine, dati logistici e
disponibilità, da usare anche come modulo d'ordine.

## Soluzione

`send_agent_product_report.py` genera un Excel per ogni cliente configurato:

1. legge da ARCA gli articoli del listino del cliente (`LSArticolo` con
   l'`Id_LSRevisione` configurato), con marca, categoria, EAN, grammatura,
   pezzi per cartone, scadenza, giacenza del magazzino `00001` e URL
   dell'immagine (`ARImg`);
2. scarica e ridimensiona le immagini, con cache in `temp_images/`;
3. compone l'Excel con openpyxl: logo, intestazione per i dati d'ordine, tabella
   formattata e immagini; le righe con giacenza inferiore a 150 unità sono
   evidenziate in rosso e marcate con `!`;
4. può inviare il file ai destinatari del cliente tramite Gmail API; il blocco
   di invio è attualmente commentato nel codice.

## Utilizzo

1. Configurare il `.env` alla radice del repository: connessione ad ARCA,
   `GMAIL_SENDER`, `GMAIL_TOKEN_FILE`, `REPORT_OUTPUT_DIR`.
2. Nel dizionario `CLIENTS` impostare per ogni cliente etichetta,
   `id_ls_revisione`, file di output, oggetto e destinatari.
3. Facoltativo: `logo.png` nella cartella dello script; `token.pickle` con le
   credenziali OAuth Gmail (scope `gmail.send`) per l'invio.
4. Eseguire:
   ```bash
   python send_agent_product_report.py              # tutti i clienti
   python send_agent_product_report.py CLIENTE_A    # un solo cliente
   ```

`run_report.bat` esegue lo script dalla propria cartella, ad esempio da
Utilità di pianificazione di Windows, e accoda l'output in
`logs\run_report.log`. I log applicativi ruotano nella cartella `logs/`.
