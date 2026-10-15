# expiry-discounts

Preparazione dei file di sconto per i prodotti in scadenza.

## Problema

I prodotti vicini alla scadenza vengono messi in promozione con sconti
crescenti in base al mese di scadenza. Calcolare i prezzi scontati B2C e B2B
per ogni articolo e preparare i file di import per il sito era un lavoro
manuale.

## Soluzione

`expiry_discounts_gui.py` (GUI PyQt5) parte dallo scadenzario esportato da ARCA,
con le colonne `Cd_AR`, `AR_Descrizione`, `Cd_ARMisura`, `Quantita`,
`QuantitaDisp`, `p_b2c`, `p_b2b`, `QxCF`, `UltCost`, `Scadenza` (`gg-mm-aaaa`): lo
stesso formato del file grezzo di `expiry-report` o del modello di stampa
giacenze di ARCA (Magazzini → Stampe → Giacenze).

| Funzione | Risultato |
|----------|-----------|
| Crea File Scadenzario | `scadenzario.xlsx` formattato e ordinato per scadenza |
| Crea CSV sconti B2C e B2B | per i primi tre mesi di scadenza chiede la percentuale di sconto di ciascun mese e genera `Sconti_B2C.csv` (prezzo al pezzo) e `Sconti_B2B.csv` (prezzo a cartone) |
| Crea File Rimuovi Sconti | `rimuovi_sconti.csv` con prezzo in offerta vuoto, per chiudere le promozioni |

I CSV hanno le colonne `COD`, `Nome`, `Prezzo in offerta`, nel formato di
import del sito e-commerce.

## Utilizzo

```bash
python expiry_discounts_gui.py
```

Selezionare il file Excel e usare il pulsante corrispondente all'operazione.
