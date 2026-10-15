# pptx-catalog

Compilazione automatica di un catalogo PowerPoint a partire da un file Excel.

## Problema

Il catalogo prodotti è un PowerPoint con centinaia di schede, ciascuna con
testi, immagine e loghi. Aggiornarlo a mano a ogni edizione richiedeva giorni
di lavoro e introduceva errori di trascrizione.

## Soluzione

Due script:

| File | Ruolo |
|------|-------|
| `catalog_gui.py` | interfaccia tkinter per scegliere Excel e modello PowerPoint e definire per ogni colonna font, dimensione, colore, grassetto, corsivo e allineamento |
| `generate_pptx_catalog.py` | elaborazione: sostituisce i segnaposto, scarica le immagini, inserisce i loghi e salva `catalogo_pronto.pptx` |

Ogni riga dell'Excel corrisponde a un prodotto e i nomi delle colonne sono i
segnaposto. Alcune colonne hanno un trattamento specifico:

- `Immagini`: URL dell'immagine del prodotto;
- `Peso` e `Categoria`: unità `ml` per la categoria "bevande", `g` per le altre;
- `ean_pz`, `ean_cf`, `pz_cf`, `cf_ly`, `cf_plt`, `ly_plt`: formattati come interi;
- `Prezzo_Rivendita_Suggerito`: formattato come valuta;
- `TOP_SELLER`, `NOVITA`, `TREND`, `GLUTEN_FREE`, `PROTEICO`, `VEGANO`, `HALAL`:
  se veri inseriscono il logo corrispondente.

Il file prodotto da `arca-evolution-toolkit/data-export/product-catalog-export`
ha già questa struttura.

Requisiti del modello PowerPoint:

- ogni prodotto è un gruppo di forme; lo script associa un gruppo a ogni riga;
- ogni forma di testo contiene il nome di una sola colonna (es. `NOME_PRODOTTO`);
- la forma con testo `Immagine` viene sostituita dalla foto del prodotto;
- le forme `LOGO1` … `LOGO8` ricevono i loghi in ordine e quelle inutilizzate
  vengono svuotate;
- le posizioni dei gruppi vengono salvate e ripristinate dopo l'elaborazione.

I loghi sono PNG nella cartella `loghi/` accanto agli script, uno per flag
(`TOP_SELLER.png`, `NOVITA.png`, ...), e non sono inclusi nel repository.

## Utilizzo

```bash
python catalog_gui.py
```

Da riga di comando:
`python generate_pptx_catalog.py <excel> <pptx> "<dizionario stili>"`.
