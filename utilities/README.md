# utilities

Strumenti indipendenti per operazioni ricorrenti su Excel, immagini e PDF. Non
richiedono il database.

## excel-price-list-images

Problema: i listini inviati ai clienti contengono link alle immagini dei
prodotti, poco leggibili rispetto alle immagini stesse.

Soluzione: `price_list_with_images.py` (GUI tkinter) copia il listino e
sostituisce il link della colonna `Immagini` con l'immagine ridimensionata a
100×100 pixel. Se il link punta a una pagina web, usa il primo tag `<img>`. Le
colonne attese sono definite in `col_indices`.

Utilizzo: `python price_list_with_images.py`, poi selezionare file di origine e
destinazione.

## image-downloader

Problema: le immagini degli articoli vanno scaricate in blocco e rinominate
con il codice articolo, ad esempio per caricarle su un e-commerce.

Soluzione: `download_images_from_excel.py` legge codice articolo (colonna A) e
URL (colonna B) e salva ogni immagine come `<codice>.<estensione>` in
`immagini_articoli/`, ricavando l'estensione dall'URL o dal `Content-Type`.
File, colonne, timeout e pausa tra i download sono configurabili in testa allo
script.

Utilizzo: `python download_images_from_excel.py`.

## pdf-price-extractor

Problema: i prezzi di rivendita suggeriti di un fornitore sono disponibili solo
in un catalogo PDF.

Soluzione: `extract_prices_from_pdf.py` cerca in ogni pagina le coppie
"CODICE INTERNO" / "Prezzo rivendita suggerito" e le salva in
`Dati_Estratti_Catalogo.xlsx`. Per altri cataloghi basta adattare le due
espressioni regolari.

Utilizzo: `python extract_prices_from_pdf.py catalogo.pdf`.
