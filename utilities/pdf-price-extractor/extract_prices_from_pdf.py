"""
Estrae da un catalogo PDF le coppie "Codice Interno" / "Prezzo rivendita
suggerito" e le salva in Excel.

Uso: python extract_prices_from_pdf.py <catalogo.pdf>
"""
import sys

import fitz  # PyMuPDF
import re
import pandas as pd

# Funzione per estrarre dati dal PDF
def estrai_dati_pdf(percorso_pdf):
    # Aprire il documento PDF
    documento = fitz.open(percorso_pdf)
    dati_estratti = []

    # Scorrere tutte le pagine
    for pagina_num in range(documento.page_count):
        pagina = documento[pagina_num]
        testo = pagina.get_text()

        # Trova tutte le occorrenze di "Codice Interno" e "Prezzo rivendita suggerito"
        codici_interni = re.findall(r"CODICE\s+INTERNO\s+(\w+)", testo)
        prezzi_suggeriti = re.findall(r"Prezzo\s+rivendita\s+suggerito[:\s€]*([\d,.]+)", testo)

        # Aggiungere i risultati trovati alla lista
        for codice, prezzo in zip(codici_interni, prezzi_suggeriti):
            dati_estratti.append({"Codice Interno": codice, "Prezzo Rivendita Suggerito (€)": float(prezzo.replace(",", "."))})

    # Chiudere il documento PDF
    documento.close()
    return dati_estratti

# Utilizzo della funzione per estrarre i dati
if len(sys.argv) != 2:
    sys.exit("Uso: python extract_prices_from_pdf.py <catalogo.pdf>")
percorso_pdf = sys.argv[1]
dati = estrai_dati_pdf(percorso_pdf)

# Convertire i dati in un DataFrame
df = pd.DataFrame(dati)

# Salvare i dati in un file Excel
nome_file_excel = "Dati_Estratti_Catalogo.xlsx"
df.to_excel(nome_file_excel, index=False)

print(f"Dati salvati con successo nel file {nome_file_excel}")