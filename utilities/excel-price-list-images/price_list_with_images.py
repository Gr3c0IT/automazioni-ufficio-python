import openpyxl
import requests
from io import BytesIO
from PIL import Image
from bs4 import BeautifulSoup
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
from tqdm import tqdm
from tkinter import ttk

# Indici delle colonne
col_indices = {
    "Codice Cliente": 0,
    "Codice Articolo": 1,
    "Marca": 2,
    "Descrizione": 3,
    "EAN": 4,
    "PZ": 5,
    "Prezzo Listino PZ": 6,
    "PZxCF": 7,
    "Listino CF": 8,
    "PZxLayer": 9,
    "Listino LY": 10,
    "PZxPLT": 11,
    "Listino PL": 12,
    "Immagini": 13
    }


def process_excel(col_indices):
    # Ottieni i percorsi dai campi di input
    xlsx_path = input_path.get()
    modified_xlsx_path = output_path.get()

    # Apri il file XLSX originale
    workbook = openpyxl.load_workbook(xlsx_path)
    sheet = workbook.active

    # Crea un nuovo foglio per il file XLSX modificato
    new_workbook = openpyxl.Workbook()
    new_sheet = new_workbook.active

    # Conta il numero totale di righe per la barra di caricamento
    total_rows = sheet.max_row

    # Creazione della barra di caricamento
    progress_bar = ttk.Progressbar(temp_root, length=200, mode='determinate', maximum=total_rows)
    progress_bar.pack(pady=5)

    # Copia i dati dalla vecchia sheet alla nuova sheet, includendo le immagini
    for row_num, row in enumerate(tqdm(sheet.iter_rows(values_only=True), total=sheet.max_row), start=1):
        # Copia i dati tranne le immagini
        for col_name, col_index in col_indices.items():
            cell_value = row[col_index]
            new_cell = new_sheet.cell(row=row_num, column=col_index + 1, value=cell_value)
            new_cell.alignment = openpyxl.styles.Alignment(horizontal="center", vertical="center")
            new_cell.border = openpyxl.styles.Border(top=openpyxl.styles.Side(style='thick', color='000000'),
                                                     right=openpyxl.styles.Side(style='thick', color='000000'),
                                                     bottom=openpyxl.styles.Side(style='thick', color='000000'),
                                                     left=openpyxl.styles.Side(style='thick', color='000000'))

        # Aggiungi i nomi delle colonne nella prima riga
        if row_num == 1:
            for col_name, col_index in col_indices.items():
                new_sheet.cell(row=row_num, column=col_index + 1, value=col_name)
                new_sheet.column_dimensions[openpyxl.utils.get_column_letter(col_index + 1)].width = len(col_name) * 1.5  # Regola la   larghezza della colonna

        # Aggiorna la barra di caricamento
        progress_bar['value'] = row_num
        temp_root.update_idletasks()  # Aggiorna l'interfaccia grafica

        link = row[col_indices["Immagini"]]  # Ottieni il link dalla colonna "Immagini"

        # Aggiungi un controllo per evitare link non validi
        if not link or not link.startswith("http"):
            continue

        # Verifica se il link contiene direttamente un'immagine o un link al sito
        if link.endswith((".jpg", ".png", ".jpeg")):
            direct_image_url = link
        else:
            # Accedi alla pagina del link e cerca il link diretto all'immagine
            response = requests.get(link)
            if response.status_code == 200:
                soup = BeautifulSoup(response.content, "html.parser")
                img_tags = soup.find_all("img")
                if img_tags:
                    direct_image_url = img_tags[0]["src"]
                else:
                    direct_image_url = None
            else:
                direct_image_url = None

        # Se è stato trovato un link diretto all'immagine, scarica l'immagine
        if direct_image_url:
            response = requests.get(direct_image_url)
            if response.status_code == 200:
                image_data = BytesIO(response.content)
                img = Image.open(image_data)

                # Ridimensiona l'immagine a 100x100 pixel
                img.thumbnail((100, 100))

                # Inserisci l'immagine sotto la colonna "Immagini" (colonna N)
                img_cell = new_sheet.cell(row=row_num, column=col_indices["Immagini"] + 1)
                img_cell.value = ""
                img_obj = openpyxl.drawing.image.Image(image_data)
                img_obj.width = 100
                img_obj.height = 100
                new_sheet.column_dimensions[openpyxl.utils.get_column_letter(col_indices["Immagini"] + 1)].width = 100 / 7.5
                new_sheet.row_dimensions[row_num].height = 100
                new_sheet.add_image(img_obj, img_cell.coordinate)

    # Imposta la riga 2 tutta in grassetto
    for cell in new_sheet["1:1"]:
        cell.font = openpyxl.styles.Font(bold=True)

    # Salva il nuovo file XLSX
    new_workbook.save(modified_xlsx_path)

    # Mostra un messaggio di completamento
    messagebox.showinfo("Completato", "Operazione completata con successo!")

# Creazione di una finestra temporanea per ottenere i temi disponibili
temp_root = tk.Tk()
style = ttk.Style()

# Imposta le dimensioni della finestra
window_width = 400
window_height = 300
screen_width = temp_root.winfo_screenwidth()
screen_height = temp_root.winfo_screenheight()
x_coordinate = (screen_width - window_width) // 2
y_coordinate = (screen_height - window_height) // 2
temp_root.geometry(f"{window_width}x{window_height}+{x_coordinate}+{y_coordinate}")

# Campo di input per il percorso del file originale
input_path_label = tk.Label(temp_root, text="Percorso file originale:")
input_path_label.pack()
input_path = tk.Entry(temp_root)
input_path.pack()

# Pulsante per selezionare il percorso del file originale
select_input_button = tk.Button(temp_root, text="Scegli file originale", command=lambda: input_path.insert(0, filedialog.askopenfilename()))
select_input_button.pack()

# Campo di input per il percorso di destinazione
output_path_label = tk.Label(temp_root, text="Percorso file modificato:")
output_path_label.pack()
output_path = tk.Entry(temp_root)
output_path.pack()

# Pulsante per selezionare il percorso di destinazione
select_output_button = tk.Button(temp_root, text="Scegli percorso destinazione", command=lambda: output_path.insert(0, filedialog.asksaveasfilename(defaultextension=".xlsx")))
select_output_button.pack(pady=10)  # Aggiungi spazio verticale tra i pulsanti

# Pulsante per avviare il processo di modifica
process_button = tk.Button(temp_root, text="Avvia Modifica", command=lambda col_indices=col_indices: process_excel(col_indices))

process_button.pack(pady=10)  # Aggiungi spazio verticale tra i pulsanti

# Avvia l'interfaccia grafica
temp_root.mainloop()