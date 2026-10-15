import tkinter as tk
from tkinter import filedialog, colorchooser, font
import os
import subprocess
import sys
import pandas as pd
from tkinter import messagebox

# Mappatura tra etichette italiane e valori inglesi
alignment_map = {
    "Sinistra": "left",
    "Centro": "center",
    "Destra": "right"
}

def check_if_styles_saved():
    if not column_styles:
        messagebox.showwarning("Nessuna Modifica", "Non è stata selezionata ancora alcuna modifica.\nDecidere se modificare o meno la Slide.")

# Variabili globali per memorizzare i percorsi dei file caricati e gli stili definiti
excel_file_path = None
ppt_file_path = None
column_styles = {}
columns = []  # Variabile per memorizzare i nomi delle colonne

# Funzione per caricare un file Excel e leggere i nomi delle colonne
def load_excel_file():
    global excel_file_path, columns
    excel_file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
    if excel_file_path:
        print(f"File Excel caricato: {excel_file_path}")
        df = pd.read_excel(excel_file_path)
        columns = list(df.columns)
        create_style_section()

# Funzione per caricare un file PowerPoint
def load_ppt_file():
    global ppt_file_path
    ppt_file_path = filedialog.askopenfilename(filetypes=[("PowerPoint files", "*.pptx")])
    if ppt_file_path:
        print(f"File PowerPoint caricato: {ppt_file_path}")
        check_if_styles_saved()  # Controlla se ci sono stili salvati
        # Puoi continuare a eseguire altre operazioni necessarie


# Funzione per selezionare il colore del font
def choose_color(label):
    color_code = colorchooser.askcolor(title="Scegli colore del font")[1]
    if color_code:
        label.config(text=color_code, bg=color_code)

# Funzione per salvare lo stile per le colonne selezionate
def save_selected_styles(font_var, size_spinbox, color_label, bold_var, italic_var, alignment_var, selected_columns, checkboxes):
    total_columns = len(selected_columns)
    processed_columns = 0
    
    for column_name in selected_columns:
        column_styles[column_name] = {
            'font': font_var.get(),
            'size': int(size_spinbox.get()),
            'color': color_label.cget("text"),
            'bold': bold_var.get(),
            'italic': italic_var.get(),
            'alignment': alignment_var.get()
        }
        processed_columns += 1
        print(f"Stile salvato per {column_name}: {column_styles[column_name]}")
    
    # Controlla se tutte le colonne sono state processate
    if processed_columns == total_columns:
        # Dopo aver iterato su tutte le colonne selezionate, deseleziona le checkbox
        for var in checkboxes.values():
            var.set(False)
        
        # Cancella la lista delle colonne selezionate
        selected_columns.clear()
        print("Le caselle selezionate sono state deselezionate.")

def create_style_section():
    # Rimuovi tutti i widget attuali
    for widget in root.winfo_children():
        widget.destroy()

    # Creazione del frame principale
    main_frame = tk.Frame(root)
    main_frame.pack(fill="both", expand=True, padx=20, pady=20)

    # Creazione della sezione di selezione delle colonne
    columns_frame = tk.LabelFrame(main_frame, text="Seleziona le colonne a cui applicare lo stile", padx=10, pady=10)
    columns_frame.grid(row=0, column=0, padx=20, pady=20, sticky="n")

    selected_columns = []

    def on_checkbox_select():
        selected_columns.clear()
        for column, var in checkboxes.items():
            if var.get():
                selected_columns.append(column)
        print(f"Colonne selezionate: {selected_columns}")

    checkboxes = {}
    for column in columns:
        if column.lower() in ["immagine", "immagini", "top_seller", "novita", "trend", "gluten_free", "proteico", "vegano", "halal"]:
            continue
        print(f"Creazione checkbox per colonna: {column}")
        var = tk.BooleanVar()
        checkbox = tk.Checkbutton(columns_frame, text=column, variable=var, command=on_checkbox_select)
        checkbox.pack(anchor="w")
        checkboxes[column] = var  # Usa il nome della colonna come chiave

    if not columns:
        print("Nessuna colonna trovata")

    # Creazione della sezione di stile del font
    style_frame = tk.LabelFrame(main_frame, text="Definisci lo stile del font", padx=10, pady=10)
    style_frame.grid(row=0, column=1, padx=20, pady=20, sticky="n")

    font_var = tk.StringVar(value="Poppins")
    font_entry = tk.Entry(style_frame)
    font_entry.pack(pady=5)

    font_listbox = tk.Listbox(style_frame, height=4)
    font_listbox.pack(pady=5)

    for f in font.families():
        font_listbox.insert(tk.END, f)

    def on_font_select(event):
        selected_font = font_listbox.get(font_listbox.curselection())
        font_var.set(selected_font)
        font_entry.delete(0, tk.END)
        font_entry.insert(0, selected_font)

    font_listbox.bind("<<ListboxSelect>>", on_font_select)
    font_entry.bind("<KeyRelease>", lambda event: update_font_list(font_entry, font_var, font_listbox))

    size_spinbox = tk.Spinbox(style_frame, from_=6, to=72)
    color_label = tk.Label(style_frame, text="#000000", bg="#000000", width=10)
    bold_var = tk.IntVar()
    italic_var = tk.IntVar()
    # Variabile di allineamento con valore predefinito
    alignment_var = tk.StringVar(value="left")

    # Funzione per aggiornare alignment_var in base alla selezione italiana
    def update_alignment(value):
        alignment_var.set(alignment_map[value])


    tk.Label(style_frame, text="Dimensione Font:").pack(pady=5)
    size_spinbox.pack(pady=5)

    tk.Label(style_frame, text="Colore Font:").pack(pady=5)
    color_button = tk.Button(style_frame, text="Scegli Colore", command=lambda lbl=color_label: choose_color(lbl))
    color_button.pack(pady=5)
    color_label.pack(pady=5)

    tk.Checkbutton(style_frame, text="Grassetto", variable=bold_var).pack(pady=5)
    tk.Checkbutton(style_frame, text="Corsivo", variable=italic_var).pack(pady=5)

    tk.Label(style_frame, text="Allineamento Testo:").pack(pady=5)
    alignment_menu = tk.OptionMenu(style_frame, tk.StringVar(value="Sinistra"), *alignment_map.keys(), command=update_alignment)
    alignment_menu.pack(pady=5)

    tk.Button(style_frame, text="Salva stile per le colonne selezionate",
          command=lambda: save_selected_styles(font_var, size_spinbox, color_label, bold_var, italic_var, alignment_var, selected_columns, checkboxes)).pack(pady=10)

    # Creazione dei pulsanti per caricare il file PowerPoint ed eseguire lo script
    ppt_button = tk.Button(main_frame, text="Carica File PowerPoint", command=load_ppt_file)
    execute_button = tk.Button(main_frame, text="Esegui Script", command=esegui_main_script)

    ppt_button.grid(row=1, column=0, pady=10, sticky="e")
    execute_button.grid(row=1, column=1, pady=10, sticky="e")



# Funzione per aggiornare l'elenco dei font in base alla ricerca
def update_font_list(entry, font_var, listbox):
    search_term = entry.get().lower()
    fonts = [f for f in font.families() if search_term in f.lower()]
    listbox.delete(0, tk.END)
    for f in fonts:
        listbox.insert(tk.END, f)

# Funzione per eseguire il file generate_pptx_catalog.py
def esegui_main_script():
    if excel_file_path or ppt_file_path:
        script = os.path.join(os.path.dirname(os.path.abspath(__file__)), "generate_pptx_catalog.py")
        args = [sys.executable, script, excel_file_path, ppt_file_path, str(column_styles)]
        try:
            subprocess.run(args, check=True)
            print("Script eseguito con successo!")
            # Mostra il pop-up dopo il completamento dello script
            messagebox.showinfo("Completato", "Lavoro eseguito!")
        except subprocess.CalledProcessError as e:
            print(f"Errore nell'esecuzione dello script: {e}")
            messagebox.showerror("Errore", "Errore nell'esecuzione dello script.")
    else:
        print("Caricare almeno un file prima di eseguire lo script.")
        messagebox.showwarning("Attenzione", "Caricare almeno un file prima di eseguire lo script.")


# Creazione della finestra principale
root = tk.Tk()
root.title("Caricamento File")
root.geometry("650x650")

excel_button = tk.Button(root, text="Carica File Excel", command=load_excel_file)
excel_button.pack(pady=20)

root.mainloop()
