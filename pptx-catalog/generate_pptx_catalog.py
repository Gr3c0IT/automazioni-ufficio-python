import sys
import pandas as pd
from tkinter import Tk, ttk
from pptx import Presentation
from pptx.util import Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
import re
import requests
from io import BytesIO
import ast
import os
from pptx.enum.dml import MSO_FILL

exclude_alignment_shapes = ['PZ X CF', 'CF X LY', 'CF X PLT', 'LY X PLT']

# Unisci tutte le colonne in una lista unica
logo_columns = ['TOP_SELLER', 'NOVITA', 'TREND', 'GLUTEN_FREE', 'PROTEICO', 'VEGANO', 'HALAL']

# Rinomina le shape in LOGO1, LOGO2, ..., LOGO8
logo_shapes_names = [f'LOGO{i+1}' for i in range(8)]

def set_shape_picture_fill(shape, image_path):
    # Sostituisci la shape con l'immagine
    left = shape.left
    top = shape.top
    width = shape.width
    height = shape.height
    shapes = shape._parent
    shapes._spTree.remove(shape._element)
    shapes.add_picture(image_path, left, top, width, height)

def set_shape_picture_fill_from_stream(shape, image_stream):
    left = shape.left
    top = shape.top
    width = shape.width
    height = shape.height
    shapes = shape._parent
    shapes._spTree.remove(shape._element)
    shapes.add_picture(image_stream, left, top, width, height)

# def process_top_bottom_shapes(shapes, data_row, columns, shapes_names, image_folder):
#     # Creare una lista delle immagini da inserire
#     images_to_insert = []
#     for column in columns:
#         value = data_row.get(column, '')
#         if str(value).strip().upper() == 'VERO':
#             images_to_insert.append(f"{column}.png")

def process_logo_shapes(shapes, data_row, columns, shapes_names, image_folder):
    # Creare una lista delle immagini da inserire
    images_to_insert = []
    
    # Raccoglie tutte le immagini disponibili
    for column in columns:
        value = data_row.get(column, '')
        if value is True or str(value).strip().upper() == 'VERO' or str(value).strip() == 'True':
            images_to_insert.append(f"{column}.png")
    
    # Dizionario per tenere traccia delle shape LOGO
    logo_shapes = {}
    
    # Prima passata: trova tutte le shape LOGO e salvale nel dizionario
    for shape in shapes:
        if shape.shape_type == 6:  # Se è un gruppo
            process_logo_shapes(shape.shapes, data_row, columns, shapes_names, image_folder)
        elif shape.has_text_frame:
            text = shape.text.strip()
            if text in shapes_names:
                logo_shapes[text] = shape
    
    # Seconda passata: inserisci le immagini in ordine, partendo sempre da LOGO1
    for i, image_name in enumerate(images_to_insert):
        logo_shape_name = f'LOGO{i+1}'
        if logo_shape_name in logo_shapes:
            shape = logo_shapes[logo_shape_name]
            image_path = os.path.join(image_folder, image_name)
            if os.path.exists(image_path):
                set_shape_picture_fill(shape, image_path)
            else:
                print(f"Immagine non trovata: {image_path}")
    
    # Pulisci le shape LOGO rimanenti
    for i in range(len(images_to_insert), len(shapes_names)):
        logo_shape_name = f'LOGO{i+1}'
        if logo_shape_name in logo_shapes:
            logo_shapes[logo_shape_name].text = ''

def format_currency(value):
    # Verifica se il valore è un numero
    if pd.notna(value) and isinstance(value, (int, float)):
        # Formatta il numero come valuta, con due decimali e il simbolo €
        return f"{value:.2f} €"
    return ''  # Restituisce una stringa vuota se il valore non è valido

def process_shapes(shapes, data_row, column_styles):
    int_columns = ["pz_cf", "cf_ly", "cf_plt", "ly_plt", "ean_pz", "ean_cf"]
    skip_texts = logo_shapes_names
    for shape_idx, shape in enumerate(shapes):
        if shape.shape_type == 6:
            process_shapes(shape.shapes, data_row, column_styles)
        elif shape.has_text_frame:
            text = shape.text.strip()

            if text in exclude_alignment_shapes:
                continue
            if text in skip_texts:
                continue
            if shape.text.strip() == "Immagine":
                image_url = data_row['Immagini']
                image_stream = download_image(image_url)
                if image_stream:
                    set_shape_picture_fill_from_stream(shape, image_stream)
                else:
                    print(f"Immagine non scaricata: {image_url}")
            else:
                for paragraph in shape.text_frame.paragraphs:
                    paragraph.alignment = PP_ALIGN.LEFT
                    for run in paragraph.runs:
                        text = run.text
                        for column, style in column_styles.items():
                            placeholder_pattern = re.compile(rf'\b{column}\b', re.IGNORECASE)
                            if re.search(placeholder_pattern, text):
                                if column == "Prezzo_Rivendita_Suggerito":
                                    replacement = format_currency(data_row.get(column, None))
                                elif column in int_columns:
                                    try:
                                        replacement = str(int(float(data_row[column]))) if pd.notna(data_row[column]) else ''
                                    except ValueError:
                                        replacement = str(data_row[column]) if pd.notna(data_row[column]) else ''
                                elif column == "Peso":
                                    peso_value = data_row.get(column, None)
                                    categoria_value = data_row.get('Categoria', '')
                                    if pd.notna(peso_value):
                                        try:
                                            peso_float = float(peso_value)
                                            if peso_float.is_integer():
                                                replacement = str(int(peso_float))
                                            else:
                                                replacement = str(peso_float)
                                            if pd.notna(categoria_value) and isinstance(categoria_value, str):
                                                categoria = categoria_value.strip().lower()
                                                if categoria == 'bevande':
                                                    replacement += " ml"
                                                else:
                                                    replacement += " g"
                                            else:
                                                replacement += " g"
                                        except ValueError:
                                            replacement = str(peso_value)
                                    else:
                                        replacement = ''  
                                else:
                                    replacement = str(data_row[column]) if pd.notna(data_row[column]) else ''

                                # Sostituisci il testo nel run
                                new_text = re.sub(placeholder_pattern, replacement, text)
                                run.text = new_text

                                # Applica lo stile al run
                                run.font.name = style['font']
                                run.font.size = Pt(style['size'])
                                run.font.bold = True if style['bold'] else False
                                run.font.italic = True if style['italic'] else False
                                run.font.color.rgb = RGBColor.from_string(style['color'].lstrip('#'))

                                # Imposta l'allineamento del paragrafo
                                if text not in exclude_alignment_shapes:
                                    if style['alignment'] == "left":
                                        paragraph.alignment = PP_ALIGN.LEFT
                                    elif style['alignment'] == "center":
                                        paragraph.alignment = PP_ALIGN.CENTER
                                    elif style['alignment'] == "right":
                                        paragraph.alignment = PP_ALIGN.RIGHT
        else:
            continue



# Funzioni per ricordare e ripristinare le posizioni delle shape e dei gruppi
def remember_group_positions(shapes, group_positions):
    for shape in shapes:
        group_positions[shape.shape_id] = {
            'left': shape.left,
            'top': shape.top,
            'width': shape.width,
            'height': shape.height
        }
        if shape.shape_type == 6:  # GROUP
            # Ricorsione nei gruppi interni
            remember_group_positions(shape.shapes, group_positions)

def restore_group_positions(shapes, group_positions):
    for shape in shapes:
        if shape.shape_id in group_positions:
            pos = group_positions[shape.shape_id]
            shape.left = pos['left']
            shape.top = pos['top']
            shape.width = pos['width']
            shape.height = pos['height']
        if shape.shape_type == 6:  # GROUP
            # Ricorsione nei gruppi interni
            restore_group_positions(shape.shapes, group_positions)

# Funzioni ausiliarie
def download_image(url):
    try:
        response = requests.get(url)
        response.raise_for_status()
        return BytesIO(response.content)
    except requests.exceptions.RequestException as e:
        print(f"Errore nel download dell'immagine: {url} - Errore: {e}")
        return None
    
def apply_text_style(paragraph, text_style):
    if paragraph.runs:
        for run in paragraph.runs:
            # Applica le proprietà del font, mantenendo i valori originali se non specificati
            run.font.name = text_style.get('font', run.font.name)
            run.font.size = Pt(text_style.get('size', run.font.size.pt))
            run.font.bold = True if text_style.get('bold', run.font.bold) else False
            run.font.italic = True if text_style.get('italic', run.font.italic) else False
            if 'color' in text_style and text_style['color']:
                run.font.color.rgb = RGBColor.from_string(text_style['color'].lstrip('#'))
        # Applicazione dell'allineamento solo se definito nello stile
        if 'alignment' in text_style and text_style['alignment']:
            if text_style['alignment'] == "left":
                paragraph.alignment = PP_ALIGN.LEFT
            elif text_style['alignment'] == "center":
                paragraph.alignment = PP_ALIGN.CENTER
            elif text_style['alignment'] == "right":
                paragraph.alignment = PP_ALIGN.RIGHT
    else:
        print("Il paragrafo non contiene testo o runs da stilizzare.")




def main(excel_path, ppt_path, column_styles_str):
    global df
    df = pd.read_excel(excel_path)
    column_styles = ast.literal_eval(column_styles_str)
    presentation = Presentation(ppt_path)

    # Definisci il percorso della cartella delle immagini
    image_folder = os.path.join(os.path.dirname(os.path.abspath(__file__)), "loghi")

    # Creazione della finestra di progresso
    # ... codice esistente per la barra di progresso ...

    row_idx = 0
    for slide_idx, slide in enumerate(presentation.slides):
        print(f"Processing Slide {slide_idx + 1}")

        group_positions = {}
        remember_group_positions(slide.shapes, group_positions)

        for shape in slide.shapes:
            if shape.shape_type == 6:
                if row_idx < len(df):
                    data_row = df.iloc[row_idx]
                    # Processa le shape LOGO
                    process_logo_shapes([shape], data_row, logo_columns, logo_shapes_names, image_folder)
                    # Processa le altre shape
                    process_shapes([shape], data_row, column_styles)
                    row_idx += 1
                else:
                    print("No more data rows available to populate.")
            else:
                continue

        restore_group_positions(slide.shapes, group_positions)

        # Aggiornare la barra di progresso e la percentuale
        # ... codice esistente ...

    presentation.save("catalogo_pronto.pptx")

    # Chiudere la finestra di progresso alla fine
    # ... codice esistente ...


if __name__ == "__main__":
    if len(sys.argv) == 4:
        excel_path = sys.argv[1]
        ppt_path = sys.argv[2]
        column_styles_str = sys.argv[3]
        main(excel_path, ppt_path, column_styles_str)
    else:
        print("Uso: python generate_pptx_catalog.py <excel_path> <ppt_path> <column_styles>")
