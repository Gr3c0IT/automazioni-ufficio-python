#!/usr/bin/env python3
"""
Scarica immagini da un file Excel.
Colonna A: codice articolo (es. ART002)
Colonna B: URL immagine (es. https://...)
Le immagini vengono salvate come <codice>.<estensione> nella cartella di output.
"""

import sys
import os
import time
import requests
import pandas as pd
from pathlib import Path
from urllib.parse import urlparse

# ─── Configurazione ────────────────────────────────────────────────────────────
EXCEL_FILE   = "articoli.xlsx"          # file xlsx di input
OUTPUT_DIR   = "immagini_articoli"      # cartella di destinazione
SHEET_NAME   = 0                        # 0 = primo foglio, oppure nome stringa
SKIP_HEADER  = True                     # True se la prima riga è un'intestazione
COL_CODICE   = 0                        # indice colonna codice  (A = 0)
COL_URL      = 1                        # indice colonna URL     (B = 1)
TIMEOUT      = 15                       # secondi per ogni richiesta HTTP
DELAY        = 0.3                      # pausa tra un download e l'altro (sec)
# ───────────────────────────────────────────────────────────────────────────────


def estensione_da_url(url: str, content_type: str = "") -> str:
    """Ricava l'estensione dal path dell'URL o dal Content-Type."""
    # Prima prova dall'URL
    path = urlparse(url).path
    ext = Path(path).suffix.lower()  # es. ".jpg"
    if ext in {".jpg", ".jpeg", ".png", ".gif", ".webp", ".bmp", ".svg"}:
        return ext

    # Fallback sul Content-Type
    ct_map = {
        "image/jpeg": ".jpg",
        "image/png":  ".png",
        "image/gif":  ".gif",
        "image/webp": ".webp",
        "image/bmp":  ".bmp",
        "image/svg+xml": ".svg",
    }
    for mime, e in ct_map.items():
        if mime in content_type:
            return e

    return ".jpg"  # default


def scarica_immagine(codice: str, url: str, output_dir: Path) -> bool:
    """Scarica una singola immagine e la salva con il nome del codice articolo."""
    url = url.strip()
    if not url:
        print(f"  [SKIP] {codice}: URL vuoto")
        return False

    try:
        resp = requests.get(url, timeout=TIMEOUT, stream=True)
        resp.raise_for_status()

        content_type = resp.headers.get("Content-Type", "")
        ext = estensione_da_url(url, content_type)

        # Sanifica il codice per usarlo come nome file
        nome_file = "".join(c for c in codice if c.isalnum() or c in "-_") + ext
        dest = output_dir / nome_file

        with open(dest, "wb") as f:
            for chunk in resp.iter_content(chunk_size=8192):
                f.write(chunk)

        size_kb = dest.stat().st_size / 1024
        print(f"  [OK]   {nome_file}  ({size_kb:.1f} KB)")
        return True

    except requests.exceptions.HTTPError as e:
        print(f"  [ERR]  {codice}: HTTP {e.response.status_code} → {url}")
    except requests.exceptions.ConnectionError:
        print(f"  [ERR]  {codice}: impossibile connettersi → {url}")
    except requests.exceptions.Timeout:
        print(f"  [ERR]  {codice}: timeout dopo {TIMEOUT}s → {url}")
    except Exception as e:
        print(f"  [ERR]  {codice}: {e}")

    return False


def main():
    # DOPO (corretto — relativo alla cartella dello script)
    script_dir = Path(__file__).parent
    excel_path = Path(sys.argv[1]) if len(sys.argv) > 1 else script_dir / EXCEL_FILE

    if not excel_path.exists():
        sys.exit(f"Errore: file non trovato → {excel_path}")

    # Leggi il foglio
    header = 0 if SKIP_HEADER else None
    df = pd.read_excel(
        excel_path,
        sheet_name=SHEET_NAME,
        header=header,
        dtype=str,          # tutto come stringa, evita conversioni indesiderate
        usecols=[COL_CODICE, COL_URL],
    )
    df.columns = ["codice", "url"]
    df.dropna(subset=["codice", "url"], inplace=True)
    df = df[df["url"].str.startswith("http")]  # filtra righe senza URL valido

    if df.empty:
        sys.exit("Nessuna riga valida trovata nel file Excel.")

    # Crea cartella di output
    output_dir = Path(OUTPUT_DIR)
    output_dir.mkdir(parents=True, exist_ok=True)

    print(f"\nFile Excel  : {excel_path}")
    print(f"Cartella out: {output_dir.resolve()}")
    print(f"Articoli    : {len(df)}\n")

    ok = ko = 0
    for _, row in df.iterrows():
        if scarica_immagine(row["codice"], row["url"], output_dir):
            ok += 1
        else:
            ko += 1
        time.sleep(DELAY)

    print(f"\nCompletato — Scaricati: {ok}  Errori: {ko}")


if __name__ == "__main__":
    main()