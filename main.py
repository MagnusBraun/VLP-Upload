from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse
import shutil
import uuid
import os
import pandas as pd
import pdfplumber
import difflib
import re
import warnings

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["https://magnusbraun.github.io"],  # exakt deine GitHub Pages Domain
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)
HEADER_MAP = {
    "Kabelnummer": ["kabelnummer", "kabel-nummer", "Kabel-nummer", "Kabel-Nummer", "Kabel-Nr","Kabel-Nr.", "Kabel-nr", "Kabel-nr.", "kabel-nr", "kabel-nr."],
    "Kabeltyp": ["kabeltyp", "typ", "Kabeltype", "Kabel-type", "Kabel-Type"],
    "Ømm": ["durchmesser", "ø", "Ø", "ømm", "mm", "Durchmesser in mm","durchmesser in mm", "Durch-messer in mm", "durch-messer in mm"],
    "Trommelnummer": ["Trommel", "trommelnummer", "Trommel-nummer"],
    "von Ort": ["von ort", "start ort"],
    "von km": ["von km", "start km", "anfang km"],
    "Metr.(von)": ["metr", "meter", "metr.", "Metr."],
    "bis Ort": ["bis ort", "ziel ort", "end ort"],
    "bis km": ["bis km", "ziel km", "end km"],
    "Metr.(bis)": ["metr", "meter", "metr.","Metr."],
    "SOLL": ["soll", "sollwert", "soll m"],
    "IST": ["ist", "istwert", "ist m"],
    "Verlegeart": ["verlegeart", "verlegung", "verlegungsart"],
    "Bemerkung": ["Bemerkungen","bemerkung","bemerkungen", "notiz", "kommentar", "Kommentar", "Anmerkung", "anmerkung"]
}

def make_unique(columns):
    seen = {}
    result = []
    for col in columns:
        col = str(col) if not isinstance(col, str) else col
        if col in seen:
            seen[col] += 1
            result.append(f"{col}_{seen[col]}")
        else:
            seen[col] = 0
            result.append(col)
    return result

def match_header(text):
    if not isinstance(text, str): return None
    t = text.strip().lower()
    t = re.sub(r"[^a-z0-9]", "", t)  # entferne Trennzeichen
    for key, syns in HEADER_MAP.items():
        for candidate in [key] + syns:
            if re.sub(r"[^a-z0-9]", "", candidate.lower()) in t:
                return key
    return None

def extract_data_from_pdf(pdf_path):
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        with pdfplumber.open(pdf_path) as pdf:
            # Phase 1: normale Header-Zeile suchen
            for seite in pdf.pages:
                try:
                    tables = seite.extract_tables()
                except Exception:
                    continue
                if not tables:
                    continue

                for tabelle in tables:
                    for zeile_idx, row in enumerate(tabelle):
                        score = sum(1 for cell in row if match_header(cell))
                        if score >= 10:
                            header = make_unique(row)
                            try:
                                df = pd.DataFrame(tabelle[zeile_idx + 1:], columns=header)
                                if not df.empty:
                                    return df
                            except Exception as e:
                                print("⚠️ Fehler beim Aufbau des DataFrames (normal):", e)
                                continue

            # Phase 2: Fallback – kombinierte Headerzeile (2 Zeilen)
            for seite in pdf.pages:
                try:
                    tables = seite.extract_tables()
                except Exception:
                    continue
                if not tables:
                    continue

                for tabelle in tables:
                    for zeile_idx in range(len(tabelle) - 1):
                        zeile1 = tabelle[zeile_idx]
                        zeile2 = tabelle[zeile_idx + 1]
                        if not zeile1 or not zeile2:
                            continue
                        combined = [
                            f"{(zeile1[i] or '').strip()} {(zeile2[i] or '').strip()}".strip()
                            for i in range(min(len(zeile1), len(zeile2)))
                        ]
                        score = sum(1 for cell in combined if match_header(cell))
                        if score >= 6:
                            header = make_unique(combined)
                            try:
                                df = pd.DataFrame(tabelle[zeile_idx + 2:], columns=header)
                                if not df.empty:
                                    return df
                            except Exception as e:
                                print("⚠️ Fehler beim Aufbau des DataFrames (fallback):", e)
                                continue
    return pd.DataFrame()


def map_columns_to_headers(df):
    mapped = {}
    metr_spalten = []
    for col in df.columns:
        header = match_header(col)
        if header in ["Metr.(von)", "Metr.(bis)"]:
            metr_spalten.append(col)
            continue
        if header:
            values = df[col].dropna().astype(str).tolist()
            mapped.setdefault(header, []).extend(values)
    for i, col in enumerate(metr_spalten):
        values = df[col].dropna().astype(str).tolist()
        target = "Metr.(von)" if i == 0 else "Metr.(bis)"
        mapped.setdefault(target, []).extend(values)
    return mapped

@app.post("/process")
def process_pdf(file: UploadFile = File(...)):
    if not file.filename.lower().endswith(".pdf"):
        raise HTTPException(status_code=400, detail="Nur PDF-Dateien erlaubt")
    file_id = str(uuid.uuid4())
    temp_path = os.path.join("/tmp", f"{file_id}.pdf")
    with open(temp_path, "wb") as f:
        shutil.copyfileobj(file.file, f)
    df = extract_data_from_pdf(temp_path)
    if df.empty:
        raise HTTPException(status_code=422, detail="Keine verarbeitbaren Tabellen gefunden")
    mapped = map_columns_to_headers(df)
    return JSONResponse(content=mapped)
