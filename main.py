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
    "Metr.(von)": ["metr", "meter", "metr.", "Metr.", "Metrier.", "metrier.", "Metrierung", "metrierung"],
    "bis Ort": ["bis ort", "ziel ort", "end ort"],
    "bis km": ["bis km", "ziel km", "end km"],
    "Metr.(bis)": ["metr", "meter", "metr.","Metr.", "Metrier.", "metrier.", "Metrierung", "metrierung"],
    "SOLL": ["soll", "sollwert", "soll m"],
    "IST": ["ist", "istwert", "ist m"],
    "Verlegeart": ["verlegeart", "verlegungsart", "Verlegeart Hand/Masch.", "VerlegeartHand/Masch.","verlegeart Hand/Masch.",],
    "Bemerkung": ["Bemerkungen","bemerkung","bemerkungen", "notiz", "kommentar", "Kommentar", "Anmerkung", "anmerkung","Bemerkungen Besonderheiten","BemerkungenBesonderheiten","Besonderheiten","besonderheiten"]
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
    for key, syns in HEADER_MAP.items():
        if difflib.get_close_matches(t, [key.lower()] + [s.lower() for s in syns], n=1, cutoff=0.7):
            return key
    return None

def match_header_prefer_exact(text):
    if not isinstance(text, str): 
        return None
    t = text.strip().lower()
    # 1) Exaktes Match zuerst prüfen
    for key, syns in HEADER_MAP.items():
        if t in [key.lower()] + [s.lower() for s in syns]:
            return key
    # 2) Falls kein exaktes Match → unscharfe Suche
    for key, syns in HEADER_MAP.items():
        if difflib.get_close_matches(t, [key.lower()] + [s.lower() for s in syns], n=1, cutoff=0.7):
            return key
    return None

def extract_data_from_pdf(pdf_path):
    alle_daten = []
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        with pdfplumber.open(pdf_path) as pdf:
            beste_score = 0
            beste_tabelle = None
            beste_header_zeile = None

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
                        if score > beste_score:
                            beste_score = score
                            beste_header_zeile = zeile_idx
                            beste_tabelle = tabelle

            # Standardweg: wenn ausreichend bekannte Header erkannt wurden
            if beste_tabelle and beste_score >= 10:
                daten_ab_header = beste_tabelle[beste_header_zeile:]
                header = daten_ab_header[0]
                try:
                    df = pd.DataFrame(daten_ab_header[1:], columns=make_unique(header))
                    alle_daten.append(df)
                except Exception:
                    pass

            # Fallback: spaltenorientiert Header suchen
            if not alle_daten:
                for seite in pdf.pages:
                    try:
                        tables = seite.extract_tables()
                    except Exception:
                        continue
                    if not tables:
                        continue

                    for tabelle in tables:
                        if not tabelle or len(tabelle) < 2:
                            continue

                        spalten = list(zip(*tabelle))
                        neue_header = []
                        inhalt_nach_header = []

                        for spalte in spalten:
                            header_idx = None
                            header_name = None
                            found_exact = False

                            # Erst nach exaktem Match suchen
                            for idx, zelle in enumerate(spalte):
                                if not isinstance(zelle, str):
                                    continue
                                t = zelle.strip().lower()
                                for key, syns in HEADER_MAP.items():
                                    if t in [key.lower()] + [s.lower() for s in syns]:
                                        header_idx = idx
                                        header_name = zelle
                                        found_exact = True
                                        break
                                if found_exact:
                                    break

                            # Falls kein exaktes Match gefunden, fuzzy suchen
                            if header_idx is None:
                                for idx, zelle in enumerate(spalte):
                                    header = match_header(zelle)
                                    if header:
                                        header_idx = idx
                                        header_name = zelle
                                        break

                            if header_idx is not None:
                                neue_header.append(header_name)
                                inhalt_nach_header.append(list(spalte[header_idx+1:]))
                            else:
                                neue_header.append(spalte[0] or f"unknown_{len(neue_header)}")
                                inhalt_nach_header.append(list(spalte[1:]))

                        daten_zeilen = list(zip(*inhalt_nach_header))

                        try:
                            df = pd.DataFrame(daten_zeilen, columns=make_unique(neue_header))
                            df = df.dropna(how='all')  # komplett leere Zeilen entfernen
                            df = df[df.notna().sum(axis=1) >= 6]  # nur Zeilen mit ≥ 3 Werten behalten
                            if not df.empty:
                                alle_daten.append(df)
                        except Exception:
                            continue

    return pd.concat(alle_daten, ignore_index=True) if alle_daten else pd.DataFrame()

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
