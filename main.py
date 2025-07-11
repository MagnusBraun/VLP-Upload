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
import logging
from pdf2image import convert_from_path
import pytesseract
from pdfminer.high_level import extract_pages
from pdfminer.layout import LTTextLineHorizontal
logging.basicConfig(level=logging.INFO)

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["https://magnusbraun.github.io"],
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
    "von km": ["von km", "start km", "anfang km", "von Ort Bahn km", "von ort bahn km"],
    "Metr.(von)": ["metr", "meter", "metr.", "Metr.", "Metrier.", "metrier.", "Metrierung", "metrierung", "Start Meter A"],
    "bis Ort": ["bis ort", "ziel ort", "end ort"],
    "bis km": ["bis km", "ziel km", "end km", "nach Ort Bahn km", "nach ort Bahn km"],
    "Metr.(bis)": ["metr", "meter", "metr.","Metr.", "Metrier.", "metrier.", "Metrierung", "metrierung", "Ende Meter E"],
    "SOLL": ["soll", "sollwert", "soll m"],
    "IST": ["ist", "istwert", "ist m", "Verlegte-Länge","Verlegte- Länge"],
    "Verlegeart": ["verlegeart", "verlegungsart", "Verlegeart Hand/Masch.", "VerlegeartHand/Masch.","verlegeart Hand/Masch."],
    "Bemerkung": ["Bemerkungen","bemerkung","bemerkungen", "notiz", "kommentar", "Kommentar", "Anmerkung", "anmerkung","Bemerkungen Besonderheiten","BemerkungenBesonderheiten","Besonderheiten","besonderheiten"]
}

# -------------------- NEU: KÜP Positionsbasiert --------------------

def extract_kabel_from_page(texts):
    kabelnummer_rx = re.compile(r'\bS[\s-]?\d{3,7}\b', re.I)
    kabeltyp_rx = re.compile(r'\d+[x×]\d+(?:[.,]\d+)?(?:[x×]\d+)?', re.I)
    laenge_rx = re.compile(r'\d+\s?m\b', re.I)

    kabel_liste = []

    for t in texts:
        text = t['text'].strip()
        if kabelnummer_rx.match(text):
            x0, top = t['x0'], t['top']

            kabelnummer = text
            kabeltyp = find_text_below(texts, x0, top, kabeltyp_rx)
            laenge = find_text_right(texts, x0, top, laenge_rx)

            if laenge:
                laenge = re.sub(r"\s*\(.*?\)", "", laenge).strip()  # Eingeklammerte Werte entfernen

            kabel_liste.append({
                "Kabelname": kabelnummer,
                "Kabeltyp": kabeltyp,
                "SOLL": laenge
            })

    return kabel_liste


def find_text_below(texts, ref_x, ref_top, pattern_rx, tolerance=40):
    for t in texts:
        if abs(t['x0'] - ref_x) < tolerance and t['top'] > ref_top:
            if pattern_rx.search(t['text']):
                return t['text']
    return None

def find_text_right(texts, ref_x, ref_top, pattern_rx, tolerance=40):
    for t in texts:
        if abs(t['top'] - ref_top) < tolerance and t['x0'] > ref_x:
            if pattern_rx.search(t['text']):
                return t['text']
    return None

# ------------------------------------------------------------------

def find_nearest_text(texts, ref, pattern_rx, max_dist=50):
    ref_x, ref_top = ref['x0'], ref['top']
    nearest = None
    min_dist = float('inf')

    for t in texts:
        if t == ref:
            continue
        text = t['text'].strip()
        if not pattern_rx.search(text):
            continue
        dist = ((t['x0'] - ref_x)**2 + (t['top'] - ref_top)**2)**0.5
        if dist < min_dist and dist <= max_dist:
            nearest = t
            min_dist = dist
    return nearest['text'] if nearest else None

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
    for key, syns in HEADER_MAP.items():
        if t in [key.lower()] + [s.lower() for s in syns]:
            return key
    for key, syns in HEADER_MAP.items():
        if difflib.get_close_matches(t, [key.lower()] + [s.lower() for s in syns], n=1, cutoff=0.7):
            return key
    return None

@app.post("/upload_kuep_file")
def upload_kuep_file(file: UploadFile = File(...)):
    if not file.filename.lower().endswith(".pdf"):
        raise HTTPException(status_code=400, detail="Nur PDF erlaubt")
    file_id = str(uuid.uuid4())
    path = f"/tmp/{file_id}.pdf"
    with open(path, "wb") as f:
        shutil.copyfileobj(file.file, f)
    return {"file_id": file_id}

def extract_kabel_clusters(pdf_path, page_number=0):
    kabelnummer_rx = re.compile(r'\bS\d{4,7}\b', re.I)
    kabeltyp_rx = re.compile(r'\d+\s*[x×]\s*\d+(?:[.,]\d+)?(?:\s*[x×]\s*\d+)?', re.I)
    laenge_rx = re.compile(r'\d{2,5}\s*m\b', re.I)

    kabelinfos = []
    page_found = False

    for page_idx, page_layout in enumerate(extract_pages(pdf_path)):
        if page_idx != page_number:
            continue
        page_found = True

        textboxes = [
            {
                "text": el.get_text().strip(),
                "x0": el.bbox[0],
                "y0": el.bbox[1],
                "x1": el.bbox[2],
                "y1": el.bbox[3],
                "cx": (el.bbox[0] + el.bbox[2]) / 2,
                "cy": (el.bbox[1] + el.bbox[3]) / 2,
            }
            for el in page_layout if isinstance(el, LTTextLineHorizontal)
        ]

        for t in textboxes:
            if not kabelnummer_rx.match(t["text"]):
                continue

            kabelnummer = t["text"]
            x0, y0 = t["x0"], t["y0"]

            kabeltyp = next(
                (tb["text"] for tb in textboxes
                 if abs(tb["x0"] - x0) < 50 and 0 < (tb["y0"] - y0) < 70 and kabeltyp_rx.search(tb["text"])),
                None
            )

            laenge = next(
                (tb["text"] for tb in textboxes
                 if abs(tb["y0"] - y0) < 30 and (tb["x0"] - x0) > 100 and laenge_rx.search(tb["text"])),
                None
            )

            if laenge:
                laenge = re.sub(r"\s*\(.*?\)", "", laenge).strip()

            kabelinfos.append({
                "Kabelname": kabelnummer,
                "Kabeltyp": kabeltyp,
                "SOLL": laenge
            })

    if not page_found:
        raise ValueError("Seite nicht gefunden im PDF")

    return kabelinfos
from pdf2image import convert_from_path
import pytesseract

def extract_kabel_with_ocr(pdf_path, page_number=0):
    kabelnummer_rx = re.compile(r'\bS\d{4,7}\b', re.I)
    kabeltyp_rx = re.compile(r'\d+\s*[x×]\s*\d+(?:[.,]\d+)?(?:\s*[x×]\s*\d+)?', re.I)
    laenge_rx = re.compile(r'\d{2,5}\s*m\b', re.I)

    try:
        images = convert_from_path(pdf_path, first_page=page_number + 1, last_page=page_number + 1, dpi=300)
    except Exception as e:
        raise RuntimeError(f"PDF-Konvertierung fehlgeschlagen: {e}")

    if not images:
        return []

    ocr_text = pytesseract.image_to_string(images[0])
    
    # DEBUG: Zeige den gesamten erkannten Text im Terminal
    print(f"[OCR-DEBUG] Seite {page_number}:\n{ocr_text}")
    lines = [line.strip() for line in ocr_text.splitlines() if line.strip()]

    kabelinfos = []

    for i, line in enumerate(lines):
        if kabelnummer_rx.search(line):
            kabelnummer = kabelnummer_rx.search(line).group()
            kabeltyp = None
            laenge = None

            # Suche in den nächsten 1–3 Zeilen nach Typ und Länge
            for offset in range(1, 4):
                if i + offset < len(lines):
                    next_line = lines[i + offset]
                    if not kabeltyp and kabeltyp_rx.search(next_line):
                        kabeltyp = kabeltyp_rx.search(next_line).group()
                    if not laenge and laenge_rx.search(next_line):
                        laenge = laenge_rx.search(next_line).group()

            kabelinfos.append({
                "Kabelname": kabelnummer,
                "Kabeltyp": kabeltyp,
                "SOLL": laenge
            })

    return kabelinfos

    
@app.get("/process_kuep_page_ocr")
def process_kuep_page_ocr(file_id: str, page: int):
    temp_path = os.path.join("/tmp", f"{file_id}.pdf")

    if not os.path.exists(temp_path):
        raise HTTPException(status_code=404, detail="Datei nicht gefunden")

    try:
        kabelinfos = extract_kabel_with_ocr(temp_path, page_number=page)
        return JSONResponse(content=kabelinfos)
    except Exception as e:
        logging.error(f"OCR-Fehler auf Seite {page}: {e}")
        return JSONResponse(status_code=200, content=[])



# ----------------- Unverändert -------------------

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
            if beste_tabelle and beste_score >= 10:
                daten_ab_header = beste_tabelle[beste_header_zeile:]
                header = daten_ab_header[0]
                try:
                    df = pd.DataFrame(daten_ab_header[1:], columns=make_unique(header))
                    alle_daten.append(df)
                except Exception:
                    pass
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
                            if header_idx is None:
                                for idx, zelle in enumerate(spalte):
                                    header = match_header(zelle)
                                    if header:
                                        header_idx = idx
                                        header_name = zelle
                                        break
                            if header_idx is not None:
                                inhalt_nach_header.append(list(spalte[header_idx+1:]))
                                neue_header.append(header_name)
                            else:
                                inhalt_nach_header.append(list(spalte[1:]))
                                neue_header.append(spalte[0] or f"unknown_{len(neue_header)}")
                        daten_zeilen = list(zip(*inhalt_nach_header))
                        try:
                            df = pd.DataFrame(daten_zeilen, columns=make_unique(neue_header))
                            df = df.dropna(how='all')
                            df = df[df.notna().sum(axis=1) >= 6]
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
