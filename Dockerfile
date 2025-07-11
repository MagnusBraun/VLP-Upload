# Nutze offizielles Python-Image
FROM python:3.10-slim

# Installiere Tesseract
RUN apt-get update && apt-get install -y tesseract-ocr

# Setze Arbeitsverzeichnis
WORKDIR /app

# Kopiere Python requirements und installiere sie
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Kopiere dein Projekt (alles andere)
COPY . .

# Expose den Port, auf dem deine App läuft
EXPOSE 10000

# Startbefehl (Passe ggf. main:app an, falls dein File nicht main.py heißt)
CMD ["uvicorn", "main:app", "--host", "0.0.0.0", "--port", "10000"]
