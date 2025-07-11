#!/usr/bin/env bash
# Systempakete installieren
apt-get update
apt-get install -y tesseract-ocr
# Dann wie üblich Python requirements installieren
pip install -r requirements.txt

