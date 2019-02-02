# VTU-Results-Scrapper V3
A python script to scrap VTU Results and store result in a workbook/spreadsheet or a Database



Dependencies and Libraries:-

1) Tessaract OCR:- [tesseract-ocr-w32-setup-v4.0.0.20181030.exe](https://digi.bib.uni-mannheim.de/tesseract/tesseract-ocr-w32-setup-v4.0.0.20181030.exe) (32 bit) or
[tesseract-ocr-w64-setup-v4.0.0.20181030.exe](https://digi.bib.uni-mannheim.de/tesseract/tesseract-ocr-w64-setup-v4.0.0.20181030.exe) (64 bit).

2) Python-tesseract(pytessaract):- pip install pytesseract

3) Pillow(PIL fork):-  pip install Pillow

4) Requests  

5) lxml- pip install lxml

6)XlsxWriter:- pip install XlsxWriter

Instructions:-

1)Set path to Tesseract in ocr.py

2)Update new token in payload of start.py in accordance to new token in VTU results page

3)Execute start.py
