import pytesseract
from PIL import Image


def get_ocr(src):
    pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe" # Specify path to tesseract.exe from Tessaract installation
    image = Image.open(src)
    text = pytesseract.image_to_string(image)
    return text
