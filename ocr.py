import pytesseract
from PIL import Image


def get_ocr(src):
    pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe"
    image = Image.open(src)
    text = pytesseract.image_to_string(image)
    return text
