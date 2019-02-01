import pytesseract
from PIL import Image
"""
                                            Created by ABHISHEK KOUSHIK B N 
                                                        AND
                                                       AKASH R
                                            on 01/02/2019
 """


def get_ocr(src):
    pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe" # Specify path to tesseract.exe from Tessaract installation
    image = Image.open(src)
    text = pytesseract.image_to_string(image)
    return text
