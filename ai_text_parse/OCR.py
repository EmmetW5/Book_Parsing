

import pytesseract
from PIL import Image
from pdf2image import convert_from_path
import os

pytesseract.pytesseract.tesseract_cmd = r'D:\\TESSERACT_OCR\\tesseract.exe'

def ocr_core(image):
    try:
        text = pytesseract.image_to_string(image)
        return text
    except Exception as e:
        print(f"Error: {e}")
        return None

def write_to_file(text, filename):
    with open(filename, 'w') as file:
        file.write(text)

image = convert_from_path("images\\1989_us.pdf")
print(image)


input = "images\\1989_us.pdf"
output = "output.txt"
text = ocr_core(image)
write_to_file(text, output)