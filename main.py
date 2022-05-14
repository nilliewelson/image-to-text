import os
import cv2
from docx import Document
from docx.text.paragraph import Paragraph
from pytesseract import image_to_string


doc = Document(os.getcwd() + "/cases.docx")
images = os.listdir(os.getcwd())
images.sort()
print(images)


def valid_xml_char_ordinal(c):
    codepoint = ord(c)
    # conditions ordered by presumed frequency
    return (
        0x20 <= codepoint <= 0xD7FF or
        codepoint in (0x9, 0xA, 0xD) or
        0xE000 <= codepoint <= 0xFFFD or
        0x10000 <= codepoint <= 0x10FFFF
        )


for image in images:
	if "case" in image and "study" in image:
		img = cv2.imread(image)
		(h, w) = img.shape[:2]
		img = cv2.resize(img, (w*6, h*6))
		gry = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
		thr = cv2.threshold(gry, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)[1]
		txt = image_to_string(thr)
		clean_txt = ''.join(c for c in txt if valid_xml_char_ordinal(c))
		doc.add_paragraph(clean_txt)


doc.save("cases.docx")