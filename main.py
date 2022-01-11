# pip install python-docx
# pip install pillow

from docx import Document
from docx.shared import Inches
from PIL import Image
from os import listdir, path
from sys import argv
from docx.enum.text import WD_ALIGN_PARAGRAPH

directory = '.'
if len(argv) > 1:
    directory = argv[1]

document = Document()
image_max_width = Inches(6.1)
image_max_height = Inches(7.6)
# image_max_height = Inches(8.6)
breakpoint_aspect_ratio = image_max_width / image_max_height

files = listdir(directory)
sectionIdx = 0
items = []

for item in files:
    parts = item.split('.')
    if parts[len(parts)-1] != 'jpg' and parts[len(parts)-1] != 'png':
        continue
        
    items.append(item);

for item in items:
    
    section = document.sections[sectionIdx]
    section.header.is_linked_to_previous = False
    section.header.paragraphs[0].text = "\t\tGlobale ID"
    section.footer.is_linked_to_previous = False
    section.footer.paragraphs[0].text = "Lokale ID\t" + item + "\tSeite " + str(sectionIdx + 1)
    image_path = path.join(directory, item)
    img = Image.open(image_path)
    image_aspect_ratio = img.width / img.height
    if image_aspect_ratio > breakpoint_aspect_ratio:
        document.add_picture(image_path, width=image_max_width)
    else:
        image = document.add_picture(image_path, height=image_max_height)
        document.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
    paragraph = document.add_paragraph(item)
    paragraph.paragraph_format.space_before = 0
    paragraph.paragraph_format.space_after = 0
    for i in range(0, 4):
        paragraph = document.add_paragraph()
        paragraph.paragraph_format.space_before = 0
        paragraph.paragraph_format.space_after = 0
        
    if item != items[len(items)-1]:
        document.add_page_break()
        document.add_section()
        sectionIdx += 1

document.save('output.docx')
