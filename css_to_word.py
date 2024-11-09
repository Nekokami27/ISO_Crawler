from docx.shared import Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx import Document

def apply_css_to_word(doc, text, css_styles):
    # Create a new paragraph in the document
    paragraph = doc.add_paragraph()
    run = paragraph.add_run(text)

    # font size
    if 'font-size' in css_styles:
        font_size = float(css_styles['font-size'].replace("px", ""))
        font_size = int(font_size)
        run.font.size = Pt(font_size)
        #print(f"Font size: {font_size}")
    # font weight (bold)
    if 'font-weight' in css_styles and css_styles['font-weight'] >= '700':
        run.bold = True

    # text alignment
    if 'text-align' in css_styles:
        alignment = css_styles['text-align']
        if alignment == 'center':
            paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        elif alignment == 'right':
            paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
        elif alignment == 'left':
            paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
