from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import os
import platform

def convert_txt_to_pdf(input_path: str, output_filename: str = None) -> str:
    if output_filename:
        output_path = os.path.join(os.path.dirname(input_path), output_filename)
    else:
        output_path = os.path.splitext(input_path)[0] + ".pdf"
    
    # Register Unicode fonts
    try:
        if platform.system() == "Windows":
            pdfmetrics.registerFont(TTFont('UniFont', 'C:/Windows/Fonts/arial.ttf'))
        else:
            pdfmetrics.registerFont(TTFont('UniFont', '/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf'))
        font_name = 'UniFont'
    except:
        font_name = 'Helvetica'
    
    c = canvas.Canvas(output_path, pagesize=letter)
    c.setFont(font_name, 11)
    
    with open(input_path, "r", encoding="utf-8", errors="ignore") as f:
        text = f.readlines()

    x, y = 50, 750
    for line in text:
        if y < 50:
            c.showPage()
            c.setFont(font_name, 11)
            y = 750
        try:
            c.drawString(x, y, line.strip())
        except:
            # If drawing fails, try to encode properly
            try:
                c.drawString(x, y, line.strip().encode('utf-8').decode('utf-8'))
            except:
                c.drawString(x, y, "Unable to render line")
        y -= 15
    c.save()
    return output_path
