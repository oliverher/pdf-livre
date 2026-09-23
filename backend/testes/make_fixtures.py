# Gera dois decks de teste: um nativo e um "achatado" (slide = imagem inteira)
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from PIL import Image, ImageDraw
import io

# --- deck nativo ---
prs = Presentation()
prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
s = prs.slides.add_slide(prs.slide_layouts[6])
tb = s.shapes.add_textbox(Inches(0.8), Inches(0.7), Inches(11), Inches(1.2))
tb.text_frame.text = "Transformacao de Servicos Publicos"
tb.text_frame.paragraphs[0].runs[0].font.size = Pt(40)
tb.text_frame.paragraphs[0].runs[0].font.color.rgb = RGBColor(0x1E,0x27,0x61)
body = s.shapes.add_textbox(Inches(0.8), Inches(2.2), Inches(6), Inches(3))
body.text_frame.text = "Indicadores de jornada do cidadao e metricas de satisfacao ao longo do ciclo."
from pptx.enum.shapes import MSO_SHAPE
sp = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(7.5), Inches(2.2), Inches(4.5), Inches(2.5))
sp.fill.solid(); sp.fill.fore_color.rgb = RGBColor(0x02,0x80,0x90)
s2 = prs.slides.add_slide(prs.slide_layouts[6])
t2 = s2.shapes.add_textbox(Inches(0.8), Inches(0.7), Inches(11), Inches(1))
t2.text_frame.text = "Resultados"
tbl = s2.shapes.add_table(3, 3, Inches(0.8), Inches(2), Inches(8), Inches(2)).table
tbl.cell(0,0).text = "Etapa"; tbl.cell(0,1).text = "NPS"; tbl.cell(0,2).text = "Volume"
prs.save("fix_nativo.pptx")

# --- deck achatado ---
img = Image.new("RGB", (1920, 1080), (30, 39, 97))
d = ImageDraw.Draw(img)
d.rectangle([120, 120, 1800, 300], fill=(2, 128, 144))
d.text((160, 400), "Slide exportado como imagem", fill=(255, 255, 255))
buf = io.BytesIO(); img.save(buf, format="PNG"); buf.seek(0)

prs2 = Presentation()
prs2.slide_width, prs2.slide_height = Inches(13.333), Inches(7.5)
for _ in range(3):
    sl = prs2.slides.add_slide(prs2.slide_layouts[6])
    buf.seek(0)
    sl.shapes.add_picture(buf, 0, 0, Inches(13.333), Inches(7.5))
prs2.save("fix_achatado.pptx")

# --- deck misto ---
prs3 = Presentation()
prs3.slide_width, prs3.slide_height = Inches(13.333), Inches(7.5)
sl = prs3.slides.add_slide(prs3.slide_layouts[6])
tt = sl.shapes.add_textbox(Inches(0.8), Inches(0.6), Inches(6), Inches(1))
tt.text_frame.text = "Diagrama do processo (arte importada)"
buf.seek(0)
sl.shapes.add_picture(buf, Inches(0.8), Inches(2), Inches(8), Inches(4.5))
prs3.save("fix_misto.pptx")
print("fixtures ok")
