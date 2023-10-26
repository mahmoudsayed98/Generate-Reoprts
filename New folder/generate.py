import pyodbc 
import tabulate 
from reportlab.pdfgen import canvas
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import inch, cm
from reportlab.lib.pagesizes import letter
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer,PageBreak,Image,BaseDocTemplate,PageTemplate,Table, TableStyle
from reportlab.lib.styles import ParagraphStyle
from reportlab.pdfbase import pdfmetrics
import reportlab
from reportlab.pdfbase.ttfonts import TTFont
import arabic_reshaper
from bidi.algorithm import get_display
from reportlab.platypus.frames import Frame
import fitz

"""
from reportlab.platypus.flowables import KeepInFrame
from reportlab.rl_config import defaultPageSize

(MAXWIDTH, MAXHEIGHT) = defaultPageSize
"""

# Connect to the SQL database
conn = pyodbc.connect('Driver={SQL Server};'
                      'Server=.;'
                      'Database=TansikRules_2;'
                      'Trusted_Connection=yes;')
# Retrieve the data from the SQL database
cur = conn.cursor()
cur.execute("SELECT referenceCode, governorateCode, name FROM Faculties")
#cur.execute("SELECT SeatNumber ,studentname,facultyname from  DEL_RES1 ORDER BY s")
rows = cur.fetchall()

pdfmetrics.registerFont(TTFont('Arabic', '29ltbukraregular.ttf'))
#init the style sheet
styles = getSampleStyleSheet()
arabic_text_style = ParagraphStyle(
    'none',#'border', # border on
    parent = styles['Normal'] , # Normal is a defaul style  in  getSampleStyleSheet
    borderColor= '#333333',
    borderWidth =  0,
    borderPadding  =  2,
    fontName="Arabic" #previously we named our custom font "Arabic"
)
storys = []

seatNumber = get_display(arabic_reshaper.reshape('رقم الجلــوس'))
name = get_display(arabic_reshaper.reshape('الإســــــــــم'))
faculty = get_display(arabic_reshaper.reshape('الكلــــــية'))

# Create the report
data = [[faculty, name ,seatNumber]]

for row in rows:
    data.append([get_display(arabic_reshaper.reshape(str(row[2]))),
                 get_display(arabic_reshaper.reshape(str(row[1]))),
                 get_display(arabic_reshaper.reshape(str(row[0])))])    

# Create the table and apply the style
table = Table(data, colWidths=[3.5*inch, 2*inch, 1*inch])

# Apply table style
table.setStyle(TableStyle([('TEXTCOLOR', (0, 0), (-1, 0), '#000000'),
                           ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                           ('FONTNAME', (0, 0), (-1, -1), 'Arabic'),
                           ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                           ('TOPPADDING', (0, 0), (-1, -1), 25),
                           ('GRID', (0, 0), (-1, -1), 1, '#888888'),
                           ]))
                           
# Build the PDF document with the table
storys = [table, Spacer(2, 8)]  # Add a spacer after the table
#close the database connection
cur.close()
conn.close()

def AllPageSetup(canvas, doc):

    canvas.saveState()
    canvas.setFont('Arabic', 12)
    #canvas.drawImage('logo2.jpg', 0,0, width=None,height=None,mask=None)
    #header
    canvas.drawString(0.5 * inch, 10.5 * inch, get_display(arabic_reshaper.reshape(u"نتيجة الثانوية العامة")))
    canvas.drawRightString(8 * inch, 10.5 * inch, get_display(arabic_reshaper.reshape(u"وزارة التعليم العالي")))
    canvas.drawRightString(8 * inch, 10.25 * inch, get_display(arabic_reshaper.reshape(u"مكتب تنسيق القبول بالكليات والمعاهد")))

    #footers
    canvas.drawString(0.5 * inch, 0.5 * inch, get_display(arabic_reshaper.reshape(u"تنسيق المرحلة الاولى")))
    canvas.drawRightString(8 * inch, 0.5 * inch, 'Page %d' % (doc.page))

    canvas.setFont("Arabic", 20)
    canvas.setFillGray(0.2)
    canvas.rotate(45)
    watermark_text = "وزارة التعليم العالي - مكتب تنسيق القبول بالجامعات والمعاهد"
    reshaped_text = arabic_reshaper.reshape(watermark_text)
    display_text = get_display(reshaped_text)
    watermark_positions = [(500, 500),(400, 400),(300, 300), (200, 200),(600, 200),(100, 100),(600, 100),(500, 20),(30, 20), (600, -70), (60, -70), (150, -150), (80, -250), (90, -350)]

    for x, y in watermark_positions:
        canvas.drawString(x, y, display_text)
    
    canvas.restoreState()

doc = BaseDocTemplate('report.pdf',pagesize = letter)
page_template = PageTemplate(id="fund_notes",frames=[Frame(2.5*cm, 2.5*cm, 15*cm, 25*cm, id='F1')], onPage=AllPageSetup, pagesize=letter)
doc.addPageTemplates(page_template)
#Now, every page will have headers, footers, and a watermark

## add the storys array to the pdf document
doc.build(storys)

#Convert pdf to pdf of images
pdf_file = "report.pdf"
pdf_document = fitz.open(pdf_file)
new_pdf_document = fitz.open()

for page_num in range(pdf_document.page_count):
    page = pdf_document[page_num]

    pixmap = page.get_pixmap(matrix=fitz.Matrix(150/72, 150/72))  # Adjust resolution
    new_page = new_pdf_document.new_page(width=pixmap.width, height=pixmap.height)
    new_page.insert_image(new_page.rect, pixmap=pixmap)

new_pdf_filename = "report_images.pdf"
new_pdf_document.save(new_pdf_filename, deflate = True) #deflate to compress
new_pdf_document.close()

pdf_document.close()

print("PDF with images saved as:", new_pdf_filename)