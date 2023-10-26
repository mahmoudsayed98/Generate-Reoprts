import pyodbc
import tabulate 
from reportlab.pdfgen import canvas
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.pagesizes import letter
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer,PageBreak,Table, TableStyle
from reportlab.lib.units import inch
from reportlab.lib.styles import ParagraphStyle
from reportlab.pdfbase import pdfmetrics
import reportlab
from reportlab.pdfbase.ttfonts import TTFont
import arabic_reshaper
from bidi.algorithm import get_display
import PyPDF2
import fitz

# Connect to the SQL database
conn = pyodbc.connect('Driver={SQL Server};'
                      'Server=.;'
                      'Database=;'
                      'Trusted_Connection=yes;')
# Retrieve the data from the SQL database
cur = conn.cursor()
cur.execute("SELECT name , Code,reference FROM Fac")
rows = cur.fetchall()

pdfmetrics.registerFont(TTFont('Arabic', '29ltbukraregular.ttf'))

seatNumber = get_display(arabic_reshaper.reshape(''))
name = get_display(arabic_reshaper.reshape(''))
faculty = get_display(arabic_reshaper.reshape(''))
# Create the report
data = [[faculty, name ,seatNumber]]

for row in rows:
    data.append([
        get_display(arabic_reshaper.reshape(str(row[0]))),
        get_display(arabic_reshaper.reshape(str(row[1]))),
        get_display(arabic_reshaper.reshape(str(row[2])))])
    

# Create a header function
def add_header(canvas, doc):
    canvas.saveState()
    canvas.setFont("Arabic", 16)  # Set the font and size for the header
    canvas.setFillGray(0.5, 0.5)
    header_text = " Write Water mark here  "
    reshaped_header = arabic_reshaper.reshape(header_text)
    display_header = get_display(reshaped_header)
    canvas.drawString(150, 750, display_header)
    canvas.restoreState()


# Initialize the PDF document
doc = SimpleDocTemplate('report.pdf', pagesize=letter)

# Create the table and apply the style
table = Table(data, colWidths=[4*inch, 3*inch, 1*inch])

# Apply table style
table.setStyle(TableStyle([('BACKGROUND', (0, 0), (-1, 0), '#CCCCCC'),
                           ('TEXTCOLOR', (0, 0), (-1, 0), '#000000'),
                           ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                           ('FONTNAME', (0, 0), (-1, -1), 'Arabic'),
                           ('FONTSIZE', (0,0), (-1,0), 11),
                           ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                           ('BACKGROUND', (0, 1), (-1, -1), '#EEEEEE'),
                           ('GRID', (0, 0), (-1, -1), 1, '#888888'),
                           ]))

# Build the PDF document with the table
storys = [table, Spacer(1, 8)]  # Add a spacer after the table

# Create a watermark function
def add_watermark(canvas, doc):
    canvas.saveState()
    canvas.setFont("Arabic", 20)
    canvas.setFillGray(0.5, 0.5)
    canvas.rotate(45)
    watermark_text = "Write Water mark here"
    reshaped_text = arabic_reshaper.reshape(watermark_text)
    display_text = get_display(reshaped_text)
    watermark_positions = [(500, 500),(400, 400),(300, 300), (200, 200),(600, 200),(100, 100),(600, 100),(500, 20),(30, 20), (600, -70), (60, -70), (150, -150), (80, -250), (90, -350)]

    for x, y in watermark_positions:
        canvas.drawString(x, y, display_text)
    canvas.restoreState()

doc.build(storys, onFirstPage=add_header, onLaterPages=add_watermark)
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

cur.close()
conn.close()

# # Encrypt the PDF using PyPDF2
# with open('report.pdf', 'rb') as pdf_file:
#     pdf_reader = PyPDF2.PdfReader(pdf_file)
#     pdf_writer = PyPDF2.PdfWriter()

#     for page_num in range(len(pdf_reader.pages)):
#         pdf_writer.add_page(pdf_reader.pages[page_num])

#     # Set encryption options
#     pdf_writer.encrypt("user_password", "owner_password", permissions_flag= fitz.PDF_PERM_PRINT)

#     with open('encrypted_report.pdf', 'wb') as encrypted_pdf:
#         pdf_writer.write(encrypted_pdf)
# # Close the database connection