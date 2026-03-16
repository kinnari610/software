from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter

c = canvas.Canvas('test_images.pdf', pagesize=letter)
# also use reportlab.lib.utils.ImageReader
from reportlab.lib.utils import ImageReader

logo = ImageReader('logo.png')
width, height = logo.getSize()
# place it at (100, 700)
c.drawImage(logo, 100, 700, width=width*0.5, height=height*0.5, preserveAspectRatio=True, mask='auto')

c.showPage()
c.save()
print('wrote test_images.pdf')
