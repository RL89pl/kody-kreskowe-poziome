
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4, ZZ, landscape
from reportlab.lib.units import mm
from reportlab.graphics.barcode import code39
import xlrd
from barcode import generate

plik = xlrd.open_workbook("numery.xls")
strona = plik.sheet_by_index(0)
total_rows = strona.nrows
c=canvas.Canvas("barcode.pdf",pagesize=landscape(ZZ))
licznik = total_rows // 16

i=0
for d in range(licznik-1):
    c.setFont("Helvetica", 25)
    

    code = strona.cell(i,0).value
    barcode=code39.Standard39(code,checksum=0, barWidth=0.35*mm,barHeight=30*mm)
    barcode.drawOn(c,4*mm,41*mm)
    c.drawString(14*mm,30*mm,code)
    i += 1
    code = strona.cell(i,0).value
    barcode=code39.Standard39(code,checksum=0, barWidth=0.35*mm,barHeight=30*mm)
    c.drawString(79*mm,35*mm,code)
    barcode.drawOn(c,69*mm,4*mm)
    i += 1
    code = strona.cell(i,0).value
    barcode=code39.Standard39(code,checksum=0, barWidth=0.35*mm,barHeight=30*mm)
    barcode.drawOn(c,144*mm,41*mm)
    c.drawString(154*mm,30*mm,code)
    i += 1
    code = strona.cell(i,0).value
    barcode=code39.Standard39(code,checksum=0, barWidth=0.35*mm,barHeight=30*mm)
    c.drawString(220*mm,35*mm,code)
    barcode.drawOn(c,210*mm,4*mm)
    i += 1
    # code = strona.cell(i,0).value
    # barcode=code39.Standard39(code,checksum=0, barWidth=0.35*mm,barHeight=30*mm)
    # c.drawString(190*mm,35*mm,code)
    # barcode.drawOn(c,184*mm,41*mm)
    # i += 1
    code = strona.cell(i,0).value
    barcode=code39.Standard39(code,checksum=0, barWidth=0.35*mm,barHeight=30*mm)
    c.showPage()


#####
#####
# now create the actual PDF
c.save()