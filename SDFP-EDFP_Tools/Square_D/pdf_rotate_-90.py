'''
Import list from current file. split the name, run the text tool for each
'''

import os

location = os.getcwd()

clear = lambda: os.system('cls')

# importing os module
import os
# what directory are we interested in
directory = os.getcwd() + "\source"
# getting the list of files
files = os.listdir(directory);
fif = []
for i in files:
    if i[-3:] == "pdf":
        fif.append(i)
PLISNs = []
for i in fif:
    
    a,b = i.split("_")

    from pyPdf import PdfFileWriter, PdfFileReader
    import StringIO
    from reportlab.pdfgen import canvas
    from reportlab.lib.pagesizes import letter
    from reportlab.lib.units import inch
    
    packet = StringIO.StringIO()
    # create a new PDF with Reportlab
    can = canvas.Canvas(packet, pagesize=letter)
    
    a = a.split(",")
    temp = a[:]
    start = -806
    col = 220
    first = 0
    can.rotate(-90)
    for o in a:
        if first  == 0:
            if len(temp) > 1:
                can.drawString(col, start, o+",")
            else:
                can.drawString(col, start, o)
        else:
            if len(temp) > 1:
                if first == 2:
                    can.drawString(col, start, o)
                else:
                    can.drawString(col, start, o+",")
            else:
                can.drawString(col, start, o)
        temp=temp[1:]
        #col+=320
        first +=1
        if col > 54:
            if len(temp) >0:
                start -= 15
                col = 120
                first = 0
        
        
    start -= 15
    
##    can.translate(inch,inch)
    # define a large font
##    can.setFont("Helvetica", 14)
    # choose some colors
##    User Guide Chapter 2 Graphics and Text with pdfgen
##    Page 11 c.setStrokeColorRGB(0.2,0.5,0.3)
##    can.setFillColorRGB(1,0,1)
    
    # draw a rectangle
    #x is negative up
    x = -5.06
    #y is to the right
    y = 4.65
    #can.rect(x*inch,y*inch,1.25*inch,.75*inch, fill=0)
    # make text go straight up
    
    # change color
    can.setFillColorRGB(0,0,0)
    # say hello (note after rotate the y coord needs to be negative!)
##    print inch
##    can.drawString(0.3*inch, -inch, "Hello World")
    # draw some lines
    can.setStrokeColorRGB(255,0,0)

    can.setFillColorRGB(255,0,0)
    can.line(x*inch-5,y*inch-5,x*inch+90,y*inch-5)
    can.line(x*inch+90,y*inch-5,x*inch+90,y*inch+58)
    
    can.line(x*inch-5,y*inch-5,x*inch-5,y*inch+58)
    can.line(x*inch-5,y*inch+58,x*inch+90,y*inch+58)
    

    can.setFillColorRGB(0,0,0)
    can.drawString(x*inch,y*inch, "FSC: 5925")
    can.drawString(x*inch,(y*inch + 15), "SCC: 00002")
    can.drawString(x*inch,(y*inch + 30), "PCCN: W9E215")
    can.drawString(x*inch,(y*inch + 45), "PLISN: " + b[:4])
   
   
    can.save()

    #move to the beginning of the StringIO buffer
    packet.seek(0)
    name = 'Dave'
    can.beginForm(name, lowerx=0, lowery=0, upperx=None, uppery=None)
    can.endForm()
    new_pdf = PdfFileReader(packet)
    
    # read your existing PDF
    fname = 'source\\' + i
    existing_pdf = PdfFileReader(file(fname, "rb"))
    output = PdfFileWriter()
    # add the "watermark" (which is the new pdf) on the existing page
    nump = existing_pdf.getNumPages()
    page = existing_pdf.getPage(0)
    for l in range(nump):
        output.addPage(existing_pdf.getPage(l))
    page.mergePage(new_pdf.getPage(0))
    # finally, write "output" to a real file
    outputStream = file(a[0]+"_"+b, "wb")
    output.write(outputStream)
    outputStream.close()
    print a[0]+"_"+b + " written", i
    if a[0] in PLISNs:
        print a[0]
    else:
        PLISNs.append(a[0])
