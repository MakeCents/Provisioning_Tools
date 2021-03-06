'''
Import list from current file. split the name, run the text tool for each
'''

import os

location = os.getcwd()

clear = lambda: os.system('cls')

# importing os module
import os
# what directory are we interested in
directory = os.getcwd() + "\McMaster-Carr_Source_files"
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

    packet = StringIO.StringIO()
    # create a new PDF with Reportlab
    can = canvas.Canvas(packet, pagesize=letter)
    
    a = a.split(",")
    temp = a[:]
    start = 760
    col = 450
    first = 0
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
        col+=42
        first +=1
        if col > 546:
            if len(temp) >0:
                start -= 15
                col = 450
                first = 0
        
        
    start -= 15
    cage = "55910"
    can.drawString(450, start, "ECPVG2")
    can.drawString(450, start - 15, "CAGE: " + cage)

    
    can.save()

    #move to the beginning of the StringIO buffer
    packet.seek(0)
    
    new_pdf = PdfFileReader(packet)
    
    # read your existing PDF
    fname = 'McMaster-Carr_Source_files\\' + i
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
    
