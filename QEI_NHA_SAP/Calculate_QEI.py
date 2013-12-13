import sys
sys.path.insert(0,'lib')

#This would be done first, before this module.
import LOAD
import NHA
import PLISNS
import SAP
import QTY
import QEI

#Load will return a list from file 
lines = LOAD.Load()
     
def cal_QEI(line):
    QEI =  1
    plisn = line[6:11].strip()
    while NHA.NHA[plisn] != '':
        QEI*=QTY.QTY[plisn]
        plisn = NHA.NHA[plisn]
    return QEI

def fixline(line):
    plisn = line[6:11].strip()
    if line[-4:-1] == '01C':
        if SAP.SAP[plisn] == "":
            qei = str(QEI.QEI[cpn])
            qei = "0"*(5-len(qei))+qei
        else:
            if plisn in NHA.AllNHA:
                qei = 'REFX '
            else:
                qei = 'REF  '
        line = line[:25] + qei + line[30:]
    return line

for line in lines:
    if len(line)>80 and line[-2] != " ":
        #This keeps the PLISNs in order for later
        if line[-4:-1] == '01A':
            plisn = line[6:11].strip()
            cpn = line[13:50].strip()
            PLISNS.add(plisn)
            SAP.pp(cpn,plisn)
        elif line[-4:-1] == '01C':
            nha = line[12:17].strip()
            sap = line[59:64].strip()
            qty = int(line[21:25])
            SAP.add(plisn, sap)
            NHA.add(plisn, nha)
            QTY.add(plisn, qty)
            #calculate qei for this part
            qei = cal_QEI(line)
            QEI.add(cpn, qei)
               
newlines = []
for line in lines:
    global cpn
    if line[-4:-1] == '01A':
        cpn = line[13:50].strip()
    newlines.append(fixline(line))


##Order of PLISNS
#PLISNS.PLISNS

##Writes each item in list without '\n' after
LOAD.Write(newlines)

