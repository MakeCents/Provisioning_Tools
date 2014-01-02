import sys
import os

sys.path.insert(0,'lib')

l = [x for x in os.listdir(os.getcwd()) if x[-4:] == '.036']
l.insert(0, "Not here")

#This would be done first, before this module.
import LOAD
import NHA
import PLISNS
import SAP
import QTY
import QEI
from LINE import *
import DIFFERENCE

Cal_NHA = []

#Load will return a list from file 
#load = LOAD.Load()
answers = LOAD.loadNames(l)
lines = answers[0]
originalFile = answers[1]


def cal_QEI(line):
    '''C card line from 036 report
        Takes a C card from the 036 report, 
        extracts the Qty and calculates the QEI
        and adds it to the over all QEI of that cpn
    '''
    QEI =  1
    plisn = line[6:11].strip()
    while NHA.NHA[plisn] != '':
        QEI*=QTY.QTY[plisn]
        plisn = NHA.NHA[plisn]
    return QEI

def Ind_NHA(ind, bool=True):
    global Cal_NHA
    last = NHA.IND[Cal_NHA[-1]]
    if last >= ind: 
        Cal_NHA.pop()
        return Ind_NHA(ind, False)
    else:
        nha = Cal_NHA[-1]
        return (nha + (" " * (5-len(nha))), True)

def fixline(line):
    '''line from the 036 report
        Takes any line form the 036 report and 
        decides what to do with it due to calculations
    '''
    plisn = line[6:11].strip()
    if line[-4:-1] == '01C':
        if SAP.SAP[plisn] == "":
            qei = str(QEI.QEI[CPN])
            qei = "0"*(5-len(qei))+qei
        else:
            if plisn in NHA.AllNHA:
                qei = 'REFX '
            else:
                qei = 'REF  '
        ind = NHA.IND[plisn]
        if ind == 'A' or ind == None:
            nha = ("     ", True)
        else:
            nha = Ind_NHA(ind)
        if nha[1]:
            addCal_NHA(plisn)
        line = line[:12] + nha[0] + line[17:25] + qei + line[30:59] +  SAP.get(CPN, plisn) + line[64:]
    return line
def addCal_NHA(plisn):
    global Cal_NHA
    if plisn == None:
        pass
    elif Cal_NHA==[]:
        Cal_NHA.append(plisn)
    elif plisn != Cal_NHA[-1]:
        Cal_NHA.append(plisn)
        
for line in lines:
    if len(line)>80 and line[-2] != " ":
        #This keeps the PLISNs in order for later
        plisn = line[6:11].strip()
        if line[-4:-1] == '01A':
            cpn = line[13:50].strip()
            PLISNS.add(plisn)
            SAP.pp(cpn,plisn)
            NHA.ind(line[12], plisn)
        elif line[-4:-1] == '01C':
            #nha = line[12:17].strip()
            sap = line[59:64].strip()
            if line[21:25].strip() != "":
                qty = int(line[21:25])
            else:
                qty = 0
            SAP.add(plisn, sap)
            QTY.add(plisn, qty)
            #fix NHA
            ind = NHA.IND[plisn]
            if Cal_NHA == []:
                addCal_NHA(plisn)
            if ind == 'A' or ind == None:
                nha = ("     ", True)
            else:
                nha = Ind_NHA(ind)
            if nha[1]:
                addCal_NHA(plisn)
            NHA.add(plisn, nha[0].strip())
            #calculate qei for this part
            qei = cal_QEI(line)
            QEI.add(cpn, qei)
        
               

CPN = ""
newlines = []
for line in lines:
    CPN = get_cpn(line, CPN)
    newlines.append(fixline(line))

##Order of PLISNS
#PLISNS.PLISNS

##Writes each item in list without '\n' after
newFile = LOAD.Write(newlines)

DIFFERENCE.Difference(originalFile, newFile)


#shows that it is done
clear = lambda: os.system('cls')
indent = '\t' +'\t'
def done():
    clear()
    print
    print
    print indent + 'DDDD       DD     DD    D   DDDDDD'
    print indent + 'DD DDD   DD  DD   DDDD DD   DD'
    print indent + 'DD DDD   DD  DD   DD DDDD   DDDD'
    print indent + 'DD DDD   DD  DD   DD  DDD   DD'
    print indent + 'DDDD       DD     DD   DD   DDDDDD'
    print
    raw_input(indent + '\t' +   'CLOSE THIS WINDOW')
    print
done()
