
Records = []
lopn = {}
class Crecords(object):
        def __init__(self,card):
                self.name = card[6:11]
                self.pn = card[18:50]

def Load(FileName = 'W9E215.036'):
        if FileName == '':
            FileName = 'W9E215.036'
        try:
                FileO = open(FileName)
        except:
            try:
                    FileO = open(FileName+".036")
                    FileName = FileName+".036"
            except:
                    print '***File "' + str(FileName) + '" not found. Please try again.***'
                    print '***Remember to include the extenstion.***'
                    print '=============================================================='
                    FileName = str(raw_input('What is the  file name?  '))
                    FileO, FileName = Load(FileName)
        return FileO, FileName


def QEI(FileName):
        for i in FileName:
                if i[-3:-1] == '1A':
                        Records.append(Crecords(i))
                elif i[-3:-1] == '1C':
                        for e in Records:
                                if e.name == i[6:11]:
                                        e.qpa = i[21:25]
                                        e.NHA = i[12:17]
                                        e.SaP = i[59:64]
                                        break
        for r in Records:
                n = r.NHA
                try:
                        q = int(r.qpa)
                        while n != '     ' or n != '     ':
                                for o in Records:
                                        
                                        if o.name == n:
                                                q*= int(o.qpa)
                                                n = o.NHA
                                                break
                                        
                        if r.pn not in lopn:
                                lopn[r.pn] = q                        
                        else:
                                lopn[r.pn]+= q
                except:
                        q = r.qpa
                                
        for r in Records:
                try:
                        q = (5 - len(str(lopn[r.pn]))) * '0'
                        r.qei = q + str(lopn[r.pn])
                except:
                        r.qei = r.qpa + " "
                
def writefile(FN, ans = 'n'):
        #'test2.txt'
        F, FN = Load(FN)
        f = open(str('New file ' + str(FN)), 'w')
        line =''
        for i in F:
                if i[-3:-1] == '1C':
                        for r in Records:
                                if r.name == i[6:11]:
                                        if ans != 'y':
                                                line = i[:25] + r.qei + i[30:]
                                        else:
                                                if r.SaP == "     ":
                                                        line = i[:25] + r.qei + i[30:]
                                                else:
                                                        line = i[:25] + 'REF  ' + i[30:]
                                                
                                        line = line[:59] + r.SaP + line[64:]
                                        break
                else:
                        line = i
                f.write(line)
        
        f.close()
sapl = []
def SAP(FileName):
        for i in FileName:
                if i[-3:-1] == '1A':
                        if i[18:50] not in sapl:
                                sapl.append(i[18:50])
                                name = '     '
                                for r in Records:
                                        if r.name == i[6:11]:                               
                                                r.SaP = name
                        else:
                                for e in Records:
                                        if e.pn == i[18:50]:
                                                name = e.name
                                                break
                                for r in Records:
                                        if r.name == i[6:11]:                               
                                                r.SaP = name
                                                break
                        
        
                
        

##F = Load('test2.txt')
print '=============================================================='
File = str(raw_input('What is the file name?  '))
print '=============================================================='
F, FN = Load(File)
print '=============================================================='
r = str(raw_input('Would you like to replace repeated part quantities with REF? y/n '))
print 'Calulating Quatities Per End Item...'

QEI(F)

F.close()
F, FN = Load(FN)
print 'Calulating Same as PLISN...'
SAP(F)
F.close()

print 'Writing file as "New file ' + str(FN) + '"'

writefile(FN,r)
print 'Finished'

