def openfile(x):
    from urllib import urlopen
    import re
##    x = "https://www.dlis.dla.mil/webflis/pub/pub_search.aspx?niin=3040013055697&newpage=1"
    f = urlopen(x)
    return f
    pass
import re

class Database(object):
    def __init__(self, NSN, pns, cages):
        self.NSN = NSN
        self.parts = []
        try:
            for i in range(len(pns)):
                self.parts.append((pns[i],cages[i]))
            
        except:
            print NSN, pns, cages
        DataBaseNSN[NSN] = (self)
        
nfile = open('NSNs.txt','r')
l = []
for i in nfile:
    if i[:-1] not in l:
        l.append(i[:-1])
nfile.close()	
DataBaseNSN = {}
for nsn in l:
    f = openfile("https://www.dlis.dla.mil/webflis/pub/pub_search.aspx?niin={0}".format(nsn))
    p=', '.join([str(x) for x in [i for i in f]])

    parts = [i[2:34].strip() for i in re.findall(r'(">.{32}</font)',p)]
    cages = [i[7:12].strip() for i in re.findall(r'(blank">.{5}</a>)',p)]
    Database(nsn,parts, cages)

for i in DataBaseNSN:
    print  DataBaseNSN[i].parts
    print DataBaseNSN[i].NSN



