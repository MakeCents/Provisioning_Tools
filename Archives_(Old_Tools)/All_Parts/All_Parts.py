import re
from browser import *
data = {}
l = []
def pline():
	print '============================'
#Build list
for i in filetoopen:
    l.append(i[:-1])
filetoopen.close()

def ppr(pr):
    c = 'y'
    pl = ''
    if len(pr)>1:
        pl = 's'
    print 'I have ' + str(len(pr)) + ' result' + pl
    co = 0
    while c == 'y':
        co +=1
        if co <= len(pr):
            print '(' + str(co) + ')', pr[co-1]
        else:
            return ch(pr)
        if co%12 == 0:
            c = str(raw_input("Continue?  (y/n)"))
            if c == '':
                c = 'y'
    return ch(pr)

#Choice from list
def ch(pr):
    pline()
    print 'Enter a number to make your selection, or enter n for no selection'
    try:
        choice = str(raw_input("Choice?  "))
        if choice == 'n' or choice == 'N' or int(choice) not in list(range(1,len(pr)+1)):
            return
        choice = int(choice)-1
        if choice < 0:
            print 'Invalid choice'
            ch(pr)
        else:
            return wtd(pr[int(choice)])
    except:
        print 'Invalid choice'
        ch(pr)
        
#What to do?
def wtd(cho):
    options = {1:'List AKAs',2:'Update AKAs',3:'Update description'}
    ansneed = {1:'no',2:'yes',3:'yes'}
    for i in range(len(options)):
        print '(' + str(i+1) + ')' + options[i+1]
    dec = str(raw_input('Do?  '))
    do = {1:"cho.getaka()",2:"cho.updateaka(str(raw_input('aka with? ')))",3:"cho.updatedesc(str(raw_input('desc with? ')))"}
    if dec == '1':
        print eval(do[int(dec)])
    else:
        eval(do[int(dec)])

######################
#### OBJECTS #########
######################

class Part(object):
    def __init__(self,number, cage, desc = '',aka = []):
        self.number = number
        self.cage = cage
        self.desc = desc
        self.aka = aka
        if (number,cage) in data:
            data[(number,cage)].number = number
            data[(number,cage)].cage = cage
            data[(number,cage)].desc = desc
        else:
            try:
                if (number.number, number.cage) in data:
                    data[(number.number, number.cage)].cage = cage
                    data[(number.number, number.cage)].desc = desc
            except:
                self.number = str(number)
                self.aka = []
                self.cage = cage
                if desc == '':
                    self.desc = str(raw_input('Description for ' + str(self.number) + '?  '))
                else:
                    self.desc = str(desc)
                for i in aka:
                    if i != "":  
                        try:
                            i.number
                            self.aka = aka[:]
                            break
                        except:
                            if (i,'NSN') in data:
                                self.aka = list(set(self.aka).union(set([data[(i,'NSN')]])))
                                data[(i,'NSN')].updateaka(self)
                            else:
                                tem = Part(i,'NSN',self.desc,[self])
                                data[(tem.number,'NSN')] = tem
                                self.aka = list(set(self.aka).union(set([tem])))
            data[(self.number,self.cage)] = self
            
    def __repr__(self):
        temp = self.getaka()
        return 'P/N: {0} * Desc: {1} * CAGE: {2} {3}'.format(self.number, self.desc, self.cage,temp)
    def updatedesc(self, desc):
        self.desc = desc
    def updateaka(self, aaka):
        try:
            if (aaka.number, aaka.cage) in data:
                    data[(aaka.number, aaka.cage)].aka = list(set(aaka.aka).union(set(self.aka)))[:]
        except:
            if (aaka,'NSN') in data:
                temp = data[(aaka,'NSN')]
            else:
                temp = Part(aaka,'NSN',self.desc,self.aka[:])
        if aaka not in self.aka:
            self.aka = list(set(self.aka).union(set([aaka])))
        if self not in self.aka:
            self.aka = list(set(self.aka).union(set([self])))
        self.aka = list(set(aaka.aka).union(set(self.aka)))
        ll = self.aka[:]
        if (self.number, self.cage) not in data:
            print (self.number, self.cage)
        for p in self.aka:
            now = ll[:]
            now.pop(now.index(p))
            try:
                data[(p.number,p.cage)].aka = now[:]
            except:
                data[(p.number,p.cage)] = p
                data[(p.number,p.cage)].aka = now[:]

    def getaka(self):
        res = '\n' + '   AKA:'
        for i in self.aka:
            if i != "":
                res+='\t{0} * {1} * {2}\n'.format(i.number, i.desc, i.cage)
        if res == '\n' + '   AKA:':
            return 'No AKA for ' + self.number
        else:
            return res[:-1]

######################
#### TOOLS ###########
######################

def addparts(pt):
    pass

def deleteparts(pt):
    pass

#Look in data for anything that matches what you are searching for
def lookup(se):
    if se == '':
        pass
    else:
        pr = [i for i in data.values() if set(se.lower().split()).issubset(re.findall('\w+',i.number.lower()))]
        pr += [i for i in data.values() if set(se.lower().split()).issubset(re.findall('\w+',i.desc.lower()))]
        pr = [i for i in data.values() if se.lower() in i.number.lower()] #searches for any match
        pr += [i for i in data.values() if se.lower() in i.desc.lower()]
        if pr != []:
            ppr(pr)
        else:
            print se + ' not found'


#search tool for searching whole words, or subsets, or anymatch
se = 'x'
for i in l:
    Part(i[3:i.find('</p>')].rstrip(),i[i.find('<c>')+3:i.find('</c>')].rstrip() ,i[i.find('<d>')+3:i.find('</d>')].rstrip(),i[i.find('<aka>')+5:i.find('</aka>')].rstrip().split(","))

while se != '': #what to search for
    options = {0:'Exit',1:'Look up parts',2:'Add parts',3:'Delete Parts'}
    for i in options: print '(' + str(i) + ') ' + options[i]
    do = {0:"print('Exiting')",1:"lookup(str(raw_input('Lookup what?  ')))",2:"addparts(str(raw_input('Add part? ')))",3:"cho.updatedesc(str(raw_input('Delete part? ')))"}
    try:
        pline()
        choice = int(raw_input('What would you like to do?  '))
        exec(do[choice])
        if choice == 0:
		break
    except:
        print('**Invalid selection**')
        pline()
        
####
##
##
##data['5305015990721'].updateaka(data['01944'])
##data['5305015990721'].updateaka(data['5315015174827'])
##print '01994'
##for i in data['01944'].aka:
##	print i
##print '5315015174827'
##for i in data['5315015174827'].aka:
##	print i
##print '5305015990721'
##for i in data['5305015990721'].aka:
##	print i
##print '91375A533'
##for i in data['91375A533'].aka:
##	print i
##print '98338A270'
##for i in data['98338A270'].aka:
##	print i
