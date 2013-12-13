from urllib import urlopen
def openfile(x):
    import re
    f = urlopen(x)
    return f
    pass
    
def CAGE(search):
    return 'https://www.logisticsinformationservice.dla.mil/bincs/details.aspx?CAGE={0}'.format(search)

print('Please enter a 5 digit cage code?  ')
print('For multiple look up seperate CAGE codes with a comma')
x = str((raw_input('Cage code:  ')))
x = x.split(',')
x = [x.strip() for x in x]
#x = [] #this is where you manual make it work

print 'Finding information for ' + str(x)

l = ['ParentCAGEData">','CompNameData">','lblCAGEData">','AddressData">',
	'POBoxData">','CityData">','ZipData">','CaoAdpData">','CountyData">',
	'PhoneData">','FaxData">','EstablishedData">','UpdatedData">','PocData">','StateData">']

class CAGECODE(object):
	def __init__(self):
		self.ParentCAGE=""
		self.CompName=""
		self.lblCAGE=""
		self.Address=""
		self.POBox=""
		self.City=""
		self.Zip=""
		self.CaoAdp=""
		self.County=""
		self.Phone=""
		self.Fax=""
		self.Established=""
		self.Updated=""
		self.Poc=""
		
def findit(x,i):
	try:
		return i[i.index(x)+len(x):i.index("</")]
	except:
		return None
def loaditup(CC,f):
	for i in f:
		for o in l:
			if findit(o,i) != None:
				setattr(CC, o[:-6], findit(o,i))
datalist = []		
for c in x:
	if c != '':
		f = openfile(CAGE(str(c)))
		CC = CAGECODE()	
		loaditup(CC,f)
		datalist.append(CC)
		print('=================')
		print CC.lblCAGE, CC.CompName, CC.Address, CC.City, CC.Zip
		







	
