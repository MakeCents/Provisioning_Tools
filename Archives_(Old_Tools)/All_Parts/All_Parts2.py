from browser import *
#Create a dicitionary of parts
parts = {}
l = []

class Parts(object):
	def __init__(self, number, cage, description, aka = []):
		self.number  = number
		self.cage = cage
		self.aka = aka
		self.description = description
		parts[(self.number, self.cage)] = self
		
	def update(self, att, ans):
		'''att = number, cage, description, and ans = update to what?
		'''
		setattr(self, att, ans)
		
for i in filetoopen:
	#parts[(i[3:i.find('</p>')].rstrip(),i[i.find('<c>')+3:i.find('</c>')].rstrip())]= \
	NUM = i[3:i.find('</p>')].rstrip()
	CAG = i[i.find('<c>')+3:i.find('</c>')].rstrip()
	DES = i[i.find('<d>')+3:i.find('</d>')].rstrip()
	AKA = i[i.find('<aka>')+5:i.find('</aka>')].rstrip().split(",")
	Parts(NUM,CAG,DES,AKA)
	if AKA != ['']:
		
		if (AKA[0],'NSN') not in parts:
			Parts(AKA[0],'NSN',DES,[NUM])
		else:
			if NUM not in parts[(AKA[0],'NSN')].aka:
				parts[(AKA[0],'NSN')].aka.append(NUM)
filetoopen.close()
for i in parts:
	if parts[i].cage == 'NSN':
		print parts[i].aka
