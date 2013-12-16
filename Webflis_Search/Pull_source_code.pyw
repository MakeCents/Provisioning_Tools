import urllib
import sys
sys.path.insert(0,'lib')


import PART
import NSN
import WRITE

nsntopart = {}
nsnText = []
with open('NSNtoParts.txt', 'r') as n:
	lines = n.readlines()
	nsns = [x.split()[0] for x in lines]
	parts = [x.split()[1] for x in lines]
	for i in range(len(nsns)):
		nsntopart[nsns[i]]=parts[i]
	for part in nsns:
		nsn = NSN.findPart(part)
		if nsntopart[part] not in nsn:
			WRITE.write('Search_Results/Errors/',nsn, part)
			if nsn not in nsnText:
				nsnText.append(nsn)
		else:
			WRITE.write('Search_Results/',nsn, part)
	print len(nsnText)

#Put way to chose what to do and then load list of what it is from text file
#give option for one search or batch file

#Example below

#PART.findPart(['15003'])
	
#url = 'https://www.logisticsinformationservice.dla.mil/bincs/details.aspx?CAGE=15434'

#print urllib.urlopen(url).read()
