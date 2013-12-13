First_pos = {
	'P':['A','B','C','D','E','F','G','H','R','Z'],
	'K':['D', 'F', 'B'],
	'M':['O','F','H','L','G','D'],
	'A':['O','F','H','L','G','D'],
	'X':['A','B','C','D']
	}
Third_pos = ['O','2','3','4','5','6','F','G','H','K','L','D','Z']
Fourth_pos = ['O','2','3','4','5','6','F','G','H','K','L','D','Z','B']
Fith_pos = ['O', 'F','G','H','K','L','D','Z','A']
#Sixth_pos = ['1','2','3','6','8','9','E','J','P','R','T']

vsc = []
for f in First_pos:
	for o in First_pos[f]:
		for t in Third_pos:
			for fo in Fourth_pos:
				for fi in Fith_pos:
					vsc.append(f + o + t + fo + fi)

pax = []
digits = [
        '1st digits',
        '2nd digits',
        '3rd digits',
        '4th digits',
        '5th digits',
        ]

#BEAR used
##include = [ #things to include
##['P', 'M', 'X'] #start with
##,['A','B','C','F'] #second
##,['F', 'D'] #third
##,['D', 'F'] #fourth
##,['D', 'Z', 'F'] #fith
##]
#ECP used
include = [ #things to include
['P', 'M', 'X'] #start with
,['A','B','C','F'] #second
,['F', 'D', 'H'] #third
,['D', 'Z', 'F', 'H'] #fourth
,['D', 'Z', 'F', 'H'] #fith
]
print "The following filters have been applied:"

def pinclude(st):
        for i in include:
                print digits[st]+":", ' and '.join(i)
                st+=1

pinclude(0)
print 
print 'If you would like to change what SMR codes are accepted, enter the number \
        digit you would like to edit.'
print 'When finished, or if you would not like to edit, press enter or enter "N".'
print
def lineit(statement):
        print
        print '*' * 55
        print statement
        print '*' * 55
        print
while True:
        change = raw_input("Would you like to change this? Which one? (1-5)  ")
        
        if change == "" or change== 'n' or change == 'N':
                break
        elif change.isdigit() == True:
                nc = int(change)- 1
                if nc in range(0,6):
                        print
                        ans = raw_input('What would you like ' +
                                digits[nc] + ' to be? ' + ','.join(include[nc]) + '?  ')
                        if ans != "":
                                include[nc] = [x.strip().upper() for x in ans.split(',')]
                else:
                        
                       lineit('I did not recognize that response. Please try again.')                
        else:
                lineit('I did not recognize that response. Please try again.')
        pinclude(0)
                
for i in vsc:
	inc = True
	for o in range(len(i)):
		if [x for x in i][o] not in include[o]:
			inc = False
	if inc == True:
		pax.append(i)



print
print 'Total possible: ' + str(len(vsc))
with open('SMR\\All_SMR_Codes.txt', 'w') as v:
	for i in vsc:
		v.write(i + '\n')
print 'Total in filtered: ' + str(len(pax))
with open('SMR\\Filtered_SMR_Codes.txt', 'w') as v:
	for i in pax:
		v.write(i + '\n')
