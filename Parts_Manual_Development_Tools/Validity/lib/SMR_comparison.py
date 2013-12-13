
############################################
######### Locate file to analyze ###########
############################################
def mline(statement):
        print
        print '#' * 80
        print statement.center(80)
        print '#' * 80
        print
        
import os
'''Gets the files in the Reports folder
'''
def loopit(times=0,ran=0):
        def paths(times=0):
                mline('Select a file to Analyze')
                        
                lof = ['Update Valid SMR Codes']
                if times == 0:
                        print '[' + '0' + ']', lof[0]
                pathtof = ['Error']
                start = 0
                mypath = os.getcwd()[:-8]
                walker = os.walk(mypath+"\\Reports\\")
                for i in walker:
                        if "svn" not in i[0]:
                                for o in i[2]:
                                        if o[-4:] == '.036' and len(o) < 11:
                                                start+=1
                                                print '[' + str(start) + ']' , o[:-4]
                                                lof.append(o[:-4])
                                                pathtof.append(i[0])
                return lof, pathtof
        
        def askfnum(times):
                if times == 0:
                        lof, pathtof = paths()
                else:
                        lof, pathtof = paths(1)
                print
                ans = raw_input('Select a number please: ' + '\n' + 'Which file would you like to analyze?  ')
                print
                try:
                        test = lof[int(ans)]
                except:
                        ans, lof, pathof = askfnum(1)
                return ans, lof, pathtof
        
        
        ans, lof, pathtof = askfnum(times)
        
        while ans == '0':
                if ran == 0:
                        mline('Build valid SMR codes')
                        import Valid_SMR_codes
                        ans, lof, pathtof = askfnum(1)
                        ran = 1
                else:
                        ans = '0'
                        mline('You already ran this, choose something else.')
                        ans, lof, pathtof = askfnum(1)
                        
                
        filename = lof[int(ans)]
        ptf = pathtof[int(ans)]
        return filename, ptf


############################################
############ Begin to analyze ##############
############################################

def runanalyzer(filename, ptf, PLISNs, Parts, SMR_Codes):
        file036 = ptf + '\\' + filename + ".036"
        print 'Analyzing 036 file...'
        print
        try:
                with open(file036) as newfile:
                        count = 10000
                        for i in newfile:
                                PCCN = i[0:6].strip()
                                PLISN = i[6:11].strip()
                                if PLISN not in PLISNs:
                                        try:
                                                PLISNs[PLISN] = _PLISN(PLISN, PCCN.strip(), i[12].strip(),
                                                                       i[13:18].strip(), i[18:50].strip(), count)
                                                count +=1
                                                Parts.add(i[18:50].strip())
                                        except IndexError:
                                                pass
                                elif i[-3:-1] == '1B':
                                        PLISNs[PLISN].Update(i[15:28].strip(), i[64:70].strip())
                                        SMR_Codes.add(i[64:70].strip())
                print 'Writing "TEXT\\SMR_Codes_used_on_'+filename+'.txt"...'
                with open('TEXT\\SMR_Codes_used_on_'+filename+'.txt', 'w') as t:
                        for i in SMR_Codes:
                                t.write(i + '\n')
                                
                print 'Calculating SMR and NSN differences for each part number...'
                Mistakes = {}
                for i in PLISNs:
                        P = PLISNs[i]
                        if P.PN not in Mistakes:
                                try:
                                        Mistakes[P.PN] = {'PLISN': [P.PLISN], 'SMR': [P.SMR], 'NSN': [P.NSN], 'PROB': []}
                                except:
                                        pass
                        for o in PLISNs:
                                NP = PLISNs[o]
                                if NP.PN in Mistakes:
                                        if NP.SMR not in Mistakes[NP.PN]['SMR'] or NP.NSN not in Mistakes[NP.PN]['NSN']:
                                                if NP.SMR not in Mistakes[NP.PN]['SMR']:
                                                        if 'SMR' not in Mistakes[NP.PN]['PROB']:
                                                                Mistakes[NP.PN]['PROB'].append('SMR')
                                                if NP.NSN not in Mistakes[NP.PN]['NSN']:
                                                        if 'NSN' not in Mistakes[NP.PN]['PROB']:
                                                                Mistakes[NP.PN]['PROB'].append('NSN')
                                                Mistakes[NP.PN]['SMR'].append(NP.SMR)
                                                Mistakes[NP.PN]['PLISN'].append(NP.PLISN)
                                                Mistakes[NP.PN]['NSN'].append(NP.NSN)

                print 'Loading "SMR\\Filtered_SMR_Codes.txt"...'
                with open('SMR\\Filtered_SMR_Codes.txt', 'r') as filt:
                        filtered = []
                        for f in filt:
                                filtered.append(f[:-1])
                print 'Writing "CSV\\SMR and NSN differences in '+ filename +'".csv...'
                with open('CSV\\SMR_and_NSN_Differences_'+ filename +'.csv', 'w') as Tracker:
                        Tracker.write('PART NUMBER,DIFFERENCE,PLISN,SMR,NSN' + '\n')
                        
                        for i in Mistakes:
                                if len(Mistakes[i]['PLISN']) > 1:
                                        Tracker.write('"' +  i.strip() + '"' + ',' + \
                                                '"' + ', '.join(Mistakes[i]['PROB']) + '"' + ',' + \
                                                '"' + ', '.join(Mistakes[i]['PLISN']) +'"' + ',' + \
                                                '"' + ', '.join(Mistakes[i]['SMR']) + '"' + ',' + \
                                                '"' + ', '.join(Mistakes[i]['NSN']) + '"' + '\n')
                                else:
                                        good = True
                                        for o in Mistakes[i]['SMR']:
                                                if o not in filtered:
                                                        Mistakes[i]['PROB'].append('Invalid SMR')
                                                        Tracker.write('"' +  i.strip() + '"' + ',' + \
                                                                '"' + ', '.join(Mistakes[i]['PROB']) + '"' + ',' + \
                                                                '"' + ', '.join(Mistakes[i]['PLISN']) +'"' + ',' + \
                                                                '"' + ', '.join(Mistakes[i]['SMR']) + '"' + ',' + \
                                                                '"' + ', '.join(Mistakes[i]['NSN']) + '"' + '\n')
                                                        break
                                                
                                        
                                        
                        

        except:
                print "No such file. Please try again."
                filename, ptf = loopit(1,1)
                runanalyzer(filename, ptf,PLISNs, Parts, SMR_Codes)

PLISNs = {}
Parts = set([])
SMR_Codes = set([])

class _PLISN(object):
	"""PLISN class"""
	def __init__(self, PLISN,PCCN,IND,CAGE,PN,ID):
		self.PLISN = PLISN
		self.PCCN = PCCN
		self.IND = IND
		self.CAGE = CAGE
		self.PN = PN
		self.ID = ID
	def Update(self, NSN, SMR):
                if NSN == "":
                        NSN = "None"
		self.NSN = NSN
		self.SMR = SMR

filename, ptf = loopit()
runanalyzer(filename, ptf,PLISNs, Parts, SMR_Codes)
mline('Done')
