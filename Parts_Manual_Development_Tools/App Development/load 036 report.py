class Report(object):
        def __init__(self):
                self.PLISNs = {}
                self.PLISNo = []
        def Load(self, FileName):
                if FileName == '':
                        FileName = 'W9E215.036'
                try:
                    self.file = open(FileName)
                except:
                    print 'File ' + FileName + ' not found. Please try again. Remember to include the extenstion.'
                    self.Load(str(raw_input('FileName?  ')))
                for e in self.file:
                        if e[79] == ' ':
                            e = e.replace('\n', '')
                            self.Report = e
                        else:
                            self.AP(e)
                self.UI()
        def AP(self, PLISN):
                """ Add PLISN
                """
                if PLISN[6:10] not in self.PLISNs:
                        self.PLISNo.append(PLISN[6:10])
                        self.PLISNs[PLISN[6:10]]={}
                        self.UC(PLISN)
                else:
                        self.UC(PLISN)
        
        def UC(self, PLISN):
                """ Update Card
                """
                PLISN = PLISN.replace('\n','')
                try:
                    self.PLISNs[PLISN[6:10]][PLISN[79]][(int(PLISN[78])-1)] = PLISN
                    
                except:
                    self.PLISNs[PLISN[6:10]][PLISN[79]] = ['', '', '', '', '', '']
                    self.PLISNs[PLISN[6:10]][PLISN[79]][(int(PLISN[78])-1)] = PLISN

        def whead(self,special,begin,answer):
                if special == 'w':
                        if begin == 1:
                                f = open(str(answer) + '.txt', 'w')
                                try:
                                    f.write(self.Report + '\n')
                                    print answer + ' written as txt file.'
                                except:
                                    f.write('No Header' + '\n')
                                    print answer + ' written as txt file.'
                else:
                        if begin == 1:
                            try:
                                print  self.Report
                            except:
                                print 'No Header'

        def returnit(self, answer):
                print '==============================================================================='
                count = 1
                special = ''
                if answer == '':
                    pass
                elif answer == 'new':
                    self.Load(str(raw_input('FileName?  ')))
                else:
                    begin = 1
                    if len(answer) > 2:
                            if answer[-2] == '/':
                                special = answer[-1]
                                answer = answer[:len(answer)-2]
                    if answer == 'all':
                        answer = answer.upper()
                        for e in self.PLISNo:
                            for i in self.PLISNs[e]:
                                for o in range(5):
                                    if self.PLISNs[e][i][o] == '':
                                        pass
                                    else:
                                        self.whead(special,begin,answer)
                                        begin +=1
                                        if special == 'w':
                                            f = open(str(answer) + '.txt', 'a')
                                            f.write(self.PLISNs[e][i][o] + '\n')
                                        if special == 'p':
                                            count +=1
                                            print self.PLISNs[e][i][o]
                                            if count > 40 and self.PLISNs[e][i][o][79] >= 'H':
                                                count = 1
                                                self.Cont()
                    elif answer == 'pList':
                        self.pList(special,begin,answer)
                    else:
                        begin = 1
                        answer = answer.upper()
                        for e in self.PLISNo:
                            Cardl = []
                            if e[:len(answer)] == answer:
                                for i in self.PLISNs[e]:
                                    Cardl.append(i)
                                Cardl.sort()
                                for c in Cardl:
                                    for o in range(5):
                                        if self.PLISNs[e][c][o] == '':
                                            pass
                                        else:
                                                self.whead(special,begin,answer)
                                                begin +=1
                                                if special == 'w':
                                                        f = open(str(answer) + '.txt', 'a')
                                                        f.write(self.PLISNs[e][c][o] + '\n')
                                                elif special == 'p':
                                                    count +=1
                                                    print self.PLISNs[e][c][o]
                                                    if count > 40 and self.PLISNs[e][c][o][79] >= 'H':
                                                        count = 1
                                                        self.Cont()
                                                else:
                                                    print self.PLISNs[e][c][o]
                                                    if count > 40 and self.PLISNs[e][c][o][79] >= 'H':
                                                        count = 1
                                                        self.Cont()
                                                    

        def pList(self,special,begin,answer):
            self.whead(special,begin,answer)
            begin +=1
            count = 1
            self.whead(special,begin,answer)
            begin +=1
            for e in self.PLISNo:
                if special == 'w':
                        f = open(str(answer) + '.txt', 'a')
                        f.write(e + '\n')
                elif special == 'p':
                    count +=1
                    print e
                    if count > 40:
                        count = 1
                        self.Cont()
                else:
                        print e
        def Cont(self):
            if str(raw_input('Continue listing? Type "n" for no.   ')) == 'n':
                self.UI()
            print '======================================================================'
            return
        
        def UI(self):
            answer = ' '
            while answer != '':
                answer = str(raw_input('Please enter?   '))
                self.returnit(answer)
                if answer == 'quit':
                    try:
                        exit()
                    except SystemExit:
                        print 'You hit cancel'
            print 'Commands:'
            print '"all" -  for entire list'
            print '"new" - to load a new report'
            print '"pList" - list all PLISNs'
            print 'add a "/p" to the end for list with pause'
            print 'add a "/w" to the end to write the list'
            print 'PLISN, such as "AAAA", or part of PLISN, "AAA"'
            print '"quit" - will quit.'
            self.UI()
print 'FileName is W9E215.036 by default. Pressing enter will load that file.'
Report().Load(str(raw_input('FileName?  ')))

                   
