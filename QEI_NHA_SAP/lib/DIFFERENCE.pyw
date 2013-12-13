def Difference(originalFile, newFile):
        oldfiles = open(originalFile)
        newfiles = open(newFile)
        oldfile = [x[:-1] for x in oldfiles]
        newfile = [x[:-1] for x in newfiles]
        #dictionary of ranges per card
        Cards = {'A':[[12,13],[13,18],[18,50],[50,51],[51,52],[52,53],[53,54],[54,55],[55,74],[74,75],[75,77]],
                'B':[[12,15],[15,28],[28,32],[32,34],[34,44],[44,46],[46,56],[56,61],[61,64],[64,70],[70,71],[71,73],[73,74],[74,75],[75,76],[76,77]],
                'C':[[12,17],[17,18],[18,21],[21,25],[25,30],[30,38],[38,46],[46,53],[53,59],[59,64],[64,69],[69,73],[73,74],[74,77]],
                'D':[[12,20],[20,52],[52,53],[53,54],[54,55],[55,56],[56,57],[57,59],[59,62],[62,65],[65,68],[68,71],[71,74],[74,77]],
                'E':[[12,14],[14,16],[16,18],[18,20],[20,22],[22,24],[24,26],[26,29],[29,32],[32,35],[35,38],[38,41],[41,44],[44,47],[47,50],[50,53],[53,56],[56,59],[59,65],[65,71],[71,73],[73,74],[74,75],[75,76],[76,77]],
                'F':[[12,27],[27,29],[29,39],[39,49],[49,51],[51,56],[56,57],[57,63],[63,69],[69,77]],
                'G':[[12,33],[33,39],[39,77]],
                'H':[[12,30],[30,32],[32,77]],
                'J':[[12,15],[15,19],[19,23],[23,25],[25,26],[26,29],[29,40],[40,45],[45,53],[53,54],[54,55],[55,60],[60,68],[68,69],[69,70],[70,71],[71,76],[76,77]],
                'K':[[12,15],[15,19],[19,23],[23,77]],
                'M':[[12,77]]}
        class plisn(object):
                def __init__(self, PLISN):
                        self.old = {}
                        self.new = {}
                        self.oldcards = []
                        self.newcards = []
                        self.consolcards = []
                        self.compared = ""
                        self.PLISN = PLISN
                        self.convert()

                def convert(self):
                        self.order = ""
                        for i in self.PLISN:
                                if ord(i) > 64:
                                        self.order += str(ord(str(i))-43)
                                else:
                                        self.order += str(ord(str(i)))

                def olds(self, card):
                        self.old[card[-3:]] = card
                        self.oldcards.append(card[-3:])
                def news(self, card):
                        self.new[card[-3:]] = card
                        self.newcards.append(card[-3:])
                def compare(self):
                        ans = self.compared
                        order = ['01A', '02A', '03A', '04A', '01B', '02B','03B', '04B','01C', '02C',
                                        '03C', '04C','01D', '02D','03D', '04D','01E', '02E','03E', '04E','01F',
                                        '02F','03F', '04F', '01G', '02G','03G', '04G', '01H', '02H','03H', 
                                        '04H','01J', '02J','03J', '04J','01K', '02K','03K', '04K','01M', '02M',
                                        '03M', '04M']
                        for i in order:
                                if i in self.newcards or i in self.oldcards:
                                        self.consolcards.append(i)
                        if self.new == {}:
                                self.compared += self.old["01A"][:11] + "D " + self.old["01A"][13:50] + (" " * 27) + self.old["01A"][77:] + '\n'
                        elif self.old == {}:
                                for i in self.newcards:
                                        self.compared += self.new[i] + '\n'
                        else:
                                for i in self.consolcards:
                                        if i not in self.oldcards:
                                                self.compared += self.new[i][:11] + "M" + self.new[i][12:] + '\n'
                                        elif i not in self.newcards:
                                                self.compared += self.old[i][:11] + "G"
                                                for param in Cards[i[-1]]:		
                                                        if self.old[i][param[0]:param[1]].strip() == "":
                                                                self.compared += FixLengths(param[1] - param[0])
                                                        else:
                                                                self.compared +=  'G' + FixLengths(param[1] - param[0])[1:]
                                                self.compared += i + '\n'
                                        elif i not in self.newcards:
                                                ans += self.old[i][:11] + "G"
                                                for letter in self.old[i][12:77]:
                                                        if letter != " ":
                                                                ans +=  'G'
                                                        else:
                                                                ans+=" "
                                                ans+=self.old[i][77:]
                                                self.compared = ans + '\n'
                                        elif self.new[i] != self.old[i]:
                                                self.compared += CompareCards(self.old[i], self.new[i])				
                        self.compared = self.compared[:-1]

        def CompareCards(old, new):
                beginCard = new[:11]
                endCard = new[77:]
                mcard = ""
                gcard = ""
                comparison = ""
                keyparam = [[12,17],[12,20],[12,27],[20,52],[55,56]] #key parameters that get g cards even for modifications
                #associated data has not been incorporated yet
                associatedparams = {31:[[50,51]],68:[[50,51]],29:[[51,52]],39:[[18,21]],72:[[52,53],
                                                        [53,54]],111:[[29,39],[39,49],[33,39],[27,29],[51,56],[49,51]]} # add keyparams together for key to check associated data
                for param in Cards[new[-1]]:
                        oldparam = CheckCard(param,old)
                        newparam = CheckCard(param,new)
                        if oldparam == newparam:
                                if int(new[-2]) > 1:
                                        if param in [[13,18],[18,50]]:
                                                gcard += oldparam
                                        else:
                                                mcard+=FixLengths(param[1] - param[0])
                                                gcard+=FixLengths(param[1] - param[0])
                                else:
                                        mcard+=FixLengths(param[1] - param[0])
                                        gcard+=FixLengths(param[1] - param[0])
                        elif newparam == FixLengths(param[1] - param[0]):
                                gcard += 'G' + FixLengths(param[1] - param[0])[1:]
                                mcard += FixLengths(param[1] - param[0])
                        else:
                                mcard += newparam
                                if param in keyparam:
                                        gcard += oldparam
                                elif int(new[-2]) > 1:
                                        if param in [[13,18],[18,50]]:
                                                gcard += oldparam
                                else:
                                        gcard += FixLengths(param[1] - param[0])
                if mcard.strip() != "":
                        mcard = beginCard + "M" + mcard +endCard + '\n'
                else:
                        mcard = ""
                if gcard.strip() != "":
                        gcard = beginCard + "G" + gcard + endCard + '\n'
                        return gcard + mcard
                return mcard

        def FixLengths(length):
                return length * " "

        def CheckCard(param, card):
                return card[param[0]:param[1]]

        PLISNs = {}

        for i in oldfile:
                if i[6:11].strip() not in PLISNs:
                        PLISNs[i[6:11].strip()] = plisn(i[6:11].strip())
                        PLISNs[i[6:11].strip()].olds(i)
                else:
                        PLISNs[i[6:11].strip()].olds(i)
        for i in newfile:
                if i[6:11].strip() not in PLISNs:
                        PLISNs[i[6:11].strip()] = plisn(i[6:11].strip())
                        PLISNs[i[6:11].strip()].news(i)
                else:
                        PLISNs[i[6:11].strip()].news(i)
        printOrder = []
        with open('Differences_Report.036', 'w') as report:
                for i in PLISNs:
                        printOrder.append(PLISNs[i])
                printOrder.sort(key=lambda x: x.order)
                for i in printOrder:
                        i.compare()
                        if i.compared != "":
                                report.write(i.compared + '\n')

        oldfiles.close() 
        newfiles.close()
