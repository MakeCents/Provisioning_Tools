def loadNames(l):
    print
    answers = []
    for answer in l:
        print '\t' + '[{0}]  {1}'.format(l.index(answer), answer)
    print
    for name in ['036']:
        answer = str(raw_input('\t' +'\t' +'Please select the ' + name + ' file.   '))
        if answer == '0':
            print "The file must be located in this folder and be of the .036 file extention"
            break
        else:
            #Load will return a fileName
            try:
                return Load(l[int(answer)])
            except:
                return loadNames(l)
    return answers

def Load(fName = ""):
        if fName == "":
                fName = str(raw_input("What is the file name to read?  "))
        error = True
        while error:
                try:
                        if '.036' not in fName and '.txt' not in fName:
                                fName+='.036'
                        with open(fName, 'r') as Input:
                               return Input.readlines(), fName
                        error = False
                except:
                        print fName + " not found."
                        fName = str(raw_input("What is the file name?  "))



def Write(writeList):
        print
        print
        fName = str(raw_input('\t' +'\t' +"What name would you like to give this file?  "))
        error = True
        while error:
                try:
                        if '.036' not in fName and '.txt' not in fName:
                                fName+='.036'
                        with open(fName, 'w') as Output:
                                for item in writeList:
                                        Output.write(item)
                        error = False
                except:
                        print fName + " not found."
                        fName = str(raw_input("What is the file name?  "))

        return fName
