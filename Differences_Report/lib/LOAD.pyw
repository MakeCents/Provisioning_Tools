
def loadNames(l):
    answers = []
    for answer in l:
        print '[{0}]  {1}'.format(l.index(answer), answer)
    print
    for name in ['new', 'old']:
        answer = str(raw_input('Please select the ' + name + ' file.   '))
        if answer == '0':
            print "The file must be located in this folder and be of the .036 file extention"
            break
        else:
            #Load will return a fileName
            answers.append(Load(l[int(answer)]))
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
                               return fName
                        error = False
                except:
                        print fName + " not found."
                        fName = str(raw_input("What is the file name?  "))



