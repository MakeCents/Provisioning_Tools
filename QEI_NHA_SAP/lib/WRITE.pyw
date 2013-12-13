#This would be done first, before this module.

def Write(writeList):
        fName = str(raw_input("What is the file name to write?  "))
        error = True
        while error:
                try:
                        if '.036' not in fName and '.txt' not in fName:
                                fName+='.036'
                        with open(fName, 'w') as Output:
                                for item in writeList:
                                        Output.write(item + '\n')
                        error = False
                except:
                        print fName + " not found."
                        fName = str(raw_input("What is the file name?  "))

