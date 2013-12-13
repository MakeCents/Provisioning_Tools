
def checkCard(line):
    if line[-1] == '\n':
        return line[-4:-1].strip()
    else:
        return line[-3:].strip()

#returns the indenture code if this is the A card
def get_Ind(line):
    if checkCard(line) == "01A":
        return line[12]

def get_cpn(line, cpn):
    if checkCard(line) == '01A':
        return line[13:50].strip()
    return cpn

def get_plisn(line):
    if checkCard(line) == '01A':
        return line[6:11].strip()
    else:
        return