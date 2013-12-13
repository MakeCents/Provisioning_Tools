NHA = {}
AllNHA = []
IND = {}

def add(PLISN, nha):
    NHA[PLISN] = nha
    AllNHA.append(nha)

def ind(IC, plisn):
	IND[plisn] = IC