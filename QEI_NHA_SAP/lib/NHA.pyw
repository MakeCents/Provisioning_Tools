NHA = {}
AllNHA = set([])
IND = {}

def add(PLISN, nha):
    NHA[PLISN] = nha
    AllNHA.add(nha)

def ind(IC, plisn):
	IND[plisn] = IC
