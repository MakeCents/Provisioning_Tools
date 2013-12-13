#same as plisn
SAP = {}
#Part first PLISN used
PP = {}

def add(plisn, sap):
    SAP[plisn] = sap

def pp(cpn,plisn):
    if cpn not in PP:
        PP[cpn] = plisn
