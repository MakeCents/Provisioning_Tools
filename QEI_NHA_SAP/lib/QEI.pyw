QEI = {}

def add(cpn, qty):
        if cpn not in QEI:
                QEI[cpn] = qty
        else:
                QEI[cpn] += qty
