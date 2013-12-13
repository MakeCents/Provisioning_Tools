
def loadf(filef):
    ext = ['.036', '.txt']
    try:
        filef = str(raw_input('File name to check?  '))
        t = filef
        f = open(filef)
    except:
        try:
            for i in ext:
                filef = t+i
                try:
                    f = open(filef)
                    return f
                except:
                    pass
        except:
            pass
        filef = t
        print '==============================='
        print 'File "' + filef + '" not found.'
        print '==============================='
        return loadf('')
    
    return f


def checker(f):
    listof = []
    print 'Processing: ' + f.name.upper()
    for i in f:
        p = i[:-1]
        if p[-2:] == '1A':
            cage = p[13:18]
            pn = p[18:50]
            PLISN = p[6:11]
            listof.append((pn, cage, PLISN))
    f.close()
    print 'Checking for differences...'
    bad = []
    for i in listof:
        for o in range(len(listof)):
            if i[0] == listof[o][0]:
                if i[1] != listof[o][1]:
                    if len(bad) == 0:
                        bad.append(listof[o])
                    else:
                        found = False
                        for b in bad:
                            if b[0] == listof[o][0] and b[1] == listof[o][1]:
                                found = True
                                break
                        if found == False:
                            bad.append(listof[o])

    print 'The following have differences:'
    print 'PLISN | CAGE  |  Part Number'
    print '============================'

    for i in bad:
        print '{0} | {1} | {2}'.format(i[2],i[1], i[0])
while True:
    f = loadf('')
    checker(f)
