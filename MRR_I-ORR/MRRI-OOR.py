f = open('MRR I-ORR from AFMCI23-104.txt')
ll = {}
for i in f:
    if i[:4] not in ll:
        temp = i.split('//')[1:]
        temp[2] = temp[2][:-1]
        ll[(i.split('//'))[0]] = [temp]
    else:
        temp = i.split('//')[1:]
        temp[2] = temp[2][:-1]
        ll[(i.split('//'))[0]].append(temp)
f.close()

ans = 'start'
def search(ll,ans):
    print
    if ans[0] != '':
        print '*' * 30
        print 'FSC: ', ans[0],ans[1]
        print '*' * 30
    elif ans[1] != '':
        print '*' * 30
        print 'Description: ', ans[1]
        print '*' * 30
    print('  OOR   |  MRR   |  Description')
    print ('=' * 80)
    if '*' in ans[0]:
        for o in ll:
            if ans[0].split('*')[0] in o[:len(ans[0].split('*')[0])]:
                fsc = 's'
                for i in ll[o]:
                    if ans[1].lower() in i[0].lower():
                        if fsc != o:
                            print '-' * 10
                            print 'FSC: ', o
                            print '-' * 10
                        print ' ', i[2], ' |', i[1],'| ', i[0]
                        fsc = o
    elif ans[0] != '':
        for o in ll[ans[0]]:
            if ans[1] == '':
                print ' ', o[2], ' |', o[1],'| ', o[0]
            elif ans[1].lower() in o[0].lower():
                print (o)
    else:
        for o in ll:
            fsc = 's'
            for i in ll[o]:
                if ans[1].lower() in i[0].lower():
                    if fsc != o:
                        print '-' * 10
                        print 'FSC: ', o
                        print '-' * 10
                    print ' ', i[2], ' |', i[1],'| ', i[0]
                    fsc = o
    print('=' * 80)
    print
    
while ans !='':
    print('Type "exit" to exit or press enter with no input')
    print('Type the FSC you wish to look up' + '\n' +\
     '(use * for partial matches i.e. 530*)') + '\n' 
    print('Narrow your search, or search by description,' + '\n' +\
     'by adding a comma and a description (FSC not needed)')
    ans = str(raw_input('FSC, desc?  '))
    if ans == 'exit' or ans == '':
        break
    if ',' in ans:
        ans = ans.split(',')
        ans = [x.strip() for x in ans]
        print ans
        ans = tuple(ans)
        print ans
    else:
        ans = (ans,'')
    try:
        search(ll,ans)
    except:
        if ans[0] == '':
            found = ans[1]
        elif ans[1] == '':
            found = ans[0]
        print found, 'not found','\n','=' * 50
        
