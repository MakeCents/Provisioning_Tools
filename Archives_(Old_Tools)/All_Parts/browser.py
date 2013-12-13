import os

location = os.getcwd()

clear = lambda: os.system('cls')

def navpath(paths,location,nextf = ''):
    
    accept = ['.txt', '.036', '.py']
    paths = []
    location +=nextf
    paths.append('<- Back')
    for i in os.listdir(location):
        if i.find('.') == -1 or i[i.find('.'):] in accept:
            paths.append(i)
    for i in range(len(paths)):
       
        print '[{0}] {1}'.format(i,paths[i])
    return paths, location
nextf = ''
while True:
    clear()
    print '***********File Locator************'.upper().center(64)
    print
    print 'Current path:'
    print location
    print
    try:
        
        paths, location = navpath([],location, nextf)
        try:
            nextf = '\\' + paths[int(raw_input('Option? '))]
        except:
            nextf = 'error'
        if nextf == '\<- Back':
            nextf = ''
            n = (len(location) - 1) - location[::-1].index('\\')
            location = location[:n]
        elif nextf == 'error':
            clear()
            print '********************************'
            print '     Not a valid selection'
            print '********************************'
            nextf = ''
    except:
        located = location + nextf
        break
print '****************************************************************'
print 'File Selected:'.center(64)
print (location + '\\').center(64)
print nextf[1:].center(64)
print '****************************************************************'

filetoopen = open(located)



