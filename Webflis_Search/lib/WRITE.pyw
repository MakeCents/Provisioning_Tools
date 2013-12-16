import urllib

def write(where,pnum, part):
    name = part.replace('*',"")
    with open(where + name + '.html', 'w') as t:
        #url = url + pnum
        for o in pnum.split('<td>'):
                        if '</td>' in o:
                            t.write('<td>' + o)
                        else:
                            t.write(o)
                        if '</tr>' in o:
                            t.write('<br>')

