import urllib

def findPart(nsn):
    url = "http://www.dlis.dla.mil/webflis/pub/pub_search.aspx?niin={0}&Par=True&Expanded=True".format(nsn)
    answer =  urllib.urlopen(url).read()
    number = answer.find('-----[/Ap')
    nnumber = answer[number:].find('<!-- --------------[')
    answer = answer[number:nnumber+number].replace('pub_help.aspx','http://www.dlis.dla.mil/webflis/pub/pub_help.aspx')
    answer = answer.replace('pub_search.aspx','http://www.dlis.dla.mil/webflis/pub_search.aspx')
    answer = answer[answer.find(nsn)+13:].replace(nsn, '<a href = "' + url + '"">'+nsn+'</a>')
    answer = answer[answer.find('<div'):]
    return answer

