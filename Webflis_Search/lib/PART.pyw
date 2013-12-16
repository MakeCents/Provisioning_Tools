import urllib

def findPart(part):
	url = "http://www.dlis.dla.mil/webflis/pub/pub_search.aspx?part={0}&Par=True&Expanded=True".format(part)
	url = url+' MakeCents'
	answer =  urllib.urlopen(url).read()
	number = answer.find('<tr class="resultspick1">')
	nnumber = answer[number:].find('</table>')
	answer = answer[number:nnumber+number].replace('pub_search.aspx','http://www.dlis.dla.mil/webflis/pub/pub_search.aspx')
	return answer		
