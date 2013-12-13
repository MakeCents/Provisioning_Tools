import os
import subprocess
location = os.getcwd()

def mcm(search):
    return '<a href="http://www.mcmaster.com/#{0}">{0}</a>'.format(search)

def NSN(search):
    return '<a href="https://www.dlis.dla.mil/webflis/pub/pub_search.aspx?niin={0}&newpage=1">{0}</a>'.format(search)

def CAGE(search):
    return '<a href="https://www.logisticsinformationservice.dla.mil/bincs/details.aspx?CAGE={0}" target="_blank">{0}</a>'.format(search)

def goto(search,plisn,time = 0):
    '''Part or Cage
##    '''
    What = 'mcmaster.com/#{0}"'.format(search)
    lk = eval('mcm(search)')

    writehtmllink('<meta http-equiv="refresh" content="{1}; http://www.{0}>'.format(What,time),search,lk, plisn)
   
    print 'File ' + str(location) + '\\' + search + '.html written.'

def writehtmllink(link,to,lk, plisn):
    linkto = open(plisn + "_" + to+'.html','w')
##    linkto.write('<p>Please wait...</p>')
##    linkto.write('<p>Opening ' + lk + '</p>')
   
##    linkto.write('<p>Please click the link above if you are not redirected</p>')
    linkto.write(link)
    linkto.close()

htmls = []
with open('PLISN_Partnumber.txt', 'r') as f:
    for i in f:
        htmls.append(i[:-1])
for i in htmls:
    plisn, search = i.split("_")
    goto(search,plisn)

