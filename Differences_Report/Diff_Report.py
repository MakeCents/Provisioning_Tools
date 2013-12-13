import sys
import os
sys.path.insert(0,'lib')

l = [x for x in os.listdir(os.getcwd()) if x[-4:] == '.036']
l.insert(0, "Not here")

#This would be done first, before this module.
import LOAD
import DIFFERENCE



if len(answers) == 2:
    DIFFERENCE.Difference(answers[0], answers[1])
else:
    print 'Please try again'
    print
    loadNames()
