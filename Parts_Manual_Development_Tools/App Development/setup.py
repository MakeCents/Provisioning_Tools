#Setup file that I made.
'''
 In a cmd prompt, navigate to the folder including this.

 C:\Python27>python setup.py install
 C:\Python27>setup.py py2exe
 
'''
from distutils.core import setup
import py2exe

setup(console=['QEI and Same As PLISN.py']) # change file name here.
