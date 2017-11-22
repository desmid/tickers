import os, sys

uselib = os.path.dirname(__file__)
uselib = os.path.join(uselib, '..', 'pythonpath')
uselib = os.path.abspath(uselib)

sys.path.insert(0, uselib)

#print >>sys.stderr, sys.path
