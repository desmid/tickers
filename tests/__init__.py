import os, sys

uselib = os.path.dirname(__file__)
uselib = os.path.realpath(uselib)
uselib = os.path.split(uselib)[0]
uselib = os.path.join(uselib, 'pythonpath')

sys.path.append(uselib)
