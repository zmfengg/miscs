#! coding=utf-8 
'''
* @Author: zmFeng 
* @Date: 2018-06-22 22:15:00 
* @Last Modified by:   zmFeng 
* @Last Modified time: 2018-06-22 22:15:00
tests that I can choose to run. use along with VS's debug function
'''

import unittest
import sys
from os import path
import logging

thispath = path.dirname(__file__)
logger = logging.getLogger("tests")
logger.setLevel(logging.DEBUG)

def main():
    vl = 2
    if len(sys.argv) > 1:
        #don't know why this does not work while call by the debugger with arguments
        mdls = sys.argv[1].split(",")
        args = sys.argv[2:] if len(sys.argv) > 2 else None
        for x in mdls:
            unittest.main(x, argv = args, verbosity= vl)
    else:
        mdls = "tests.keytest".split(",")
        if mdls:
            #below won't work, find better way, it run only one test and quit
            for x in mdls:
                unittest.main(x, verbosity= vl)
            #"pytest.testutils"
        else:
            unittest.main(verbosity= vl, argv = ["discover"])

if __name__ == "__main__":
    main()
