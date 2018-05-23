# coding=utf-8
"""
 extended command help to simplify things
 
"""

import re
import __builtin__ as bi
import inspect

__all__ = ["dir"]

def dir(obj = None,args = None):
    if obj != None:
        xall = bi.dir(obj)
    else:
        frms = []
        frms.append(inspect.currentframe())
        ofrms = None
        try:
            ofrms = inspect.getouterframes(frms[0], 1)
            frms.append(ofrms[1][0])
            xall = frms[1].f_globals.keys()
        finally:
            if frms:
                for x in frms:
                    if x: del x
                del frms
            if ofrms: del ofrms
    if args:
        rc = set()
        for s in args.split(","):
            ptn = re.compile(s, re.IGNORECASE)
            rc = rc.union(set([x for x in xall if(ptn.search(x))]))
    else:
        rc = xall
    return None if not rc else sorted(rc)

if(__name__ == "__main__"):
    print("x.dir(obj,[args]) help to show items in the obj,for example: %s shows:" % 
          '"x.dir("","trans"))')
    print(dir("","trans"))