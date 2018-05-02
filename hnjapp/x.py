# coding=utf-8
"""
 extended command help to simplify things
 
"""

def xdir(obj,args = None):
    if(obj == None):
        rc = None
    else:
        if(not args):
            rc = dir(obj)
        else:
            st0 = set()
            xall = dir(obj)
            for s in args.split(","):
                st0 = st0.union(set([x for x in xall if(x.find(s) >= 0)]))
            rc = None if(len(st0) == 0) else sorted(st0)        
    return rc

if(__name__ == "__name__"):
    print("x.xdir(obj,[args]) help to show items in the obj for example: %s shows " % 
          '"xdir("","trans"))')
    xdir("","trans")