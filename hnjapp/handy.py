#! coding=utf-8
'''
* @Author: zmFeng
* @Date: 2018-10-19 13:41:06
* @Last Modified by:   zmFeng
* @Last Modified time: 2018-10-19 13:41:06
handy utils for daily life
'''

from datetime import datetime
from os import (makedirs, path, rename, sep, listdir)

from utilz import (getfiles, trimu)


class CadDeployer(object):
    """
    when (maybe C1)'s JCAD file comes, send them to related folder
    also check if they're already there, if already exists, prompt the user
    also decrease the pending list
    """

    def __init__(self, tar_fldr=None):
        self._tarfldr = tar_fldr

    def deploy(self, src_fldr, tar_fldr=None):
        """
        deploy the jcad files in src_fldr to tar_fldr, if tar_fldr is ommitted,
        deploy to self._tarfldr
        return a tuple as (list(stynos deployed.), list(dup. stynos))
        """
        fns, stynos, dups = [path.join(src_fldr, x) for x in listdir(src_fldr)], [], []
        if not fns:
            return
        if tar_fldr is None:
            tar_fldr = self._tarfldr
        for fn in fns:
            styno = path.splitext(path.basename(fn))
            styno = (trimu(styno[0]), styno[1])
            var0 = self._exists(styno[0])
            if var0:
                dups.append((fn, var0))
            else:
                stynos.append("".join(styno))
                dt = datetime.fromtimestamp(path.getmtime(fn))
                dt = "%s%s%s%s%s" % (dt.strftime("%Y"), sep, dt.strftime("%m%d"), sep, styno[0])
                dt = path.join(tar_fldr, dt)
                if not path.exists(dt):
                    makedirs(dt)
                rename(fn, path.join(dt, stynos[-1]))
        self._modlist(stynos, "delete")
        return (stynos, dups)

    def _exists(self, styno):
        #FIXME: check if styno exists current folder or child folders
        return None

    def addlist(self, stynos):
        self._modlist(stynos, "add")

    def _modlist(self, stynos, action="add"):
        #FIXME
        pass
