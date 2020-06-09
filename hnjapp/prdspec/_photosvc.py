'''
#! coding=utf-8
@Author:        zmfengg
@Created at:    2020-06-03
@Last Modified: 2020-06-03 5:16:26 pm
@Modified by:   zmfengg
Service for photo processing for the new style system
'''

from hnjapp import config
from hnjapp.svcs.misc import StylePhotoSvc

class StylePhotoSvcX(StylePhotoSvc):

    def getPhotos(self, styno, atype='styno', hints=None, **kwds):
        '''
        in the new style case, we don't need to guest any more, just get the ones with prefix
        '''
        pass
