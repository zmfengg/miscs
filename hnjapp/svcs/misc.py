'''
#! coding=utf-8
@Author:        zmFeng
@Created at:    2019-07-02
@Last Modified: 2019-07-02 1:06:29 pm
@Modified by:   zmFeng
fascade for misc services
'''

from ._misc import StylePhotoSvc, PhotoQueryAsMail
from ._stysn import StyleFinder, LKSizeFinder, JOSnIndex

__all__ = ['JOSnIndex', 'LKSizeFinder', 'PhotoQueryAsMail', 'StyleFinder', 'StylePhotoSvc']
