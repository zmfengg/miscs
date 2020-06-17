'''
#! coding=utf-8
@Author:        zmFeng
@Created at:    2019-05-29
@Last Modified: 2019-05-29 9:35:12 am
@Modified by:   zmFeng
Product Spec handler, including Acessor/Normalizer

'''

from ._fromjo import FromJO, JOFormHandler
from ._nmctrl import NameGtr, NameSvc, NameGtr
from ._photosvc import StylePhotoSvcX as StylePhotoSvc

__all__ = ['FromJO', 'JOFormHandler', 'NameGtr', 'NameSvc', 'NameGtr', 'StylePhotoSvc']
