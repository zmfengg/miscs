#! coding=utf-8
'''
* @Author: zmFeng
* @Date: 2018-11-22 20:49:07
* @Last Modified by:   zmFeng
* @Last Modified time: 2018-11-22 20:49:07
 try to build catalog page
'''

from ._miscs import getvalue
from xlwings.constants import PaperSize


class CatBase(object):
    """
    base class for this module
    """
    pass


class SettingBase(CatBase):
    """
    class to hold settings
    """
    pass

    def fromCfg(self, cfg):
        """
        read the necessary settings from the cfg provided
        the cfg can be of a dict/a json object
        """
        pass

    def toCfg(self):
        """
        return a json that can be store to a json file
        """
        pass


class Block(object):
    """
    represent a block for data displaying
    """

    def __init__(self, page, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._page = page

    @property
    def page(self):
        """ return the page this block belongs to """
        return self._page

    def address(self):
        """
        the address of this block
        """
        pass

    def size(self):
        """
        return the height and width of this block
        """
        pass

    def init(self):
        """
        init this block itself, first, check if I'm inited, if yes, do nothing
        """


class Page(object):
    """
    represent a page in the catalog
    """

    def __init__(self, config, sht):
        super().__init__()
        self._setting = PageSettings(config=config)
        
    @property
    def headers(self):
        """
        return the header defs of one page
        """
        pass

    def _add_listener(self, listener, lis_type):
        pass

    def next(self):
        """
        return next block
        """
        pass

    def get_block(self, idx):
        """
        return the idx_th block inside this page
        """


class _HeaderSettings(SettingBase):
    #holds the heights of rows
    _rows = None
    def rows(self):
        #TODO get the rows from config
        if not self._rows:
            self._rows = [10, 10, 10]
        return self._rows

    
class _FooterSettings(_HeaderSettings):
    
    def tail(self):
        """ return the tail row height that's calculated by the page height calculator"""
        return 10


class PageSettings(SettingBase):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._name, self._paper_size, self._margins, self._headers, self._footers, self._blocks = (None, ) * 6
        cfg = getvalue(kwargs, "config")
        if cfg:
            self.config(cfg)
    
    def _calc(self):
        """ calculate the total height/width that a page can hold """
        pass

    def config(self, mp_cfg):
        """
        load settings from a config
        """
        self._headers = _HeaderSettings(getvalue(mp_cfg, "header"))
    
    def headers(self):
        """ return a tuple of header heights, return None if no header is defined """
        return self._headers

    def footers(self):
        """ return a tuple of footer heights, return None if no footer is defined """
        return self._footers
