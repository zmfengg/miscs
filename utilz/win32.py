'''
#! coding=utf-8
@Author:        zmFeng
@Created at:    2019-08-07
@Last Modified: 2019-08-07 8:50:35 am
@Modified by:   zmFeng

handle the so-called win32 junction/reparse issue

finally it's proven that below code randomly failed

# https://stackoverflow.com/questions/27972776/having-trouble-implementing-a-readlink-function 
        # read symbolic link
'''

import ctypes
from ctypes import wintypes
from sys import getdefaultencoding
import platform
from os import environ, walk, path, remove, rmdir
import subprocess as sp

import logging
logger = logging.getLogger()

kernel32 = ctypes.WinDLL('kernel32', use_last_error=True)

FILE_READ_ATTRIBUTES = 0x0080
OPEN_EXISTING = 3
FILE_FLAG_OPEN_REPARSE_POINT = 0x00200000
FILE_FLAG_BACKUP_SEMANTICS   = 0x02000000
FILE_ATTRIBUTE_REPARSE_POINT = 0x0400

IO_REPARSE_TAG_MOUNT_POINT = 0xA0000003
IO_REPARSE_TAG_SYMLINK     = 0xA000000C
FSCTL_GET_REPARSE_POINT    = 0x000900A8
MAXIMUM_REPARSE_DATA_BUFFER_SIZE = 0x4000

LPDWORD = ctypes.POINTER(wintypes.DWORD)
LPWIN32_FIND_DATA = ctypes.POINTER(wintypes.WIN32_FIND_DATAW)
INVALID_HANDLE_VALUE = wintypes.HANDLE(-1).value

def IsReparseTagNameSurrogate(tag):
    return bool(tag & 0x20000000)

def _check_invalid_handle(result, func, args):
    if result == INVALID_HANDLE_VALUE:
        raise ctypes.WinError(ctypes.get_last_error())
    return args

def _check_bool(result, func, args):
    if not result:
        raise ctypes.WinError(ctypes.get_last_error())
    return args

kernel32.FindFirstFileW.errcheck = _check_invalid_handle
kernel32.FindFirstFileW.restype = wintypes.HANDLE
kernel32.FindFirstFileW.argtypes = (
    wintypes.LPCWSTR,  # _In_  lpFileName
    LPWIN32_FIND_DATA) # _Out_ lpFindFileData

kernel32.FindClose.argtypes = (
    wintypes.HANDLE,) # _Inout_ hFindFile

kernel32.CreateFileW.errcheck = _check_invalid_handle
kernel32.CreateFileW.restype = wintypes.HANDLE
kernel32.CreateFileW.argtypes = (
    wintypes.LPCWSTR, # _In_     lpFileName
    wintypes.DWORD,   # _In_     dwDesiredAccess
    wintypes.DWORD,   # _In_     dwShareMode
    wintypes.LPVOID,  # _In_opt_ lpSecurityAttributes
    wintypes.DWORD,   # _In_     dwCreationDisposition
    wintypes.DWORD,   # _In_     dwFlagsAndAttributes
    wintypes.HANDLE)  # _In_opt_ hTemplateFile 

kernel32.CloseHandle.argtypes = (
    wintypes.HANDLE,) # _In_ hObject

kernel32.DeviceIoControl.errcheck = _check_bool
kernel32.DeviceIoControl.argtypes = (
    wintypes.HANDLE,  # _In_        hDevice
    wintypes.DWORD,   # _In_        dwIoControlCode
    wintypes.LPVOID,  # _In_opt_    lpInBuffer
    wintypes.DWORD,   # _In_        nInBufferSize
    wintypes.LPVOID,  # _Out_opt_   lpOutBuffer
    wintypes.DWORD,   # _In_        nOutBufferSize
    LPDWORD,          # _Out_opt_   lpBytesReturned
    wintypes.LPVOID)  # _Inout_opt_ lpOverlapped 

class REPARSE_DATA_BUFFER(ctypes.Structure):
    class ReparseData(ctypes.Union):
        class LinkData(ctypes.Structure):
            _fields_ = (('SubstituteNameOffset', wintypes.USHORT),
                        ('SubstituteNameLength', wintypes.USHORT),
                        ('PrintNameOffset',      wintypes.USHORT),
                        ('PrintNameLength',      wintypes.USHORT))
            @property
            def PrintName(self):
                dt = wintypes.WCHAR * (self.PrintNameLength //
                                       ctypes.sizeof(wintypes.WCHAR))
                name = dt.from_address(ctypes.addressof(self.PathBuffer) +
                                       self.PrintNameOffset).value
                if name.startswith(r'\??'):
                    name = r'\\?' + name[3:] # NT => Windows
                return name
        class SymbolicLinkData(LinkData):
            _fields_ = (('Flags',      wintypes.ULONG),
                        ('PathBuffer', wintypes.BYTE * 0))
        class MountPointData(LinkData):
            _fields_ = (('PathBuffer', wintypes.BYTE * 0),)
        class GenericData(ctypes.Structure):
            _fields_ = (('DataBuffer', wintypes.BYTE * 0),)
        _fields_ = (('SymbolicLinkReparseBuffer', SymbolicLinkData),
                    ('MountPointReparseBuffer',   MountPointData),
                    ('GenericReparseBuffer',      GenericData))
    _fields_ = (('ReparseTag',        wintypes.ULONG),
                ('ReparseDataLength', wintypes.USHORT),
                ('Reserved',          wintypes.USHORT),
                ('ReparseData',       ReparseData))
    _anonymous_ = ('ReparseData',)

def islink(path):
    ''' check if given path is a junction
    '''
    data = wintypes.WIN32_FIND_DATAW()
    kernel32.FindClose(kernel32.FindFirstFileW(path, ctypes.byref(data)))
    if not data.dwFileAttributes & FILE_ATTRIBUTE_REPARSE_POINT:
        return False
    return IsReparseTagNameSurrogate(data.dwReserved0)

def readlink(path):
    ''' if given path is a junction, get the actual junction point
    '''
    n = wintypes.DWORD()
    buf = (wintypes.BYTE * MAXIMUM_REPARSE_DATA_BUFFER_SIZE)()
    flags = FILE_FLAG_OPEN_REPARSE_POINT | FILE_FLAG_BACKUP_SEMANTICS
    handle = kernel32.CreateFileW(path, FILE_READ_ATTRIBUTES, 0, None,
                OPEN_EXISTING, flags, None)
    try:
        kernel32.DeviceIoControl(handle, FSCTL_GET_REPARSE_POINT, None, 0,
            buf, ctypes.sizeof(buf), ctypes.byref(n), None)
    finally:
        kernel32.CloseHandle(handle)
    rb = REPARSE_DATA_BUFFER.from_buffer(buf)
    tag = rb.ReparseTag
    if tag == IO_REPARSE_TAG_SYMLINK:
        return rb.SymbolicLinkReparseBuffer.PrintName
    if tag == IO_REPARSE_TAG_MOUNT_POINT:
        return rb.MountPointReparseBuffer.PrintName
    if not IsReparseTagNameSurrogate(tag):
        raise ValueError("not a link")
    raise ValueError("unsupported reparse tag: %d" % tag)

_w_ver = int(platform.version().split('.')[0])

def _safe_get(fldr):
    if _w_ver < 6:
        return fldr
    try:
        # readlink is not working properly
        return readlink(fldr) if islink(fldr) else fldr
    except:
        return fldr


def users(cat=None):
    ''' return the users of current system. Support windows only

    Args:

        cat=None:   None for user name
                    home for user home folder
                    temp for user temp folders
    '''

    cmd = "net users" if platform.system().find('Windows') == 0 else None
    if cmd:
        with sp.Popen(cmd, stdout=sp.PIPE) as proc:
            lns = proc.stdout.read().decode(getdefaultencoding(), errors='ignore').split("\r\n")
        lns = [x for x in lns if x][2:-1]
        usrs = [y for x in lns for y in x.split() if y]
    else:
        usrs = None
    if usrs and cat:
        root = environ['USERPROFILE']
        root = path.dirname(path.abspath(root)) # assume root for all users are of the same
        usrs = [path.join(root, x) for x in usrs]
        if cat == 'temp':
            def _sfldrs(fldr):
                fldr = _safe_get(fldr)
                return next(iter(walk(fldr)))[1]
            mp = {}
            for root in usrs:
                try:
                    sf = _sfldrs(root)
                    sf = [x for x in sf if x.lower().find('local') >= 0]
                    if sf:
                        lst = []
                        for rt in sf:
                            rt = _safe_get(path.join(root, rt))
                            lst.extend([_safe_get(path.join(rt, x)) for x in _sfldrs(rt) if x.lower().find('temp') == 0])
                    mp[path.basename(root)] = lst
                except:
                    pass
            usrs = mp
    return usrs

def clearTempFiles(excl=None):
    ''' delete the temp folder/files
    '''
    def _safe_kill(fn, func=None):
        try:
            if not func:
                func = remove
            func(fn)
        except:
            pass
    if not excl:
        excl = set()
    mp = users('temp')
    for u, fldrs in mp.items():
        if u in excl:
            continue
        pidx, lst = 0, [x for x in fldrs]
        fns = []
        while pidx < len(lst):
            fldr = _safe_get(lst[pidx])
            for x, sfs, sfns in walk(fldr):
                sfs = [path.join(x, y) for y in sfs]
                lst.extend(sfs)
                fldrs.extend(sfs)
                sfns = [path.join(x, y) for y in sfns]
                for fn in sfns:
                    _safe_kill(fn)
                fns.extend(sfns)
            pidx += 1
        for fldr in lst[2:]:
            _safe_kill(fldr, rmdir)
        logger.debug('(%s): folder_count=%d, file_count=%d' % (u, len(lst), len(fns)))
