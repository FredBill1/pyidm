#!/usr/bin/env python
#coding:utf-8
"""
  Author : Hadi Cahyadi <cumulus13@gmail.com>
  Purpose: command line to download with Internet Download Manager (IDM) on Windows OS
  Created: 04/22/19
  Support: 2.7+, 3+
"""

import os
import sys
from enum import IntEnum


class IDMNotFound(Exception):
    pass

class OSNotSupport(Exception):
    pass

class IDMConfirmOption(IntEnum):
    SHOW_DIALOG = 0
    HIDE_DIALOG = 1
    ADD_TO_QUEUE = 2

if not 'linux' in sys.platform:
    import comtypes.client as cc
    from comtypes import automation
    from pythoncom import CoInitialize
else:
    raise OSNotSupport('This only for Windows OS !')

class IDMan:
    PID = os.getpid()
    
    def __init__(self):
        CoInitialize()
        self.tlb = r'c:\Program Files\Internet Download Manager\idmantypeinfo.tlb'
        if not os.path.isfile(self.tlb):
            self.tlb = r'c:\Program Files (x86)\Internet Download Manager\idmantypeinfo.tlb'
        if not os.path.isfile(self.tlb):
            #print("It seem IDM not installed, please install first !")
            #sys.exit("It seem IDM not installed, please install first !")
            raise IDMNotFound("It seem IDM (Internet Download Manager) not installed, please install first !")

    def get_from_clipboard(self):
        try:
            import clipboard
        except ImportError:
            print("Module 'clipboard' not Installer yet, please install first [pip install clipboard]")
            q = input("Please re-input url download to:")
            if not q:
                sys.exit("You not input URL Download !")
            else:
                return q
        return clipboard.paste()

    #def download(self, link, path_to_save=None, output=None, referrer=None, cookie=None, postData=None, user=None, password=None, confirm = False, lflag = None, clip=False):
    def download(self, link, path_to_save=None, output=None, referrer=None, cookie=None, postData=None, user=None, password=None, confirm=IDMConfirmOption.HIDE_DIALOG, user_agent=None, clip=False):
        
        lflag = confirm
        
        if clip or link == 'c':
            link = self.get_from_clipboard()
        try:
            cc.GetModule(['{ECF21EAB-3AA8-4355-82BE-F777990001DD}', 1, 0])
        except:
            cc.GetModule(self.tlb)

        
        try:
            import comtypes.gen.IDManLib as idman
        except ImportError:
            raise IDMNotFound("Please install 'Internet Download Manager' first !")
        idman1 = cc.CreateObject(idman.CIDMLinkTransmitter, None, None, idman.ICIDMLinkTransmitter2)
        if path_to_save:
            os.path.realpath(path_to_save)
        #if isinstance(postData, dict):
            #postData = '\n'.join([f'{key}: {value}' for key, value in postData.items()])
        if isinstance(cookie, dict):
            cookie = '; '.join([f'{key}={value}' for key, value in cookie.items()])
        
        if isinstance(postData, dict):
            postData = '\n'.join([f'{key}={value}' for key, value in postData.items()])
        
        reserved1 = automation.VARIANT()
        if user_agent:
            reserved1.vt = automation.VT_BSTR
            reserved1.value = user_agent
        else:
            reserved1.vt = automation.VT_EMPTY
        
        # Prepare reserved2 (not used)
        reserved2 = automation.VARIANT()
        reserved2.vt = automation.VT_EMPTY
        
        #idman1.SendLinkToIDM(link, referrer, cookie, postData, user, password, path_to_save, output, lflag)
        idman1.SendLinkToIDM2(link, referrer, cookie, postData, user, password, path_to_save, output, lflag, reserved1, reserved2)

