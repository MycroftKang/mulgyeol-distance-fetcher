import json
import sys
import win32.win32api as win32api

class UpdateDriver:
    def __init__(self, _software_info_file):
        self.info_file = _software_info_file
        
    def load_info(self):
        with open(self.info_file, 'rt') as f:
            uinfo = json.load(f)['update']
        return uinfo

    def check_update(self):
        if getattr(sys, 'frozen', False):
            win32api.ShellExecute(None, "open", '..\\update\\Update.exe', '/cu', None, 1)
        else:
            print('Call Check_Update')
        
    def run_update(self):
        if getattr(sys, 'frozen', False):
            win32api.ShellExecute(None, "open", '..\\update\\Update.exe', '/ru', None, 1)
        else:
            print('Call Run_Update')
        
    def run_update_with_autorun(self):
        if getattr(sys, 'frozen', False):
            win32api.ShellExecute(None, "open", '..\\update\\Update.exe', '/ruwa', None, 1)
        else:
            print('Call Run_Update_With_Autorun')

    def isnew(self):
        return self.load_info()['isnew']

    def isready(self):
        uinfo = self.load_info()
        return uinfo['isnew'] and uinfo['download']

UDriver = UpdateDriver('../info/version.json')