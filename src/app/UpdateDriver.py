import json
import win32api

class UpdateDriver:
    def __init__(self, _software_info_file):
        self.info_file = _software_info_file
        
    def load_info(self):
        with open(self.info_file, 'rt') as f:
            uinfo = json.load(f)['update']
        return uinfo

    def check_update(self):
        print('check')
        # win32api.ShellExecute(None, "open", '..\\update\\Update.exe', '/cu', None, 0)

    def run_update(self):
        print('run')
        # win32api.ShellExecute(None, "open", '..\\update\\Update.exe', '/ru', None, 0)

    def isnew(self):
        return self.load_info()['isnew']

    def isready(self):
        uinfo = self.load_info()
        return uinfo['isnew'] and uinfo['download']

UDriver = UpdateDriver('../info/version.json')