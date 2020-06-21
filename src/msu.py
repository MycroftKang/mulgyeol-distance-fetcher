import requests
import json
import os, sys
import zipfile
import win32.win32api as win32api

from app.product import PRODUCT_CONFIG
from PyQt5.QtCore import QLockFile

def version(v):
    ls = v.split(".")
    re = list(map(int, (ls[:3])))
    return re

class Updater:
    def __init__(self):
        with open('../info/version.json', 'rt', encoding='UTF8') as f:
            self.info = json.load(f)
        
        self.cur = version(self.info['version'])
        self.cur_commit = self.info['commit']
        self.download_file_name = os.getenv('LOCALAPPDATA')+'\\Temp\\mdf-update\\MDFSetup-stable.zip'
        
    def get_release_info(self):
        try:
            res = requests.get(PRODUCT_CONFIG['VERSION_INFO_URL'])
            data = res.json()
            self.last = version(data['last-version'])
            self.last_commit = data['commit']
            self.tags = data['tags']
        except Exception as e:
            print(e)
            sys.exit()

    def isnew(self):
        self.get_release_info()
        return (self.last[:3] > self.cur[:3]) or (self.last_commit != self.cur_commit)

    def isdownload(self):
        return self.info['update']["download"]

    def isdownloading(self):
        return self.info['update']["isnew"]

    def commit_info(self):
        with open('../info/version.json', 'wt', encoding='UTF8') as f:
            json.dump(self.info, f)

    def download(self):
        self.get_release_info()
        # win32api.ShellExecute(None, "open", "taskkill", '/f /im "MK Bot.exe"', None, 0)
        uinfo = self.info['update']
        uinfo["isnew"] = True
        self.commit_info()
        
        if PRODUCT_CONFIG['UPDATE_URL']:
            r = requests.get(PRODUCT_CONFIG['UPDATE_URL'].format(self.tags))

            os.makedirs(os.path.dirname(self.download_file_name), exist_ok=True)
        
            with open(self.download_file_name, 'wb') as f:
                f.write(r.content)

            _zip = zipfile.ZipFile(self.download_file_name)
            _zip.extractall(os.path.dirname(self.download_file_name))

            uinfo["download"] = True
            self.commit_info()

    def run_update(self, autorun=False):
        if self.isdownload:
            if autorun:
                param = '/S /autorun'
            else:
                param = '/S'
            win32api.ShellExecute(None, "open", os.path.dirname(self.download_file_name)+'\\MDFSetup-stable.exe', param, None, 1)
            sys.exit()

    def check_update(self):
        if (self.isnew()) and (not self.isdownload()):
            self.download()
        
if __name__ == '__main__':

    lockfile = QLockFile(os.getenv('TEMP') + '/MDFMSU.lock')

    if not lockfile.tryLock():
        sys.exit()

    ut = Updater()
    if sys.argv[1] == "/cu":
        ut.check_update()
    elif sys.argv[1] == "/ru":
        ut.run_update()
    elif sys.argv[1] == "/ruwa":
        ut.run_update(True)