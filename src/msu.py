import json
import os
import sys
import zipfile

import requests
import win32.win32api as win32api
from PyQt5.QtCore import QLockFile

from app.product import PRODUCT_CONFIG


def version(v):
    ls = v.split(".")
    re = list(map(int, (ls[:3])))
    return re


class Updater:
    def __init__(self):
        with open('../info/version.json', 'rt', encoding='UTF8') as f:
            self.info = json.load(f)

        with open('../data/settings/settings.json', 'rt', encoding='UTF8') as f:
            self.insider = json.load(f).get('insiderupdate', False)

        self.data = None
        self.cur = version(self.info['version'])
        self.cur_commit = self.info['commit']
        self.download_file_name = os.getenv(
            'LOCALAPPDATA')+'\\Temp\\mdf-update\\MDFSetup.zip'

    def get_release_info(self):
        if self.data == None:
            try:
                res = requests.get(PRODUCT_CONFIG['VERSION_INFO_URL'])
                res.raise_for_status()
                data = res.json()
                self.last = version(data['last-version'])
                self.last_commit = data['commit']
                self.tags = data['tags']
                self.beta_commit = data['insider']['commit']
                self.beta_tags = data['insider']['tags']
            except Exception as e:
                print(e)
                sys.exit()

    def isnew(self):
        self.get_release_info()
        return (self.last[:3] > self.cur[:3]) or (self.last_commit != self.cur_commit)

    def isbeta(self):
        self.get_release_info()
        return (self.beta_commit != None) and (self.beta_commit != self.cur_commit)

    def isdownload(self):
        return self.info['update']["download"]

    def commit_info(self):
        with open('../info/version.json', 'wt', encoding='UTF8') as f:
            json.dump(self.info, f)

    def download(self, beta=False):
        if beta:
            utype = 'stable-release'
            ref = self.beta_tags
            out = 'MDFSetup-stable.exe'
        else:
            utype = 'stable-release'
            ref = self.tags
            out = 'MDFSetup-stable.exe'

        uinfo = self.info['update']
        uinfo["isnew"] = True
        uinfo['isbeta'] = beta
        self.commit_info()

        if PRODUCT_CONFIG['UPDATE_URL']:
            try:
                r = requests.get(
                    PRODUCT_CONFIG['UPDATE_URL'].format(ref, utype))
                r.raise_for_status()
            except:
                sys.exit()

            os.makedirs(os.path.dirname(
                self.download_file_name), exist_ok=True)

            with open(self.download_file_name, 'wb') as f:
                f.write(r.content)

            _zip = zipfile.ZipFile(self.download_file_name)
            _zip.extractall(os.path.dirname(self.download_file_name))

            win32api.ShellExecute(None, "open", os.path.dirname(
                self.download_file_name)+'\\{}'.format(out), '/S /init', None, 1)

            uinfo["download"] = True
            self.commit_info()

    def run_update(self, autorun=False):
        if self.isdownload:
            if autorun:
                param = '/S /autorun'
            else:
                param = '/S'

            if self.info['update']['isbeta']:
                out = 'MDFSetup-beta.exe'
            else:
                out = 'MDFSetup-stable.exe'

            win32api.ShellExecute(None, "open", os.path.dirname(
                self.download_file_name)+'\\{}'.format(out), param, None, 1)
            sys.exit()

    def check_update(self):
        if self.isnew():
            self.download()
        elif self.insider and self.isbeta():
            self.download(True)


if __name__ == '__main__':
    lockfile = QLockFile(os.getenv('TEMP') + '/MDFMSU.lock')
    if not lockfile.tryLock():
        sys.exit()
    ut = Updater()
    if len(sys.argv) > 0:
        if sys.argv[1] == "/cu":
            ut.check_update()
        elif sys.argv[1] == "/ru":
            ut.run_update()
        elif sys.argv[1] == "/ruwa":
            ut.run_update(True)
