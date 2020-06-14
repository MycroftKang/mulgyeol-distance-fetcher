import requests
import json
import os, sys
import zipfile
import win32api

def version(v):
    return list(map(int, (v.split("."))))

class Updater:
    def __init__(self):
        with open('../info/version.json', 'rt') as f:
            self.info = json.load(f)
        
        self.cur = version(self.info['version'])
        self.download_file_name = os.getenv('LOCALAPPDATA')+'\\Temp\\mdf-update\\MDFSetup-stable.zip'
        
        try:
            res = requests.get('https://mgylabs.gitlab.io/mulgyeol-distance-fetcher/version.json')
            data = res.json()
            self.last = version(data['last-version'])
            self.tags = data['tags']
        except:
            pass

    def isnew(self):
        return (self.last[:3] > self.cur[:3]) or (self.last[3] != self.cur[3])

    def isdownload(self):
        return self.info['update']["download"]

    def commit_info(self):
        with open('../info/version.json', 'wt') as f:
            json.dump(self.info, f)

    def download(self):
        # win32api.ShellExecute(None, "open", "taskkill", '/f /im "MK Bot.exe"', None, 0)
        uinfo = self.info['update']
        uinfo["isnew"] = True
        self.commit_info()
        
        r = requests.get('https://gitlab.com/mgylabs/mulgyeol-distance-fetcher/-/jobs/artifacts/{}/download?job=stable-release'.format(self.tags))
        
        with open(self.download_file_name, 'wb') as f:
            f.write(r.content)

        _zip = zipfile.ZipFile(self.download_file_name)
        _zip.extractall(os.path.dirname(self.download_file_name))

        uinfo["download"] = True
        self.commit_info()

    def run_update(self):
        if self.isdownload:
            win32api.ShellExecute(None, "open", os.path.dirname(self.download_file_name)+'\\MDFSetup-stable.exe', '/S', None, 0)

    def check_update(self):
        if self.isnew():
            self.download()
        
if __name__ == '__main__':
    ut = Updater()
    if sys.argv[1] == "/cu":
        ut.check_update()
    elif sys.argv[1] == "/ru":
        ut.run_update()