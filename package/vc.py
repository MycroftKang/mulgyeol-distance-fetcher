import json
import sys, os

def isfile(filename):
    return os.path.isfile(filename)

def version(v):
    return list(map(int, (v.split("."))))

def isnewupdate(base, last):
    return base[:-1] != last[:-1]

def signature(ver):
    with open('src/version_info.txt', 'rt') as f:
        text = f.read()

    text.replace('<version:tuple>', '{},{},{},0'.format(*ver[:3]))
    text.replace('<version:str>', '{}.{}.{}.{}'.format(*ver))

    with open('src/version_info.txt', 'wt') as f:
        f.write(text)

def build():
    with open('package/info/version.json', 'rt') as f:
        cur = json.load(f)

    cur['version'] += '.{}'.format(os.getenv('CI_COMMIT_SHORT_SHA'))

    os.makedirs('output', exist_ok=True)

    with open('output/last_version.txt', 'wt') as f:
        f.write(cur['version'])
    
    with open('package/info/version.json', 'wt') as f:
        json.dump(cur, f)

    signature(cur['version'])

def release():
    with open('package/info/version.json', 'rt') as f:
        vd = json.load(f)

    with open('public/version.json', 'rt') as f:
        pd = json.load(f)

    vd['version'] += '.{}'.format(os.getenv('CI_COMMIT_SHORT_SHA'))

    pd['last-version'] = vd['version']
    pd['tags'] = os.getenv('CI_COMMIT_TAG')

    os.makedirs('output', exist_ok=True)

    with open('output/last_version.txt', 'wt') as f:
        f.write(vd['version'])

    with open('package/info/version.json', 'wt') as f:
        json.dump(vd, f)

    with open('public/version.json', 'wt') as f:
        json.dump(pd, f)

    signature(vd['version'])

if '-b' in sys.argv:
    build()
elif '-r' in sys.argv:
    release()
