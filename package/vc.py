import json
import sys
import os
import requests


def isfile(filename):
    return os.path.isfile(filename)


def version(v):
    ls = v.split(".")
    re = list(map(int, (ls[:3])))
    return re


def isnewupdate(base, last):
    return base[:-1] != last[:-1]


def signature(ver):
    with open('product.json', 'rt') as f:
        PRODUCT_CONFIG = json.load(f)

    ver = version(ver)

    with open('src/version_info.txt', 'rt') as f:
        text = f.read()

    text = text.replace('<version:tuple>', '{},{},{},0'.format(*ver[:3]))
    text = text.replace('<version:str>', '{}.{}.{}'.format(*ver))
    text = text.replace('<name:str>', PRODUCT_CONFIG['PRODUCT_NAME'])

    with open('src/version_info.txt', 'wt') as f:
        f.write(text)


def build():
    with open('package/info/version.json', 'rt') as f:
        cur = json.load(f)

    cur['version'] = cur['version'].replace('.beta', '')
    cur['commit'] = os.getenv('CI_COMMIT_SHORT_SHA')

    os.makedirs('output', exist_ok=True)

    with open('output/last_version.txt', 'wt') as f:
        f.write(cur['version'])

    with open('package/info/version.json', 'wt') as f:
        json.dump(cur, f)

    signature(cur['version'])


def local_build():
    with open('package/info/version.json', 'rt') as f:
        cur = json.load(f)

    signature(cur['version'])


def release():
    with open('package/info/version.json', 'rt') as f:
        vd = json.load(f)

    vd['version'] = vd['version'].replace('.beta', '')
    vd['commit'] = os.getenv('CI_COMMIT_SHORT_SHA')

    if '-rc' in os.getenv('CI_COMMIT_TAG'):
        r = requests.get(os.getenv('VESION_URL'))
        pd = r.json()
        pd['insider'] = {"commit": vd['commit'],
                         "tags": os.getenv('CI_COMMIT_TAG')}
    else:
        with open('public/version.json', 'rt') as f:
            pd = json.load(f)
        pd['last-version'] = vd['version']
        pd['commit'] = os.getenv('CI_COMMIT_SHORT_SHA')
        pd['tags'] = os.getenv('CI_COMMIT_TAG')

    os.makedirs('output', exist_ok=True)

    with open('output/last_version.txt', 'wt') as f:
        f.write(vd['version'])

    with open('package/info/version.json', 'wt') as f:
        json.dump(vd, f)

    with open('public/version.json', 'wt') as f:
        json.dump(pd, f)

    signature(vd['version'])


if os.getenv('CI_COMMIT_BRANCH'):
    build()
elif os.getenv('CI_COMMIT_TAG'):
    release()
elif '-lb' in sys.argv:
    local_build()
