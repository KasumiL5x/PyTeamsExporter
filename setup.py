import sys
import subprocess

required_major = 3
required_minor = 7

# Version check.
py_info = sys.version_info
if py_info.major < required_major or py_info.minor < required_minor:
    print('Python must be at least ' + str(required_major) + '.' + str(required_minor) + '.')
    print('Please install at least this version and try again.')
    exit()

# https://stackoverflow.com/a/27496113
reqs = subprocess.check_output([sys.executable, '-m', 'pip', 'freeze'])
installed_packages = [r.decode().split('==')[0].lower() for r in reqs.split()]

# Tech stack check.
libs = [
    'flask',
    'requests-oauthlib'
]
for lib in libs:
    if lib not in installed_packages:
        print('You are missing the `'+  lib + '` library.')
        print('Run `pip install ' + lib + '` to install it and try again.')
        exit()

print('You\'re good to go! Run `python ./app.py` from this directory!')