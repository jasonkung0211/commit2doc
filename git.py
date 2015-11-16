import subprocess
import os
import codecs

import sys

def is_64_windows():
    return 'PROGRAMFILES(X86)' in os.environ


def get_program_files_32():
    if is_64_windows():
        return os.environ['PROGRAMFILES(X86)']
    else:
        return os.environ['PROGRAMFILES']


def get_program_files_64():
    if is_64_windows():
        return os.environ['PROGRAMW6432']
    else:
        return None


def git(args=[]):
    if os.path.isfile(get_program_files_64() + '/Git/bin/git.exe'):
        cmd = [get_program_files_64() + '/Git/bin/git.exe']
    elif os.path.isfile(get_program_files_32() + '/Git/bin/git.exe'):
        cmd = [get_program_files_32() + '/Git/bin/git.exe']
    else:
        print('git.exe is missing')
        sys.exit()
    return codecs.decode(subprocess.check_output(cmd + args), 'utf-8')
