import subprocess
import os


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
    cmd = [get_program_files_32() + '/Git/bin/git.exe']
    return subprocess.getoutput(cmd + args)
