from datetime import datetime
from git import git


class Commit:
    def __init__(self, hash):
        self.author_name = git(['show', '-s', '--format=%aN', hash])
        self.author_email = git(['show', '-s', '--format=%ae', hash])
        self.subject = git(['show', '-s', '--format=%s', hash])
        self.message = git(['show', '-s', '--format=%b', hash])
        self.author_date = git(['show', '-s', '--format=%at', hash])
        self.author_date = datetime.fromtimestamp(int(self.author_date))
        self.author_date = self.author_date.strftime("%Y-%m-%d %H:%M:%S")
        lines = git(['show', '--name-status', '--format=%n', hash]).strip().splitlines()
        self.files = [line.split()[1] for line in lines]
        self.mods = [line.split()[0] for line in lines]
        self.id = git(['rev-parse', hash])

    def dump(self):
        print(', '.join("%s: %s" % item for item in vars(self).items()))