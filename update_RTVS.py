import os
import subprocess

import subprocess
import os

cwd = os.getcwd()
subprocess.run(["git", "pull"], check=True, cwd=cwd, shell=True)


