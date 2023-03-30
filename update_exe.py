import os
import subprocess

# Get the current working directory
cwd = os.getcwd()

# Execute the git pull command in the current working directory
subprocess.call(f"cd {cwd} && git pull", shell=True)
