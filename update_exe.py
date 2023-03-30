import os
import subprocess

# Get the current working directory
cwd = os.getcwd()

# Open a new terminal console and run the git pull command
subprocess.Popen(["gnome-terminal", "--", "bash", "-c", f"cd {cwd} && git pull; exec bash"], stdin=subprocess.PIPE)

