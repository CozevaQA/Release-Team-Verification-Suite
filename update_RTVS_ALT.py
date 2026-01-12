import subprocess
import os

def run_command(command, cwd):
    try:
        subprocess.run(command, check=True, cwd=cwd, shell=True)
    except subprocess.CalledProcessError as e:
        print(f"Error occurred while executing {command}: {e}")
        input("Press any key to exit...")
        exit(1)

cwd = os.getcwd()

# Run git stash and git pull with error handling
run_command(["git", "stash"], cwd)
run_command(["git", "pull"], cwd)

# Wait for user input before closing the window
input("Operation completed. Press any key to exit...")
