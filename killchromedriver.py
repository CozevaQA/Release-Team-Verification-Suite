import subprocess

def terminate_process(process_name):
    try:
        subprocess.run(['taskkill', '/F', '/IM', process_name], check=True)
        print(f"Terminated process: {process_name}")
    except subprocess.CalledProcessError:
        print(f"Failed to terminate process: {process_name}")

terminate_process("chromedriver.exe")