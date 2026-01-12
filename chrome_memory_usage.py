import pandas as pd
import subprocess
from io import StringIO
from time import sleep


def get_total_chrome_usage():
    all_subprocesses = subprocess.check_output('tasklist')
    subprocess_data = StringIO(all_subprocesses.decode('utf-8'))
    df = pd.read_csv(subprocess_data, sep='\s\s+', skiprows=[2], engine='python')
    chrome_usages = df[df['Image Name'] == 'chrome.exe']

    def parse_mem(s):
        return int(s[:-2].replace(',', '')) / 1e3

    # Transform the 'Mem Usage' column of strings to a column of floats
    mem = chrome_usages['Mem Usage'].transform(parse_mem)
    # Sum to get Chrome's total memory usage
    return round((mem.sum()/ 1e3), 2)

print(get_total_chrome_usage())


