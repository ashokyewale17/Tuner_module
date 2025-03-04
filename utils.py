import os
import sys

def get_script_dir():
    return sys.path[0]

def get_json_file_path(filename):
    return os.path.join(getattr(sys, '_MEIPASS', get_script_dir()), filename)