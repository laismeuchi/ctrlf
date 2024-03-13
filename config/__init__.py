import json
import pathlib

CONFIG_FILE = "config.json"
CONFIG_PATH = pathlib.Path(__file__).parent.absolute() / CONFIG_FILE

def get_config():
    return json.load(open(CONFIG_PATH, encoding='utf-8'))