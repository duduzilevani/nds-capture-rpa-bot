#configurations.py
import configparser
import os

PATH = os.getcwd()

class Configurations:
    def __init__(self, section, path_variable):
        self.section = section
        self.path_variable = path_variable

    def read_config(self):
        config = configparser.RawConfigParser(allow_no_value=True)
        config.read(os.path.join(PATH, "config.ini"))
        return config.get(self.section, self.path_variable)
