import os
from sys import platform

def check_system():
    if platform == "linux" or platform == "linux2":
        return "l"
    else :
        return "w"
