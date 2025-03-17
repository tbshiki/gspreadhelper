from setuptools import setup, find_packages

NAME = "gspreadhelper"
VERSION = "0.1.6"

AUTHOR = "tbshiki"
AUTHOR_EMAIL = "info@tbshiki.com"
URL = "https://github.com/tbshiki/" + NAME

setup(
    name=NAME,
    author=AUTHOR,
    author_email=AUTHOR_EMAIL,
    url=URL,
    version=VERSION,
    packages=find_packages(),
)
