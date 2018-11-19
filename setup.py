#! /usr/bin/env python3
#################################################################################
#     File Name           :     setup.py
#     Created By          :     Eloi Silva
#     Creation Date       :     [2018-07-25 18:10]
#     Last Modified       :     [2018-11-19 00:19]
#     Description         :      
#################################################################################

from setuptools import setup
import maloy_xls

setup(
    name = 'maloy_xls',
    version = maloy_xls.__version__,
    author = maloy_xls.__author__,
    author_email = maloy_xls.__email__,
    packages = ['maloy_xls',],
    license = 'GNU Version 3',
    install_requires=[
        "xlrd",
        "xlsxwriter",
        ],
)
