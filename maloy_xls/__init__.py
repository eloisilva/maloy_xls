#! /usr/bin/env python3
#################################################################################
#     File Name           :     maloy_xls/__init__.py
#     Created By          :     Eloi Silva (eloi@how2security.com.br)
#     Creation Date       :     [2018-07-25 18:22]
#     Last Modified       :     [2018-11-19 00:25]
#     Description         :      
#################################################################################

import sys

__version__ = '1.0.1-dev1'
__author__ = 'Eloi Luiz da Silva'
__email__ = 'eloi@how2security.com.br'

if sys.version_info[0] < 3:
    msg = 'Python Version 3 required'
    raise ImportError(msg)

from .xls import xls_to_list
from .xls import sheet_names
from .xls import list_to_xls
from .xls import xls_rows
from .xls import xls_cols
from .xls import xls_clean
from .xls import calc_integer_to_row_position
