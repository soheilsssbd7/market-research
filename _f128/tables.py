# -*- coding: utf-8 -*-
from tables_1 import TABLES as _A
from tables_2 import TABLES as _B
from tables_3 import TABLES as _C
from tables_4 import TABLES as _D

TABLES = {}
for _d in (_A, _B, _C, _D):
    TABLES.update(_d)
