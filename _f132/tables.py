# -*- coding: utf-8 -*-
from tables_1 import TABLES as A
from tables_2 import TABLES as B
from tables_3 import TABLES as C
from tables_4 import TABLES as D

TABLES = {}
for d in (A, B, C, D):
    TABLES.update(d)
