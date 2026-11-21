# -*- coding: utf-8 -*-
"""جدول‌های سند F-134 — تولید نیروی برق دماوند (دماوند)"""
from tables_1 import T as _A
from tables_2 import T as _B
from tables_3 import T as _C
from tables_4 import T as _D

TABLES = {}
for _d in (_A, _B, _C, _D):
    TABLES.update(_d)
