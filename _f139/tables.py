# -*- coding: utf-8 -*-
"""جدول‌های سند F-138 — تولید برق عسلویه مپنا (بمپنا)"""
from tables_1 import T as _A
from tables_2 import T as _B
from tables_3 import T as _C
from tables_4 import T as _D
from tables_5 import T as _E

TABLES = {}
for _d in (_A, _B, _C, _D, _E):
    TABLES.update(_d)
