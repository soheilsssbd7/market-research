# -*- coding: utf-8 -*-
"""تجمیع جدول‌ها — پروندهٔ تکادو (وکادو)"""
from tables_1 import T as T1
from tables_2 import T as T2
from tables_3 import T as T3

TABLES = {}
for _d in (T1, T2, T3):
    TABLES.update(_d)
