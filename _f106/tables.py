# -*- coding: utf-8 -*-
"""ادغام جدول‌های F-106 — ترتیب T1..T32 برابر ترتیب ظهور در سند است."""
from tabs_a import TABLES as A
from tabs_b import TABLES as B
from tabs_c import TABLES as C

TABLES = {}
TABLES.update(A)
TABLES.update(B)
TABLES.update(C)

assert sorted(TABLES, key=lambda k: int(k[1:])) == ['T%d' % i for i in range(1, 33)], sorted(TABLES)
