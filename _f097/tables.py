# -*- coding: utf-8 -*-
from tabs_a import TABLES as A
from tabs_b import TABLES as B
from tabs_c import TABLES as C

TABLES = {}
for src in (A, B, C):
    TABLES.update(src)
