# -*- coding: UTF-8 -*-
import xlwings as xw


@xw.func
def double_sum(x, y):
    """返回两个参数之和的两倍"""
    return 2 * (x + y)