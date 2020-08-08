import xlwings as xw
import pywintypes

__requires__ = 'excel_rw==0.9.0'
__version__ = '0.9.0-2020-8-9'


class TheBook(object):
    """该模块操作的工作表"""
    pass


class StakeBook(TheBook):
    """记录单位信息的工作表
    """
    pass


class XlUnit(object):
    """量表中的对象"""

    def __init__(self, uid):
        self.__uid = uid


def main():
    """这个模块通过xlwings直接和EXCEL交互"""
    try:
        wb = xw.books.active
    except AttributeError:
        wb = xw.Book()
    # wb = xw.Book('Stake4Esports.xlsx')
    try:
        sht = wb.sheets['Test']
    except pywintypes.com_error:
        sht = wb.sheets.add("Test")
    finally:
        print("excel_battlefield run success")
    sht.range('A1').value = [['模块名称', "excel_rw"], ['当前版本', __version__], ['简介', main.__doc__]]


if __name__ == '__main__':
    main()
