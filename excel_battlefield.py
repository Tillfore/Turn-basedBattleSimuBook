import xlwings as xw
import pywintypes
import argparse

__requires__ = 'excel_rw==0.9.0'
__version__ = '0.9.0-2020-8-9'


class Battlefield(xw.Sheet):
    """展示战斗演绎的工作表"""
    pass


class XlUnit:
    """量表中的对象"""

    def __init__(self, uid):
        self.__uid = uid


class UserCommandParser():
    """用户指令"""

    def __init__(self):
        """https://docs.python.org/zh-cn/3/library/argparse.html#prog"""
        self.parser = argparse.ArgumentParser(description='可以在此操作excel或寻求帮助')
        subparsers = self.parser.add_subparsers()
        subparsers.required = True

        parser_a = subparsers.add_parser('add', help='add help')
        parser_a.add_argument('-x', type=int, default=1, help='x value')
        parser_a.add_argument('-y', type=int, default=1, help='y value')
        parser_a.add_argument('-z', action='store_const', default=1, const=2, help='y value')
        parser_a.set_defaults(func=self.add)

    def add(self, args):
        r = (args.x + args.y)*args.z
        print('x + y = ', r)


def operate(value):
    print(value)


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
    print('模块名称:{0}\n当前版本:{1}'.format('excel_rw',  __version__))

    parser = UserCommandParser().parser
    while True:
        command = input()
        args = parser.parse_args(command.split())
        args.func(args)


if __name__ == '__main__':
    main()
