import xlwings as xw
import pywintypes
import argparse
import requests
import nltk

__requires__ = 'excel_rw==0.9.0'
__version__ = '0.9.0-2020-8-9'


class Battlefield(xw.Sheet):
    """展示战斗演绎的工作表"""
    pass


class XlUnit:
    """量表中的对象"""

    def __init__(self, uid):
        self.__uid = uid


class UserCommandParser:
    """用户指令"""

    def __init__(self):
        """https://docs.python.org/zh-cn/3/library/argparse.html#prog"""
        self.parser = argparse.ArgumentParser(description='可以在此操作excel或寻求帮助')
        subparsers = self.parser.add_subparsers()
        subparsers.required = True

        parser_a = subparsers.add_parser('push battle', help='add help')
        parser_a.set_defaults(func=push_battle)


def push_battle(sht):
    if not sht:
        sht = start_info()
    url = 'http://192.168.0.230:12123/battle'
    data = {'memberinfos': sht.range('A10').value}
    r = requests.post(url, data)
    sht.range('A11').value = "已提交"
    sht.range('A12').value = nltk.clean_html(r.text)


def operate(value):
    print(value)


def read_command_loop():
    parser = UserCommandParser().parser
    while True:
        command = input()
        args = parser.parse_args(command.split())
        args.func(args)


def start_info():
    try:
        wb = xw.books.active
    except AttributeError:
        wb = xw.Book()
    # wb = xw.Book('Stake4Esports.xlsx')
    try:
        sht = wb.sheets['Info']
    except pywintypes.com_error:
        sht = wb.sheets.add("Info")
    finally:
        print("excel_battlefield run success")
    sht.range('A1').value = [['模块名称', "excel_rw"], ['当前版本', __version__], ['简介', main.__doc__]]
    sht.range('A9').value = "在A10输入阵容后push battle"
    print('模块名称:{0}\n当前版本:{1}'.format('excel_rw', __version__))
    return sht


def main():
    """这个模块通过xlwings直接和EXCEL交互"""
    sht = start_info()
    push_battle(sht)
    #read_command_loop()


if __name__ == '__main__':
    main()
