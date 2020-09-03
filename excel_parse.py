# -*- coding: UTF-8 -*-
import argparse
import report_parse
import xlwings as xw
import excel_battlefield as xlbf


class UserCommandParser:
    """用户指令"""

    def __init__(self):
        """https://docs.python.org/zh-cn/3/library/argparse.html#prog"""
        self.parser = argparse.ArgumentParser(description='可以在此操作excel或寻求帮助')
        subparsers = self.parser.add_subparsers()
        subparsers.required = True

        parser_a = subparsers.add_parser('push battle', help='add help')
        parser_a.set_defaults(func=xlbf.push_battle)


def read_command_loop():
    parser = UserCommandParser().parser
    while True:
        command = input()
        args = parser.parse_args(command.split())
        args.func(args)
