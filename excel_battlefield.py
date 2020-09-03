import os
import xlwings as xw
import pywintypes
import requests
from bs4 import BeautifulSoup
import time
import report_parse
import json

__requires__ = 'excel_rw==0.9.0'
__version__ = '0.9.1-2020-9-3'
__run_time = 0
BATTLE_REPORTS_PATH = 'E:/Projects/BattleReports/'
BATTLE_URL = 'http://192.168.0.230:12123/battle'
BOOK_NAME = 'excel_battlefield.xlsx'
SHEET_TEST = 'Test'
SHEET_INFO = 'Info'
SHEET_CONTENTS = 'Contents'
battleInput = ''


def push_battle(shts, run_time=1, simple_parse=0):
    if not shts:
        shts = start_work()
    url = BATTLE_URL
    sht = shts[SHEET_INFO]
    data = {'memberinfos': battleInput}
    run_time = min(run_time, 100)
    n = 0
    n_sp = 1
    while n < run_time:
        n += 1
        sht.cells(10, 1).value = '提交中({0}/{1})'.format(n, run_time)
        r = requests.post(url, data)
        soup = BeautifulSoup(r.text, 'html.parser')
        json_data = delete_battle_report_head(soup.get_text(), battleInput)
        sht.cells(10 + n, 1).value = save_as_json(json_data)
        if simple_parse:
            while sht.cells(n_sp+10, 6).value:
                n_sp += 1
            rp = report_parse.BattleReport(json.loads(json_data, strict=False))
            shts[SHEET_INFO].range((n_sp+10, 6)).value = \
                [n_sp, rp.win, len(rp.rounds), len(rp.heroes_final), sht.cells(10 + n, 1).value]
        time.sleep(1)
    sht.cells(10, 1).value = '已提交'


def delete_battle_report_head(text, *args):
    head_string = {'\n', '战斗测试', '上传新表', '显示公式'}
    for s in args:
        head_string.add(s)
    for hs in head_string:
        text = text.replace(hs, '', 1)
    return text.replace(u'\xa0', '').replace(u'&nbsp', '')


def save_as_json(text, file_name='', path=''):
    if path == '':
        path = time.strftime('%Y%m%d') + '/'
    path = BATTLE_REPORTS_PATH + path
    if file_name == '':
        file_name = time.strftime('%Y%m%d%H%M%S') + '.json'
    if not os.path.exists(path):
        os.makedirs(path)
    with open(path + file_name, 'a+', encoding='utf-8') as f:
        f.write(text)
        f.close()
    return file_name


def start_work():
    wb = xw.Book(BOOK_NAME)
    try:
        wb.sheets[SHEET_INFO]
    except pywintypes.com_error:
        wb.sheets.add(SHEET_INFO)
    finally:
        shts = wb.sheets
        print('excel_battlefield run success')
    # shts[SHEET_INFO].range('A:E').clear()
    shts[SHEET_INFO].range((1, 1)).value = [['模块名称', 'excel_rw'], ['当前版本', __version__], ['简介', main.__doc__]]
    print('模块名称:{0}\n当前版本:{1}'.format('excel_rw', __version__))
    fresh_contents(shts)
    return shts


def fresh_contents(shts):
    # 根据Sheets'Contents'创建索引
    s_key = shts[SHEET_CONTENTS].range('B1:B3')
    s_value = shts[SHEET_CONTENTS].range('C1:C3')
    s_dict = dict(zip(s_key.value, s_value.value))
    global battleInput
    battleInput = s_dict["battleInput"]


def main():
    """这个模块通过xlwings直接和EXCEL交互"""
    shts = start_work()
    user_command = shts[SHEET_TEST].range('A1:F1')
    uc = user_command.value
    # 0 跑战报 参数1:次数
    if uc[0] == 0:
        push_battle(shts, uc[1])
    if uc[0] == 1:
        push_battle(shts, uc[1], uc[2])
    # 2 模拟战斗 参数1:战报
    if uc[0] == 2:
        pass


if __name__ == '__main__':
    main()
