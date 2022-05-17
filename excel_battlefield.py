import os
import xlwings as xw
import pywintypes
import requests
from bs4 import BeautifulSoup
import time
import report_parse
import json
import configparser

__requires__ = 'excel_rw==0.9.3.1'
__version__ = '0.9.3.1-2021-12-2'
__run_time = 0
battleInput = ''
battleInputAttacker = ''
battleInputDefenser = ''


def take_config(config_key, default, value_type=''):
    config = configparser.ConfigParser()
    config.read('cfg.ini', encoding='UTF-8')
    if config_key == 'TOOL_DEBUG':
        try:
            print('TOOL_DEBUG = ', config.getboolean('Config', config_key))
            return config.getboolean('Config', config_key)
        except configparser.NoOptionError:
            return False
        except configparser.NoSectionError:
            return False
    try:
        if value_type == 'int':
            if TOOL_DEBUG:
                print(config_key, ' = ', config.getint('Config', config_key))
            return config.getint('Config', config_key)
        elif value_type == 'boolean':
            if TOOL_DEBUG:
                print(config_key, ' = ', config.getboolean('Config', config_key))
            return config.getboolean('Config', config_key)
        elif value_type == 'list':
            if TOOL_DEBUG:
                print(config_key, ' = ', list(config['Config'][config_key].split(',')))
            return list(config['Config'][config_key].split(','))
        else:
            if TOOL_DEBUG:
                print(config_key, ' = ', config.get('Config', config_key))
            return config.get('Config', config_key)
    except configparser.NoOptionError:
        if TOOL_DEBUG:
            print(config_key, ' 读取默认配置')
        return default
    except configparser.NoSectionError:
        return default
    except KeyError:
        if TOOL_DEBUG:
            print(config_key, ' 读取默认配置')
        return default


TOOL_DEBUG = take_config('TOOL_DEBUG', False, 'boolean')
BATTLE_REPORTS_PATH = take_config('BATTLE_REPORTS_PATH', 'E:/Projects/BattleReports/')
BATTLE_URL = take_config('BATTLE_URL', 'http://192.168.0.89:12126/battle')
BATTLE_PARSE_LEN = take_config('BATTLE_PARSE_LEN', 65, 'int')
BOOK_NAME = take_config('BOOK_NAME', 'excel_battlefield.xlsx')
SHEET_TEST = take_config('SHEET_TEST', 'Test')
SHEET_INFO = take_config('SHEET_INFO', 'Info')
SHEET_CONTENTS = take_config('SHEET_CONTENTS', 'Contents')
SHEET_TEAM_CREATOR = take_config('SHEET_CONTENTS', 'TeamCreator')
TRIAL_LEVEL_START = take_config('TRIAL_LEVEL_START', 7000, 'int')
TRIAL_LEVEL_END = take_config('TRIAL_LEVEL_END', 8000, 'int')
TRIAL_ROUND_LIMIT = take_config('TRIAL_ROUND_LIMIT', 8, 'int')
GUILD_LEVEL_START = take_config('GUILD_LEVEL_START', 7000, 'int')
GUILD_LEVEL_END = take_config('GUILD_LEVEL_END', 8000, 'int')
GUILD_ROUND_LIMIT = take_config('GUILD_ROUND_LIMIT', 8, 'int')
HEAD_STRING = take_config('HEAD_STRING', [], 'list')


def push_battle(shts, run_time=1, simple_parse=0, show_debug=0, max_round=0, monster_group=0, max_times=1):
    if not shts:
        shts = start_work()
    fresh_contents(shts)
    url = BATTLE_URL
    sht = shts[SHEET_INFO]
    run_time = min(run_time, 100)
    try:
        # 可能手动配置快跑起始行
        n_sp = sht.cells(9, 7).value + 0
    except TypeError:
        n_sp = 0
    n = 0
    print('开始自动战斗中……')
    while sht.cells(n_sp + 10, 6).value:
        n_sp += 1
    take_parse_title(sht, n_sp)
    n_sp += 1
    sht.cells(10, 1).value = '提交中'
    while n < run_time:
        n += 1
        br_row = 10 + n
        # 请求转为multipart/form-data格式
        data = {'memberinfos': battleInput, 'showDebugFormula': show_debug, 'maxRound': max_round,
                'monsterGroup': monster_group, 'maxTimes': max_times}
        r = requests.post(url, data=data)
        soup = BeautifulSoup(r.text, 'html.parser')
        json_data = delete_battle_report_head(soup.get_text(), battleInput, *HEAD_STRING)
        sht.cells(br_row, 1).value = save_as_json(json_data)
        if simple_parse:
            take_simple_parse(sht, simple_parse, json_data, br_row, n_sp, n)
            n_sp += 1
    # time.sleep(1)
    sht.cells(10, 1).value = '已提交'


def rush_pve_battle(shts, run_time=1, simple_parse=0, show_debug=0, max_round=0, start_row=10, try_sl=0, max_times=1):
    if not shts:
        shts = start_work()
    fresh_contents(shts)
    url = BATTLE_URL
    sht = shts[SHEET_INFO]
    # 请求转为multipart/form-data格式
    run_time = min(run_time, 100)
    try:
        n_sp = sht.cells(9, 7).value + 0
    except TypeError:
        n_sp = 0
    n = 0
    print('开始批量PVE战斗中……')
    while sht.cells(n_sp + 10, 6).value:
        n_sp += 1
    take_parse_title(sht, n_sp)
    n_sp += 1
    monster_groups = shts[SHEET_TEST].range('BL' + str(start_row) + ':BL9999').value
    groups_count = len(monster_groups)
    sht.cells(10, 1).value = '提交中'
    while n < groups_count:
        if not monster_groups[n]:
            sht.cells(10, 1).value = '已提交'
            return
        mg = int(monster_groups[n])
        if TOOL_DEBUG:
            print(mg)
        # 材料本特殊处理
        if TRIAL_LEVEL_START <= mg <= TRIAL_LEVEL_END:
            max_round = 8
        # 公会副本特殊处理
        elif GUILD_LEVEL_START <= mg <= GUILD_LEVEL_END:
            max_round = 8
        else:
            max_round = 0
        data = {'memberinfos': battleInput, 'showDebugFormula': show_debug, 'maxRound': max_round,
                'monsterGroup': mg, 'maxTimes': max_times}
        n = n + 1
        i = 0
        sl_time = 0
        while i < run_time:
            br_row = 11 + i
            r = requests.post(url, data=data)
            soup = BeautifulSoup(r.text, 'html.parser')
            json_data = delete_battle_report_head(soup.get_text(), battleInput, *HEAD_STRING)
            sht.cells(br_row, 1).value = save_as_json(json_data, custom=str(mg))
            if simple_parse:
                take_simple_parse(sht, simple_parse, json_data, br_row, n_sp, mg)
                n_sp += 1
            i += 1
            if try_sl >= sl_time:
                if sht.range((n_sp + 10, 7)).value == 'False':
                    sl_time += 1
                    i = 0
            # time.sleep(1)
    sht.cells(10, 1).value = '已提交'


def rush_pvp_battle(shts, run_time=1, simple_parse=0, show_debug=0, start_row=10, max_times=1):
    if not shts:
        shts = start_work()
    fresh_contents(shts)
    url = BATTLE_URL
    sht = shts[SHEET_INFO]
    # 请求转为multipart/form-data格式
    run_time = min(run_time, 1000)
    try:
        n_sp = sht.cells(9, 7).value + 0
    except TypeError:
        n_sp = 0
    m = 0
    print('开始批量PVP战斗中……')
    attacker_title = shts[SHEET_TEST].range('BP' + str(start_row) + ':BP9999').value
    attacker_input = shts[SHEET_TEST].range('BQ' + str(start_row) + ':BQ9999').value
    attacker_groups_count = len(attacker_input)
    while m < attacker_groups_count:
        if not attacker_input[m]:
            sht.cells(10, 1).value = '已提交'
            return
        ai = attacker_input[m]
        n = 0
        while sht.cells(n_sp + 10, 6).value:
            n_sp += 1
        sht.range((n_sp + 10, 6)).value = time.strftime('%d%H%M')
        defenser_title = shts[SHEET_TEST].range('BR' + str(start_row) + ':BR9999').value
        defenser_input = shts[SHEET_TEST].range('BS' + str(start_row) + ':BS9999').value
        defenser_groups_count = len(defenser_input)
        sht.cells(10, 1).value = '提交中'
        while n < defenser_groups_count:
            if not defenser_input[n]:
                break
            di = defenser_input[n]
            if simple_parse < 2:
                sht.range((n_sp + 10, 7)).value = attacker_title[m]
                sht.range((n_sp + 10, 8)).value = defenser_title[n]
            else:
                sht.range((n_sp + 10, 13)).value = attacker_title[m]
                sht.range((n_sp + 10, 21)).value = defenser_title[n]
            n_sp += 1
            if TOOL_DEBUG:
                print(n)
            battle_input_pvp = '[' + ai + ',' + di + ']'
            data = {'memberinfos': battle_input_pvp, 'showDebugFormula': show_debug, 'maxTimes': max_times}
            n = n + 1
            i = 0
            sl_time = 0
            while i < run_time:
                br_row = 11 + i
                r = requests.post(url, data=data)
                soup = BeautifulSoup(r.text, 'html.parser')
                json_data = delete_battle_report_head(soup.get_text(), battle_input_pvp, *HEAD_STRING)
                sht.cells(br_row, 1).value = save_as_json(json_data, custom=str(n))
                if simple_parse:
                    take_simple_parse(sht, simple_parse, json_data, br_row, n_sp, n)
                    n_sp += 1
                i += 1
                # time.sleep(1)
        m += 1
    sht.cells(10, 1).value = '已提交'


def take_parse_title(sht, n_sp):
    sht.range((n_sp + 10, 6)).value = time.strftime('%d%H%M')
    sht.range((n_sp + 10, 13)).value = sht.range("A6:E6").value
    sht.range((n_sp + 10, 21)).value = sht.range("A8:E8").value


def take_simple_parse(sht, level, json_data, br_row, n_sp, n):
    rp = report_parse.BattleReport(json.loads(json_data, strict=False))
    if level >= 1:
        sht.range((n_sp + 10, 6)).value = \
            [n, rp.win, len(rp.rounds), len(rp.heroes_final), sht.cells(br_row, 1).value]
    if level >= 2:
        temp_value = [""]*BATTLE_PARSE_LEN
        for hero_stas in rp.stas:
            col = report_parse.hero_pos_id_trans(hero_stas['mid'], from1101=True, to19=True)
            if (col == 6) or (col == 14):
                col += hero_stas['artifact_pos'] - 1
            temp_value[16 + col] = hero_stas.get('be_hurt') or 0
            temp_value[32 + col] = hero_stas.get('hurt') or 0
            temp_value[48 + col] = hero_stas.get('cure') or 0
            temp_value[col] = temp_value[32 + col] + temp_value[48 + col]
        sht.range((n_sp + 10, 12)).value = temp_value


def delete_battle_report_head(text, *args):
    head_string = ['\n', '战斗测试', '上传新表', '显示公式', '显示被动技能', '显示属性', '显示空值', '严格检查对象池',
                   '防守方替换为MonsterGroup:', '最大回合(默认为0,如果设置了大于0就代表执行到指定回合数才结束战斗):',
                   '最大运行次数(用于检测战斗错误,直到跑出错误才停止):', '运行次数:', 'FunctionID :', '是否打印目标信息']
    for s in args:
        head_string.append(s)
    for hs in head_string:
        text = text.replace(hs, '', 1)
    return text.replace(u'\xa0', '').replace(u'&nbsp', '')


def save_as_json(text, file_name='', path='', custom=''):
    if path == '':
        path = time.strftime('%Y%m%d') + '/'
    path = BATTLE_REPORTS_PATH + path
    if file_name == '':
        file_name = time.strftime('%H%M%S') + '_' + custom + '.json'
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
    shts[SHEET_INFO].range((1, 1)).value = [['模块名称', 'excel_battlefield'], ['当前版本', __version__]]
    if TOOL_DEBUG:
        print('模块名称:{0}\n当前版本:{1}'.format('excel_rw', __version__))
    fresh_contents(shts)
    return shts


def fresh_contents(shts):
    # 根据Sheets-Contents创建索引
    s_key = shts[SHEET_CONTENTS].range('B1:B28')
    s_value = shts[SHEET_CONTENTS].range('C1:C28')
    s_dict = dict(zip(s_key.value, s_value.value))
    global battleInput
    battleInput = s_dict["battleInput"]
    global battleInputAttacker
    battleInputAttacker = s_dict["battleInput_attacker"]
    global battleInputDefenser
    battleInputDefenser = s_dict["battleInput_defenser"]


def main():
    """这个模块通过xlwings直接和EXCEL交互"""
    shts = start_work()
    while True:
        tool_command = input('回车开始战斗')
        if tool_command == 'refresh':
            shts = start_work()
        elif tool_command == 'close':
            break
        elif tool_command == 'clear':
            if os.name == 'posix':
                os.system('clear')
            else:
                os.system('cls')
        else:
            user_command = shts[SHEET_TEST].range('A1:F1')
            uc = user_command.value
            # 0 跑战报 参数1:次数 参数2:需要简析 参数3:战报显示公式 参数4:预留 参数5:怪物组ID
            if uc[0] == 1:
                # 跑战报  怪物组ID为0时进行标准PVP对战,怪物组ID为有效monster_group_id,则进行PVE对战
                push_battle(shts, run_time=int(uc[1]), simple_parse=uc[2], show_debug=int(uc[3]), monster_group=int(uc[5]))

            # 0 跑战报 参数1:次数 参数2:需要简析 参数3:战报显示公式 参数4:快跑起始行
            elif uc[0] == 3:
                # 速刷PVE  怪物组ID为0时进行标准PVP对战,怪物组ID为有效monster_group_id,则进行PVE对战
                rush_pve_battle(shts, max_times=int(uc[1]), simple_parse=uc[2], show_debug=int(uc[3]), start_row=int(uc[4]),
                                try_sl=int(uc[5]))

            # 0 跑战报 参数1:次数 参数2:需要简析 参数3:战报显示公式 参数4:数据起始行号
            elif uc[0] == 4:
                # 速刷PVP
                rush_pvp_battle(shts, max_times=int(uc[1]), simple_parse=uc[2], show_debug=int(uc[3]),
                                start_row=int(uc[4]))
            # 阵容自动创建
            elif uc[0] == 9:
                # 阵容自动创建
                import random_team_creator
                random_team_creator.team_creator(shts, SHEET_TEAM_CREATOR)

            print('完成\n========================')


if __name__ == '__main__':
    main()
