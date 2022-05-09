# -*- coding: UTF-8 -*-
import json
import io


class Dict(dict):
    """将字典解析为对象"""

    __setattr__ = dict.__setitem__
    __getattr__ = dict.__getitem__


def hero_pos_id_trans(pos_id, from16=False, from1101=False, from19=False, to16=False, to1101=False, to19=False):
    """转换英雄pos_id,
    from/to16:我方英雄1~5,神器101~103;敌方英雄6~10,神器104~106
    from/to1101:我方英雄1~5,神器6~8;敌方英雄101~105,神器106~108
    from/to19:我方英雄1~5,神器6~8;敌方英雄9~13,神器14~16
    """
    if to16:
        if from1101:
            if pos_id >= 106:
                pos_id = pos_id - 2
            elif pos_id >= 101:
                pos_id = pos_id - 95
            elif pos_id >= 6:
                pos_id = pos_id + 95
        if from19:
            if pos_id >= 14:
                pos_id = pos_id + 90
            elif pos_id >= 9:
                pos_id = pos_id - 3
            elif pos_id >= 6:
                pos_id = pos_id + 95
    elif to1101:
        if from16:
            if pos_id >= 106:
                pos_id = pos_id + 2
            elif pos_id >= 101:
                pos_id = pos_id - 95
            elif pos_id >= 6:
                pos_id = pos_id + 95
        if from19:
            if pos_id >= 14:
                pos_id = pos_id + 92
            elif pos_id >= 9:
                pos_id = pos_id + 92
    elif to19:
        if from16:
            if pos_id >= 104:
                pos_id = pos_id - 90
            elif pos_id >= 101:
                pos_id = pos_id - 95
            elif pos_id >= 6:
                pos_id = pos_id + 3
        if from1101:
            if pos_id >= 106:
                pos_id = pos_id - 92
            elif pos_id >= 101:
                pos_id = pos_id - 92
    return pos_id


class BattleReport:
    """战报解析"""

    def __init__(self, battle_report_dict):
        self.attackers = """进攻方英雄 list"""
        self.defensers = """防守方英雄 list"""
        self.init = """战斗初始化 obj"""
        self.rounds = """战斗回合 obj"""
        self.stas = """战斗统计 obj"""
        self.win = """胜负情况 bool->int"""
        self.attackers_final = """战斗结果 list"""
        self.defensers_final = """战斗结果 list"""
        self.skills = """用过的技能 list"""
        self.buffs = """用过的BUFF list"""
        if type(battle_report_dict) == dict:
            self.__dict__ = self.dict2obj(battle_report_dict)
        elif type(battle_report_dict) == str:
            self.__dict__ = self.dict2obj(json.loads(battle_report_dict, strict=False))
        try:
            if self.win:
                self.win = 1
        except AttributeError:
            self.win = 0
            pass
        if self.win:
            try:
                self.heroes_final = self.attackers_final
            except AttributeError:
                self.heroes_final = []
        else:
            try:
                self.heroes_final = self.defensers_final
            except AttributeError:
                self.heroes_final = []
        # try:
        #     self.heroes_final = self.attackers_final
        #     self.win = True
        # except AttributeError:
        #     pass
        # try:
        #     self.heroes_final = self.defensers_final
        #     self.win = False
        # except AttributeError:
        #     pass

    def dict2obj(self, dict_obj):
        if not isinstance(dict_obj, dict):
            return dict_obj
        d = Dict()
        for k, v in dict_obj.items():
            # if type(v) == list:
            #     new_v = []
            #     for i in v:
            #         i = self.dict2obj(i)
            #         new_v.append(i)
            #      v = new_v
            d[k] = self.dict2obj(v)
        return d


def main(path=""):
    """返回战斗简报"""
    if path:
        rp = BattleReport(json.loads(open(path, "r", encoding='utf-8').read(), strict=False))
        print("是否获胜: ", rp.win)
        print("总回合数: ", len(rp.rounds))
        if rp.win:
            print("进攻方剩余:")
        else:
            print("防守方剩余:")
        for v in rp.heroes_final:
            print(v)
        print()
    else:
        while True:
            __path = input("输入战报路径:\n")
            if __path:
                main(__path)
            else:
                main("battle_output.json")


if __name__ == '__main__':
    main()
