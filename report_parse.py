import json


class Dict(dict):
    """将字典解析为对象"""

    __setattr__ = dict.__setitem__
    __getattr__ = dict.__getitem__


class BattleReport:
    """战报解析"""

    def __init__(self, battle_report_dict):
        self.attackers = """进攻方英雄 list"""
        self.defensers = """防守方英雄 list"""
        self.init = """战斗初始化 obj"""
        self.rounds = """战斗回合 obj"""
        self.win = """胜负情况 bool"""
        self.stas = """战斗统计 obj"""
        self.attackers_final = """战斗结果 list"""
        self.skills = """用过的技能 list"""
        self.buffs = """用过的BUFF list"""
        if type(battle_report_dict) == dict:
            self.__dict__ = self.dict2obj(battle_report_dict)
        elif type(battle_report_dict) == str:
            self.__dict__ = self.dict2obj(json.loads(battle_report_dict, strict=False))

    def dict2obj(self, dict_obj):
        if not isinstance(dict_obj, dict):
            return dict_obj
        d = Dict()
        for k, v in dict_obj.items():
            if type(v) == list:
                new_v = []
                for i in v:
                    i = self.dict2obj(i)
                    new_v.append(i)
                v = new_v
            d[k] = self.dict2obj(v)
        return d


def main(path=""):
    """返回战斗简报"""
    if path:
        rp = BattleReport(json.loads(open(path, "r", encoding='utf-8').read(), strict=False))
        print("进攻方是否获胜: ", rp.win)
        print("总回合数: ", len(rp.rounds))
        for v in rp.attackers_final:
            print(v)
        print()
    else:
        while True:
            __path = input("输入战报路径:\n")
            main(__path)


if __name__ == '__main__':
    main()
