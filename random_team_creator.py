# -*- coding: utf-8 -*-
from array import array
from copy import deepcopy
from random import random
from bisect import bisect
import xlwings as xw

__all__ = ['WeightedRandom', ]


class WeightObj(object):
    """记录对象和其初始化权重"""

    def __init__(self, obj, rand_weight):
        self.weight, self.obj = rand_weight, obj


class WeightedRandom(object):
    def __init__(self):
        """初始化总权重、权重列表以及随机对象
        权重列表：为了节省内存空间使用array结构
        """
        self.__total_weight = 0
        self.__weight = array('d')
        self.__objects = []

    def addObject(self, obj, rand_weight):
        """赋权，加对象"""
        self.__total_weight += rand_weight
        self.__weight.append(self.__total_weight)
        self.__objects.append(WeightObj(obj, rand_weight))

    def getObject(self):
        """随机获得对象"""
        return self.__objects[bisect(self.__weight, random() * self.__total_weight)].obj

    def exclusiveGet(self, get_count):
        """排他获取: 即获得过一次不再获得,
        get_count: 最后获得几个结果，最多不超过加入的总数
        """
        result = []
        temp_total = self.__total_weight
        temp_weight, temp_objs = self.__weight[:], deepcopy(self.__objects)
        for _ in range(get_count):
            get_index = bisect(temp_weight, random() * temp_total)
            try:
                current_obj = temp_objs.pop(get_index)
                temp_weight.pop(get_index)
            except IndexError:
                '''获取数超过存储的数据数，直接返回结果'''
                return result
            result.append(current_obj.obj)
            temp_total -= current_obj.weight
        return result


def team_creator(shts, work_sheet, exclusive=0):
    print('创建队伍中……')
    sht = shts[work_sheet]
    hero_list = [[0 for i in range(7)] for i in range(64)]
    for i in range(64):
        for n in range(6):
            hero_list[i][n] = sht[i + 1, n].value
    team_list = [[0 for i in range(5)] for i in range(9999)]
    front_hero_list = []
    front_hero_choice(front_hero_list, hero_list)
    print(front_hero_list)


def front_hero_choice(front_hero_list, hero_list, front_hero=None, pos=0, exclusive=0):
    if front_hero is None:
        front_hero = [0, 0]
    for hero in hero_list:
        if pos == 0:
            if hero[4] <= 1:
                front_hero[pos] = hero[0]
                front_hero_choice(front_hero_list, hero_list, front_hero, pos+1, exclusive)
        elif pos == 1:
            if hero[4] <= 1:
                if not (exclusive == 1 and hero[0] == front_hero[0]):
                    front_hero[pos] = hero[0]
                    front_hero_list.append(front_hero)


def back_hero_choice(hero_list, back_hero, pos=0, exclusive=0):
    for hero in hero_list:
        hero[6] = 0
    for hero in hero_list:
        if pos == 0:
            if hero[4] <= 1:
                back_hero[pos] = hero[0]
                front_hero_choice(hero_list, back_hero, pos+1, exclusive)



if __name__ == '__main__':
    from collections import defaultdict
    factory = WeightedRandom()
    for str_obj, weight in (('a', 1), ('b', 4), ('c', 5)):
        factory.addObject(str_obj, weight)
    '''检查排他的方法是否正确'''
    print(factory.exclusiveGet(100))
    '''随机10000次，校验分布情况是否符合预期'''
    count_dict = defaultdict(int)
    for _ in range(10000):
        get_obj = factory.getObject()
        count_dict[get_obj] += 1
    print(dict(count_dict))