# coding=utf-8
# @Author: cfl
# @Time: 2021/1/6 21:50
import json
import xlrd
from sympy import *
import math
# 效用
HN = {"H1": -1.0, "H2": -0.8, "H3": -0.6, "H4": -0.4, "H5": -0.2,
      "H6": 0.0,
      "H7": 0.2, "H8": 0.4, "H9": 0.6, "H10": 0.8, "H11": 1.0}


def read_xlrd(excel_file, list_):
    """
    读取Excel文件
    :param excel_file:
    :param list_:
    :return:
    """
    data = xlrd.open_workbook(excel_file)
    table = data.sheet_by_index(0)
    data_lists = []

    for j in list_:
        data_list = []
        for k in range(table.nrows):
            if k > 0:
                row = table.row_values(k)[j]
                data_list.append(json.loads(str(row).replace(' ', '')))
        data_lists.append(data_list)

    print(data_lists)

    return data_lists


def get_table_2(data_list):
    """
    得到Table2
    :param data_list:
    :return:
    """
    score_list = []
    for j in data_list:
        score = 0
        for k in j.keys():
            if k is not "H":
                score += j[k]*HN[k]

        score_range = [0, 0]
        if "H" in list(j.keys()):
            score_range[0] = round(score + j["H"] * HN["H1"], 4)
            score_range[1] = round(score + j["H"] * HN["H11"], 4)
        else:
            score_range[0] = round(score, 4)
            score_range[1] = round(score, 4)

        score_list.append(score_range)

    return score_list


def min_(score_lists):
    s_12__ = []
    for j in score_lists[0]:
        s_12__.append(j[0])
    s_23__ = []
    for j in score_lists[1]:
        s_23__.append(j[0])
    s_13__ = []
    for j in score_lists[2]:
        s_13__.append(j[0])

    s_12_ = []
    for j in score_lists[0]:
        s_12_.append(j[1])
    s_23_ = []
    for j in score_lists[1]:
        s_23_.append(j[1])
    s_13_ = []
    for j in score_lists[2]:
        s_13_.append(j[1])

    x = symbols('x')
    sum_1 = 0.0
    sum_2 = 0.0
    new_g__ = g__(score_lists, x)
    new_g_ = g_(score_lists, x)
    for j in range(len(score_lists)):
        if s_12__[j] * s_23__[j] > 0:
            sum_1 += (new_g__[j] - g(math.fabs(s_12__[j]), math.fabs(s_23__[j]), x)) ** 2
        if s_12_[j] * s_23_[j] > 0:
            sum_2 += (new_g_[j] - g(math.fabs(s_12_[j]), math.fabs(s_23_[j]), x)) ** 2
    y = sum_1 + sum_2

    eq_1 = diff(y, x)
    print(eq_1)
    print(solve(eq_1, x))


def g(y, z, x):
    return (y + z - (1 + x) * y * z) / (1 - x * y * z)


def g__(score_lists, x):
    s_12__ = []
    for j in score_lists[0]:
        s_12__.append(j[0])
    s_23__ = []
    for j in score_lists[1]:
        s_23__.append(j[0])
    s_13__ = []
    for j in score_lists[2]:
        s_13__.append(j[0])

    new_ = []
    for j in range(len(s_13__)):
        if math.fabs(s_13__[j]) >= max(math.fabs(s_12__[j]), math.fabs(s_23__[j])) > 0:
            new_.append(math.fabs(s_13__[j]))
        else:
            new_.append(g(math.fabs(s_12__[j]), math.fabs(s_23__[j]), x))

    return new_


def g_(score_lists, x):
    s_12_ = []
    for j in score_lists[0]:
        s_12_.append(j[1])
    s_23_ = []
    for j in score_lists[1]:
        s_23_.append(j[1])
    s_13_ = []
    for j in score_lists[2]:
        s_13_.append(j[1])

    new_ = []
    for j in range(len(s_13_)):
        if math.fabs(s_13_[j]) >= max(math.fabs(s_12_[j]), math.fabs(s_23_[j])) > 0:
            new_.append(math.fabs(s_13_[j]))
        else:
            new_.append(g(math.fabs(s_12_[j]), math.fabs(s_23_[j]), x))

    return new_


if __name__ == '__main__':
    _excel_file = '../Data/paper_score.xlsx'
    _data_lists = read_xlrd(_excel_file, [4, 5, 6])
    # print(type(_data_lists))
    # print(len(_data_list))
    # print(_data_list)

    _score_lists = []
    for i in _data_lists:
        # print(i)
        _score_list = get_table_2(i)
        _score_lists.append(_score_list)
    # print(_score_lists)

    min_(_score_lists)

