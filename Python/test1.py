# coding=utf-8
# @Author: cfl
# @Time: 2021/1/6 21:50
import json
import xlrd
# 期望
HN = {"H1": -1.0, "H2": -0.8, "H3": -0.6, "H4": -0.4, "H5": -0.2,
      "H6": 0.0,
      "H7": 0.2, "H8": 0.4, "H9": 0.6, "H10": 0.8, "H11": 1.0}

W = {"e1": 0.15, "e2": 0.12, "e3": 0.15, "e4": 0.1, "e5": 0.08, "e6": 0.08, "e7": 0.08, "e8": 0.1, "e9": 0.14}


def read_xlrd(excel_file):
    """
    读取Excel文件
    :param excel_file:
    :return:
    """
    data = xlrd.open_workbook(excel_file)
    table = data.sheet_by_index(0)
    data_lists = []

    for j in range(3):
        data_list = []
        for k in range(table.nrows):
            if k > 0:
                row = table.row_values(k)[j+1]
                data_list.append(json.loads(str(row).replace(' ', '')))
        # print(data_list)
        data_lists.append(data_list)

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


if __name__ == '__main__':
    _excel_file = '../Data/score.xlsx'
    _data_lists = read_xlrd(_excel_file)
    # print(type(_data_lists))
    # print(len(_data_list))
    # print(_data_list)

    for i in _data_lists:
        # print(i)
        _score_list = get_table_2(i)
        print(_score_list)

