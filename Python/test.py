# coding=utf-8
# @Author: cfl
# @Time: 2021/1/6 15:05
import xlrd
import json

# 效用
HN = {"H1": -1.0, "H2": -0.8, "H3": -0.6, "H4": -0.4, "H5": -0.2,
      "H6": 0.0,
      "H7": 0.2, "H8": 0.4, "H9": 0.6, "H10": 0.8, "H11": 1.0}

# 属性权重
W = {"e1": 0.15, "e2": 0.12, "e3": 0.15, "e4": 0.1, "e5": 0.08, "e6": 0.08, "e7": 0.08, "e8": 0.1, "e9": 0.14}


def read_xlrd(excel_file, n):
    """
    读取Excel文件
    :param excel_file: excel文件名
    :param n: 两两比较的个数
    :return:
    """
    data = xlrd.open_workbook(excel_file)
    # Sheet1
    table = data.sheet_by_index(0)

    data_lists = []
    for j in range(n):

        data_list = []
        for k in range(table.nrows):
            if k > 0:
                row = table.row_values(k)[j+1]
                data_list.append(json.loads(row))
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


def get_table_3(data_list):
    """
    得到Table3
    :param data_list:
    :return:
    """
    m_n_i = [[0.0 for x in range(9)] for x in range(11)]
    m_h_i = [0.0 for x in range(9)]
    m_h_i_ = [0.0 for x in range(9)]
    m_h_i__ = [0.0 for x in range(9)]
    for q in range(0, 9):
        e_n = 'e' + str(q+1)
        sum_ = 0.0
        for p in range(0, 11):
            h_n = 'H' + str(p+1)
            if h_n not in data_list[q].keys():
                continue
            else:
                m_n_i[p][q] = W[e_n]*data_list[q][h_n]
                sum_ += data_list[q][h_n]
        m_h_i[q] = 1 - W[e_n]*sum_
        m_h_i_[q] = 1 - W[e_n]
        m_h_i__[q] = W[e_n]*(1-sum_)

    m_n = [0.0 for x in range(11)]
    m_h_ = 0.0
    m_h__ = 0.0
    # k
    sum_ = 0.0
    mul_ = [0.0 for x in range(11)]
    for j in range(11):
        mul_1 = 1.0
        for k in range(9):
            mul_1 *= (m_n_i[j][k] + m_h_i_[k] + m_h_i__[k])
        mul_[j] = mul_1
    for j in range(11):
        sum_ += mul_[j]

    mul_2 = 1.0
    for j in range(9):
        mul_2 *= (m_h_i_[j] + m_h_i__[j])
    _k = 1/(sum_ - (11 - 1)*mul_2)

    # m_n
    for j in range(11):
        mul_1 = 1.0
        mul_2 = 1.0
        for k in range(9):
            mul_1 *= m_n_i[j][k] + m_h_i_[k] + m_h_i__[k]
            mul_2 *= m_h_i_[k] + m_h_i__[k]
        m_n[j] = _k*(mul_1 - mul_2)

    # m_h_
    mul_ = 1.0
    for j in range(9):
        mul_ *= m_h_i_[j]
    m_h_ = _k*mul_

    # m_h__
    mul_1 = 1.0
    mul_2 = 1.0
    for j in range(9):
        mul_1 *= m_h_i_[j] + m_h_i__[j]
        mul_2 *= m_h_i_[j]
    m_h__ = _k*(mul_1 - mul_2)

    p_n = [0.0 for x in range(11)]
    for j in range(11):
        p_n[j] = round(m_n[j]/(1-m_h_), 4)

    p_h = round(m_h__/(1-m_h_), 4)

    dict_ = dict()
    for j in range(11):
        h_n = 'H' + str(j + 1)
        dict_[h_n] = p_n[j]
    dict_["H"] = p_h
    return dict_


def g(y, z, x):
    return (y + z - (1 + x) * y * z) / (1 - x * y * z)


def get_table_4(score_list, b):
    table = [[[0.0, 0.0] for x in range(6)] for x in range(6)]
    for j in range(len(score_list) - 1):
        table[j][j+1] = score_list[j]
        table[j+1][j] = [-table[j][j+1][1], -table[j][j+1][0]]

    for j in range(len(score_list) - 2):
        table[j][j+2] = [g(table[j][j+1][0], table[j+1][j+2][0], b), g(table[j][j+1][1], table[j+1][j+2][1], b)]
        table[j+2][j] = [-table[j][j+2][1], -table[j][j+2][0]]
    #
    # for j in range(len(score_list) - 3):
    #     table[j][j+3] = [g(table[j][j+2][0], table[j+1][j+3][0], b), g(table[j][j+2][1], table[j+1][j+3][1], b)]
    #     table[j+3][j] = [-table[j][j+3][1], -table[j][j+3][0]]
    #
    # for j in range(len(score_list) - 4):
    #     table[j][j+4] = [g(table[j][j+3][0], table[j+1][j+4][0], b), g(table[j][j+3][1], table[j+1][j+4][1], b)]
    #     table[j+4][j] = [-table[j][j+4][1], -table[j][j+4][0]]
    #
    # for j in range(len(score_list) - 5):
    #     table[j][j+5] = [g(table[j][j+4][0], table[j+1][j+5][0], b), g(table[j][j+4][1], table[j+1][j+6][1], b)]
    #     table[j+5][j] = [-table[j][j+5][1], -table[j][j+5][0]]
    print(table)


if __name__ == '__main__':
    _excel_file = '../Data/paper_score.xlsx'
    _data_lists = read_xlrd(_excel_file, 5)

    _score_lists = []
    for i in _data_lists:
        # print(i)
        _score_list = get_table_2(i)
        _score_lists.append(_score_list)
    print('table2:')
    print(_score_lists)

    _dict_list = []
    for i in _data_lists:
        _dict = get_table_3(i)
        _dict_list.append(_dict)
    print('table3')
    print(_dict_list)

    _score_list = get_table_2(_dict_list)
    print('table4')
    print(_score_list)

    print('table4')
    get_table_4(_score_list, b=-1.2867)
