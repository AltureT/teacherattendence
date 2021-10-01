import pandas as pd
import datetime

df = pd.read_excel('罗浮中学_每日统计_20210901-20210925.xlsx', sheet_name='每日统计')

# ---------数据处理------------
# 重设标题索引，丢弃无用行
columns = df[1:2].values.tolist()
df.columns = columns[0]
df = df.drop([0, 1])

# 重设行索引为userID，建立多层索引
userid = df.loc[:, 'UserId'].values.tolist()
username = df.loc[:, '姓名'].values.tolist()
array = []
array.append(userid)
array.append(username)
tuples = list(zip(*array))
index = pd.MultiIndex.from_tuples(tuples, names=['userid', '姓名'])
df.index = index
df = df.fillna(value=-1)


# print(df)


# ---------计数功能实现------------
def iswork(array: list) -> bool:
    '''

    :param array: ['请假','请假']
    :return: False
    '''
    if array[0] == array[1] == '请假':
        return False
    return True


def get_work_time(s1: str, s2: str) -> int:
    hour1 = int(s1[0:2])
    minute1 = int(s1[3:5])
    hour2 = int(s2[0:2])
    minute2 = int(s2[3:5])
    if minute1 < minute2:
        minute1 += 60
        hour1 -= 1
    worktime = (hour1 - hour2) * 60 + minute1 - minute2
    return worktime


def enough_work_time(array: list) -> bool:
    '''

    :param array: [06:52,11:45]
    :return: True
    '''

    if array[0] == -1 or array[1] == -1:
        return False
    '工作时间按3小时（180）分钟计算'
    if get_work_time(array[1], array[0]) >= 180:
        return True

    return False


def get_work_count(attendance: list, leave: list) -> list:
    '''

    :param userid:  023508212226311234
    :param username: 张三
    :param data: 21-09-01 星期三
    :param attendance: [06:52,11:45,13:52,17:14,0,0]
    :param leave:['请假','请假','请假','正常','缺卡','缺卡']
    :return:
    '''
    countresult = []
    for i in range(len(attendance)):
        c = 0
        for j in range(0, len(attendance[0]), 2):
            if c < 2:
                if iswork(leave[i][j:j + 2]) is not True:
                    c = c + 1
                elif enough_work_time(attendance[i][j:j + 2]):
                    c = c + 1

        countresult.append(c)
    return countresult




attendance = df.loc[:, ['上班1打卡时间', '下班1打卡时间', '上班2打卡时间', '下班2打卡时间', '上班3打卡时间', '下班3打卡时间']].values.tolist()
leave = df.loc[:, ['上班1打卡结果', '下班1打卡结果', '上班2打卡结果', '下班2打卡结果', '上班3打卡结果', '下班3打卡结果']].values.tolist()

# ---------每日次数统计,并添加到总数据------------

result = get_work_count(attendance, leave)
df['每日打卡次数'] = result
# print(result)
print(df)

# ---------每日迟到早退次数统计------------


# ---------每周次数统计------------
attendancedate = df.loc[:, ['日期']].values.tolist()
print(attendancedate)
print(df.columns)
# 周日作为第一天，周日晚至周五晚，至少 4.5 天弹性坐班
# 晚自习不需要满3小时算一次
# 迟到早退 30 分钟(含)以上者计缺勤半天
# 迟到早退 4 次(含合计)折 合缺勤半天 1 次
# 连续旷工超过 15 个工作日或者一年内累计旷工超过 30 个工作 日的，按规定解除聘用合同
# 未签到签退又未事前履行请假报告手续的视为缺勤
# 每周因法定节假日导致不足5天，休息日视为两次
