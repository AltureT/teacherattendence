import pandas as pd
import datetime
starttime=datetime.datetime.now()

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
def is_work(array: list) -> bool:
    '''

    :param array: ['请假','请假']
    :return: False
    '''
    if array[0] == array[1] == '请假' or '补卡审批通过' in array :
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
    # 工作时间按3小时（180）分钟计算,晚自习则根据最后打卡时间超过八点计算
    if get_work_time(array[1], array[0]) >= 180 or array[1][0:3]>='20':
        return True

    return False


def get_work_times(data: list, attendance: list, leave: list) -> list:
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
        late=0
        severly=0
        # 一天都是休息，则代表法定节假日,也算老师打卡，除周六外
        if leave.count('休息') == 6 and data[i][-3:] != '星期六':
            c = 2

        # 一日最多两次计数统计
        else:
            for j in range(0, len(attendance[0]), 2):
                if c < 2:
                    if attendance[i][j] != -1 and attendance[i][j+1] != -1:
                        worktime=get_work_time(attendance[i][j],attendance[i][j+1])
                    # 请假,补卡也算正常打卡
                    if is_work(leave[i][j:j + 2]) is not True:
                        c = c + 1
                    elif enough_work_time(attendance[i][j:j + 2]):
                        c = c + 1
                    elif 180-worktime>=30:
                        severly+=1
                    elif 0<180-worktime<30:
                        late+=1
        r=[c,late,severly]
        countresult.append(r)
    return countresult



# season='冬季打卡'
# if season=='冬季打卡':
#     night='20:40'
# else:
#     night='21:00'

name=df.loc[:, '姓名'].values.tolist()
attendance = df.loc[:, ['上班1打卡时间', '下班1打卡时间', '上班2打卡时间', '下班2打卡时间', '上班3打卡时间', '下班3打卡时间']].values.tolist()
attendancedate = df.loc[:, '日期'].values.tolist()
leave = df.loc[:, ['上班1打卡结果', '下班1打卡结果', '上班2打卡结果', '下班2打卡结果', '上班3打卡结果', '下班3打卡结果']].values.tolist()
userid = df.loc[:, 'UserId'].values.tolist()
week = ('星期日', '星期一', '星期二', '星期三', '星期四', '星期五', '星期六')
backschool = '星期日'
# ---------每日次数统计,并添加到总数据------------
# 获取样式[1，0，1]，三个值分别代表每日：合格考勤次数，迟到次数，严重迟到次数
attendancetimes = get_work_times(attendancedate, attendance, leave)

x=[]
y=[]
z=[]
for i in range(len(attendancetimes)):
    x.append(attendancetimes[i][0])
    y.append(attendancetimes[i][1])
    z.append(attendancetimes[i][2])
df['打卡次数'] = x
df['缺30分钟以内次数'] = y
df['缺20分钟以上次数'] = z
# print(result)
print(df)

# ---------每日迟到早退次数统计------------


# ---------每周次数统计------------

print(df.columns)
# 周日作为第一天，周日晚至周五晚，至少 4.5 天弹性坐班
# 晚自习不需要满3小时算一次
# 迟到早退 30 分钟(含)以上者计缺勤半天
# 迟到早退 4 次(含合计)折 合缺勤半天 1 次
# 连续旷工超过 15 个工作日或者一年内累计旷工超过 30 个工作 日的，按规定解除聘用合同
# 未签到签退又未事前履行请假报告手续的视为缺勤
# 每周因法定节假日导致不足5天，休息日视为两次

# summary汇总记录每位教师每周打卡次数
summary = {}

# 缓存单位教师每周打卡次数
normal = [0] * 7
late=[0] * 7
severly=[0] * 7
#
start=0
end=-1
for i in range(len(userid)):
    id = str(userid[i])
    d = attendancedate[i][-3:]

    if id not in summary:
        summary[id] = {}

    normal[week.index(d)] = attendancetimes[i][0]
    late[week.index(d)] = attendancetimes[i][1]
    severly[week.index(d)] = attendancetimes[i][2]
    # 由于数据在下周公布，所以在周六，记录上周情况
    if week.index(d) == 6:
        t=str(attendancedate[start])+' -> '+str(attendancedate[end])
        summary[id]['姓名']=name[i]

        # 早退次数考虑教师存在打卡次数过多问题，正常次数少于9次，才存在早退，优先统计30分钟以内迟到
        x=sum(normal)
        # y=9-x if x<9 and sum(late)>=9-x and sum(late)!=0 else 0
        if x < 9:
            if sum(late) >= 9 - x :
                y=9-x
                z=0
            else:
                y=sum(late)
                z=9-x-y

        summary[id][t] = {'正常打卡次数':x,
                          '30分钟以内早退次数':y,
                          '30分钟以上早退次数':z}
        normal = [0] * 7
        start=i+1
        end=i
    else:
        end+=1


for i in summary.items():
    print(i)

endtime=datetime.datetime.now()
print((endtime-starttime).seconds)