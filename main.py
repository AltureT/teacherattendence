import openpyxl
import pandas as pd


# from app import file_path

class User():
    def __init__(self, file_path):
        self.filename = file_path

        # 文件导入从周日或者周一开始，否则第一周情况会出现问题
        self.df = pd.read_excel(self.filename, sheet_name='每日统计')

        # ---------数据处理------------
        # 重设标题索引，丢弃无用行
        columns = self.df[1:2].values.tolist()
        self.df.columns = columns[0]
        self.df = self.df.drop([0, 1])

        # 重设行索引为userID，建立多层索引
        self.userid = self.df.loc[:, 'UserId'].values.tolist()
        self.username = self.df.loc[:, '姓名'].values.tolist()
        self.array = []
        self.array.append(self.userid)
        self.array.append(self.username)
        tuples = list(zip(*self.array))
        index = pd.MultiIndex.from_tuples(tuples, names=['userid', '姓名'])
        self.df.index = index
        self.df = self.df.fillna(value=-1)
        self.name = self.df.loc[:, '姓名'].values.tolist()
        self.attendance = self.df.loc[:, ['上班1打卡时间', '下班1打卡时间', '上班2打卡时间', '下班2打卡时间', '上班3打卡时间', '下班3打卡时间']].values.tolist()
        self.attendancedate = self.df.loc[:, '日期'].values.tolist()
        self.leave = self.df.loc[:, ['上班1打卡结果', '下班1打卡结果', '上班2打卡结果', '下班2打卡结果', '上班3打卡结果', '下班3打卡结果']].values.tolist()
        self.userid = self.df.loc[:, 'UserId'].values.tolist()
        self.week = ('星期日', '星期一', '星期二', '星期三', '星期四', '星期五', '星期六')
        self.backschool = '星期日'

    # ---------每日次数统计,并添加到总数据------------
    # 获取样式[1，0，1]，三个值分别代表每日：合格考勤次数，迟到次数，严重迟到次数
    # attendancetimes = get_work_times(attendancedate, attendance, leave)

    def every_times_count(self, attendancetimes: list)->None:
        x = []
        y = []
        z = []
        for i in range(len(attendancetimes)):
            x.append(attendancetimes[i][0])
            y.append(attendancetimes[i][1])
            z.append(attendancetimes[i][2])
        self.df['打卡次数'] = x
        self.df['缺30分钟以内次数'] = y
        self.df['缺30分钟以上次数'] = z
        # print(result)
        # print(df)

    # ---------计数功能实现------------
    def is_work(self,array: list) -> bool:
        '''

        :param array: ['请假','请假']
        :return: False
        '''
        if array[0] == array[1] == '请假' or '补卡审批通过' in array:
            return False
        return True

    def get_work_time(self,s1: str, s2: str) -> int:
        hour1 = int(s1[0:2])
        minute1 = int(s1[3:5])
        hour2 = int(s2[0:2])
        minute2 = int(s2[3:5])
        if minute1 < minute2:
            minute1 += 60
            hour1 -= 1
        worktime = (hour1 - hour2) * 60 + minute1 - minute2
        return worktime

    def enough_work_time(self, array: list) -> bool:
        '''

        :param array: [06:52,11:45]
        :return: True
        '''

        if array[0] == -1 or array[1] == -1:
            return False
        # 工作时间按3小时（180）分钟计算,晚自习则根据最后打卡时间超过八点计算
        if self.get_work_time(array[1], array[0]) >= 180 or array[1][0:3] >= '20':
            return True

        return False

    # def get_work_times(self,data: list, attendance: list, leave: list) -> list:
    def get_work_times(self) -> list:
        '''

        :param data: 21-09-01 星期三
        :param attendance: [06:52,11:45,13:52,17:14,0,0]
        :param leave:['请假','请假','请假','正常','缺卡','缺卡']
        :return:
        '''

        countresult = []
        for i in range(len(self.attendance)):
            c = 0
            late = 0
            severly = 0
            # 一天都是休息，则代表法定节假日,也算老师打卡，除周六外
            if self.leave.count('休息') == 6 and self.attendancedate[i][-3:] != '星期六':
                c = 2

            # 一日最多两次计数统计
            else:
                for j in range(0, len(self.attendance[0]), 2):
                    if c < 2:
                        if self.attendance[i][j] != -1 and self.attendance[i][j + 1] != -1:
                            worktime = self.get_work_time(self.attendance[i][j], self.attendance[i][j + 1])
                        # 请假,补卡也算正常打卡
                        if self.is_work(self.leave[i][j:j + 2]) is not True:
                            c = c + 1
                        elif self.enough_work_time(self.attendance[i][j:j + 2]):
                            c = c + 1
                        elif 180 - worktime >= 30:
                            severly += 1
                        elif 0 < 180 - worktime < 30:
                            late += 1
            r = [c, late, severly]
            countresult.append(r)
        return countresult


    # ---------每周次数统计------------

    def create_times_list(self,attendancetimes:list)->dict:
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
        late = [0] * 7
        severly = [0] * 7
        #
        start = 0
        end = -1
        for i in range(len(self.userid)):
            id = str(self.userid[i])
            d = self.attendancedate[i][-3:]

            if id not in summary:
                # summary[id] = {}
                summary[id] = []

            normal[self.week.index(d)] = attendancetimes[i][0]
            late[self.week.index(d)] = attendancetimes[i][1]
            severly[self.week.index(d)] = attendancetimes[i][2]
            # 由于数据在下周公布，所以在周六，记录上周情况
            if self.week.index(d) == 6:
                t = str(self.attendancedate[start]) + ' -> ' + str(self.attendancedate[end])

                # 早退次数考虑教师存在打卡次数过多问题，正常次数少于9次，才存在早退，优先统计30分钟以内迟到
                x = sum(normal)
                y=0
                z=0
                # 如果遇到表格从周三等情况开始，则默认周一、二、三打卡正常
                if x < 9:
                    if sum(late) >= 9 - x:
                        y = 9 - x
                        z = 0
                    else:
                        y = sum(late)
                        z = 9 - x - y

                temp = [self.name[i], t, x, y, z]
                # summary[id][t] = {'正常打卡次数':x,
                #                   '30分钟以内早退次数':y,
                #                   '30分钟以上早退次数':z}
                summary[id].append(temp)
                normal = [0] * 7
                start = i + 1
                end = i
            else:
                end += 1
        return summary

    # 字典转化成列表
    # for key,value in summary.items():
    #     for i in range(len(value)):

    # 写入excel文件
    def write_excel(self, path: str, summary: dict):
        title = ['姓名', '日期', '正常打卡次数', '缺30分钟以内次数', '缺30分钟以上次数']

        workbook = openpyxl.load_workbook(self.filename)
        sheet = workbook.create_sheet('按周统计汇总')

        for i in range(5):
            sheet.cell(row=1, column=i + 1, value=title[i])

        i = 1
        for key, val in summary.items():
            for j in range(len(val)):
                for k in range(len(val[j])):
                    sheet.cell(row=i + 1, column=k + 1, value=val[j][k])
                i += 1
        workbook.save(path)
        print('数据成功输出')

# write_excel(filename, summary)
