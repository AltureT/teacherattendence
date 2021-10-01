import tkinter as tk
import os
from tkinter import filedialog

from main import User


class Application(tk.Frame):
    def __init__(self, master=None):
        tk.Frame.__init__(self, master)
        self.pack()
        self.text1 = tk.Text(self, width=50, height=10, bg='white', font=('Arial', 12))
        self.text1.pack()
        self.createWidgets()

    def createWidgets(self):
        bt1 = tk.Button(self, text='上传文件', width=15, height=2, command=self.upload_file)
        bt1.pack()
        bt2 = tk.Button(self, text='打开文件', width=15, height=2, command=self.open_file)
        bt2.pack()

    def upload_file(self):
        '''
        打开文件
        :return:
        '''
        global file_path
        file_path = filedialog.askopenfilename(title=u'选择文件', initialdir=(os.path.expanduser('~')))
        print('打开文件：', file_path)
        if file_path is not None:
            with open(file=file_path, mode='r+', encoding='utf-8') as file:

                try:
                    data = User(file_path)
                    self.text1.insert('insert', '准备统计中·····\n')
                except:
                    self.text1.insert('insert', '文件导入出错·····\n')

                try:
                    attendancetimes = data.get_work_times()
                    self.text1.insert('insert', '正在统计每日打卡情况·····\n')
                except:
                    self.text1.insert('insert', '数据统计出错，请检查·····\n')

                try:
                    data.every_times_count(attendancetimes)
                    self.text1.insert('insert', '正在各类情况汇总·····\n')
                except:
                    self.text1.insert('insert', '数据汇总出错，请检查·····\n')

                try:
                    summary = data.create_times_list(attendancetimes)
                    self.text1.insert('insert', '正在数据生成中·····\n')
                except:
                    self.text1.insert('insert', '数据生成出错，请检查·····\n')

                try:
                    data.write_excel(file_path, summary)
                    self.text1.insert('insert', '分析完成，打开刚刚下载的 "罗浮中学_每日打卡" 文件·····\n')
                except:
                    self.text1.insert('insert', '文件生成出错，请检查·····\n')

                print('分析完成，请打开刚刚下载的 "罗浮中学_每日打卡" 文件')

    def open_file(self):
        os.system('open ' + file_path)


app = Application()
app.master.title('罗浮中学教师打卡统计程序')
app.mainloop()  # 显示
