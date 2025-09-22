#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
考勤打卡统计程序 - 多组别版本
支持教师组、行政组、后勤组不同规则
"""

import sys
import os

# 添加DLL路径修复
if hasattr(sys, '_MEIPASS'):
    os.environ['PATH'] = sys._MEIPASS + os.pathsep + os.environ.get('PATH', '')

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import warnings

warnings.filterwarnings('ignore')

# 延迟导入pandas和numpy，避免DLL问题
try:
    import pandas as pd
    import numpy as np
except ImportError as e:
    import tkinter.messagebox as msgbox

    msgbox.showerror("错误", f"无法加载必要的库: {str(e)}\n请确保已安装pandas和numpy")
    sys.exit(1)

from datetime import datetime, timedelta
from collections import defaultdict


class AttendanceChecker:
    def __init__(self, root):
        self.root = root
        self.root.title("考勤打卡统计系统 v2.0 - 多组别版本")
        self.root.geometry("1300x850")

        # 设置样式
        style = ttk.Style()
        style.theme_use('clam')

        # 数据存储
        self.df = None
        self.file_path = None
        self.special_holidays = set()  # 存储特殊休息日
        self.weekly_results = {}  # 存储每周统计结果

        # 考勤组规则配置
        self.group_rules = {
            '教师组': {
                'daily_punches': 2,  # 每天打卡次数
                'weekly_punches': 10,  # 每周打卡次数
                'punch_columns': ['上班1打卡时间', '上班1打卡结果', '下班1打卡时间', '下班1打卡结果'],
                'description': '每天2次打卡（8:30前上班，16:30后下班）'
            },
            '行政组': {
                'daily_punches': 4,  # 每天打卡次数
                'weekly_punches': 20,  # 每周打卡次数
                'punch_columns': ['上班1打卡时间', '上班1打卡结果', '下班1打卡时间', '下班1打卡结果',
                                  '上班2打卡时间', '上班2打卡结果', '下班2打卡时间', '下班2打卡结果'],
                'description': '每天4次打卡（8:00前、11:20后、13:40前、16:30后）'
            },
            '后勤组': {
                'daily_punches': 4,  # 每天打卡次数
                'weekly_punches': 20,  # 每周打卡次数
                'punch_columns': ['上班1打卡时间', '上班1打卡结果', '下班1打卡时间', '下班1打卡结果',
                                  '上班2打卡时间', '上班2打卡结果', '下班2打卡时间', '下班2打卡结果'],
                'description': '每天4次打卡（8:00前、11:20后、13:40前、16:30后）'
            }
        }

        self.current_group = '教师组'  # 默认选择教师组

        # 创建UI
        self.create_widgets()

    def create_widgets(self):
        """创建UI组件"""
        # 创建主框架
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # 配置权重
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(4, weight=1)

        # 考勤组选择区域
        group_frame = ttk.LabelFrame(main_frame, text="考勤组选择", padding="10")
        group_frame.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)

        ttk.Label(group_frame, text="选择考勤组:").grid(row=0, column=0, padx=5)

        self.group_var = tk.StringVar(value=self.current_group)
        group_combo = ttk.Combobox(group_frame, textvariable=self.group_var,
                                   values=list(self.group_rules.keys()),
                                   state='readonly', width=15)
        group_combo.grid(row=0, column=1, padx=5)
        group_combo.bind('<<ComboboxSelected>>', self.on_group_change)

        self.group_desc_label = ttk.Label(group_frame,
                                          text=self.group_rules[self.current_group]['description'],
                                          foreground='blue')
        self.group_desc_label.grid(row=0, column=2, padx=20)

        # 文件选择区域
        file_frame = ttk.LabelFrame(main_frame, text="文件操作", padding="10")
        file_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)

        ttk.Button(file_frame, text="选择Excel文件", command=self.load_file).grid(row=0, column=0, padx=5)
        self.file_label = ttk.Label(file_frame, text="未选择文件")
        self.file_label.grid(row=0, column=1, padx=5)

        # 特殊日期设置区域
        holiday_frame = ttk.LabelFrame(main_frame, text="特殊休息日设置", padding="10")
        holiday_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)

        ttk.Label(holiday_frame, text="日期 (YYYY-MM-DD):").grid(row=0, column=0, padx=5)
        self.holiday_entry = ttk.Entry(holiday_frame, width=15)
        self.holiday_entry.grid(row=0, column=1, padx=5)

        ttk.Button(holiday_frame, text="添加特殊休息日", command=self.add_holiday).grid(row=0, column=2, padx=5)
        ttk.Button(holiday_frame, text="清空特殊休息日", command=self.clear_holidays).grid(row=0, column=3, padx=5)

        self.holiday_list_label = ttk.Label(holiday_frame, text="已设置的特殊休息日：无")
        self.holiday_list_label.grid(row=1, column=0, columnspan=4, pady=5)

        # 统计控制区域
        control_frame = ttk.LabelFrame(main_frame, text="统计控制", padding="10")
        control_frame.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)

        ttk.Button(control_frame, text="开始统计", command=self.start_analysis,
                   style='Accent.TButton').grid(row=0, column=0, padx=5)
        ttk.Button(control_frame, text="导出结果", command=self.export_results).grid(row=0, column=1, padx=5)
        ttk.Button(control_frame, text="清空日志", command=self.clear_log).grid(row=0, column=2, padx=5)

        # 进度条
        self.progress = ttk.Progressbar(control_frame, mode='indeterminate')
        self.progress.grid(row=0, column=3, padx=5, sticky=(tk.W, tk.E))

        # 当前规则显示
        self.rule_label = ttk.Label(control_frame, text="", foreground='green')
        self.rule_label.grid(row=1, column=0, columnspan=4, pady=5)

        # 结果显示区域
        result_frame = ttk.LabelFrame(main_frame, text="统计结果", padding="10")
        result_frame.grid(row=4, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        result_frame.rowconfigure(0, weight=1)
        result_frame.columnconfigure(0, weight=1)

        # 创建树形视图显示结果
        columns = ('姓名', '部门', '考勤组', '周期', '应打卡', '实际打卡', '正常次数', '迟到次数', '旷工次数', '周结果')
        self.tree = ttk.Treeview(result_frame, columns=columns, show='headings', height=15)

        # 设置列标题和宽度
        column_widths = {
            '姓名': 100, '部门': 120, '考勤组': 80, '周期': 180,
            '应打卡': 70, '实际打卡': 70, '正常次数': 70,
            '迟到次数': 70, '旷工次数': 70, '周结果': 100
        }

        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=column_widths.get(col, 100))

        # 添加滚动条
        tree_scroll_y = ttk.Scrollbar(result_frame, orient=tk.VERTICAL, command=self.tree.yview)
        tree_scroll_x = ttk.Scrollbar(result_frame, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.tree.configure(yscrollcommand=tree_scroll_y.set, xscrollcommand=tree_scroll_x.set)

        self.tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        tree_scroll_y.grid(row=0, column=1, sticky=(tk.N, tk.S))
        tree_scroll_x.grid(row=1, column=0, sticky=(tk.W, tk.E))

        # 日志显示区域
        log_frame = ttk.LabelFrame(main_frame, text="运行日志", padding="10")
        log_frame.grid(row=4, column=1, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        log_frame.rowconfigure(0, weight=1)
        log_frame.columnconfigure(0, weight=1)

        self.log_text = scrolledtext.ScrolledText(log_frame, height=20, width=50, wrap=tk.WORD)
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # 统计信息显示
        stats_frame = ttk.LabelFrame(main_frame, text="统计信息", padding="10")
        stats_frame.grid(row=5, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)

        self.stats_label = ttk.Label(stats_frame, text="等待数据加载...")
        self.stats_label.grid(row=0, column=0, padx=5)

    def on_group_change(self, event=None):
        """考勤组切换事件"""
        self.current_group = self.group_var.get()
        self.group_desc_label.config(text=self.group_rules[self.current_group]['description'])
        self.log(f"已切换到: {self.current_group}")

        # 更新规则显示
        rule = self.group_rules[self.current_group]
        rule_text = f"当前规则: {self.current_group} - 每天{rule['daily_punches']}次, 每周{rule['weekly_punches']}次"
        self.rule_label.config(text=rule_text)

    def log(self, message, level='INFO'):
        """写入日志"""
        timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        log_message = f"[{timestamp}] [{level}] {message}\n"
        self.log_text.insert(tk.END, log_message)
        self.log_text.see(tk.END)
        self.root.update_idletasks()

    def clear_log(self):
        """清空日志"""
        self.log_text.delete(1.0, tk.END)

    def load_file(self):
        """加载Excel文件"""
        file_path = filedialog.askopenfilename(
            title="选择考勤Excel文件",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )

        if file_path:
            try:
                self.log(f"正在加载文件: {os.path.basename(file_path)}")

                # 读取Excel文件，跳过前两行标题
                self.df = pd.read_excel(file_path, skiprows=2, engine='openpyxl')

                # 标准化列名
                self.standardize_columns()

                self.file_path = file_path
                self.file_label.config(text=os.path.basename(file_path))

                # 显示基本信息
                self.log(f"文件加载成功！共 {len(self.df)} 条记录")

                # 检测考勤组信息
                self.detect_attendance_groups()

                # 显示日期范围
                try:
                    date_min = self.df['日期'].min()
                    date_max = self.df['日期'].max()
                    self.log(f"日期范围: {date_min} 至 {date_max}")
                except:
                    self.log("无法获取日期范围", 'WARNING')

                # 更新统计信息
                self.update_stats()

            except Exception as e:
                self.log(f"加载文件失败: {str(e)}", 'ERROR')
                messagebox.showerror("错误", f"无法加载文件: {str(e)}")

    def standardize_columns(self):
        """标准化列名"""
        # 基础列映射
        base_columns = ["姓名", "考勤组", "部门", "工号", "职位", "UserId", "日期", "workDate", "班次"]

        # 教师组列（2次打卡）
        teacher_columns = ["上班1打卡时间", "上班1打卡结果", "下班1打卡时间", "下班1打卡结果"]

        # 行政/后勤组列（4次打卡）
        admin_columns = ["上班1打卡时间", "上班1打卡结果", "下班1打卡时间", "下班1打卡结果",
                         "上班2打卡时间", "上班2打卡结果", "下班2打卡时间", "下班2打卡结果"]

        # 检测是教师组还是行政/后勤组（根据列数）
        if len(self.df.columns) > len(base_columns) + len(teacher_columns):
            # 行政/后勤组格式
            expected_columns = base_columns + admin_columns
        else:
            # 教师组格式
            expected_columns = base_columns + teacher_columns

        # 重命名列
        if len(self.df.columns) >= len(expected_columns):
            column_mapping = {self.df.columns[i]: expected_columns[i]
                              for i in range(len(expected_columns))}
            self.df.rename(columns=column_mapping, inplace=True)
            self.log(f"已标准化 {len(expected_columns)} 个列名")

    def detect_attendance_groups(self):
        """检测文件中的考勤组"""
        try:
            if '考勤组' in self.df.columns:
                groups = self.df['考勤组'].dropna().unique()
                self.log(f"检测到考勤组: {', '.join(groups)}")

                # 统计各组人数
                group_counts = self.df.groupby('考勤组')['姓名'].nunique()
                for group, count in group_counts.items():
                    self.log(f"  {group}: {count} 人")
        except Exception as e:
            self.log(f"检测考勤组失败: {str(e)}", 'WARNING')

    def add_holiday(self):
        """添加特殊休息日"""
        date_str = self.holiday_entry.get().strip()
        if date_str:
            try:
                # 验证日期格式
                datetime.strptime(date_str, '%Y-%m-%d')
                self.special_holidays.add(date_str)
                self.holiday_entry.delete(0, tk.END)
                self.update_holiday_display()
                self.log(f"添加特殊休息日: {date_str}")
            except ValueError:
                messagebox.showerror("错误", "日期格式错误！请使用 YYYY-MM-DD 格式")

    def clear_holidays(self):
        """清空特殊休息日"""
        self.special_holidays.clear()
        self.update_holiday_display()
        self.log("已清空所有特殊休息日")

    def update_holiday_display(self):
        """更新特殊休息日显示"""
        if self.special_holidays:
            holidays_text = "已设置的特殊休息日：" + ", ".join(sorted(self.special_holidays))
        else:
            holidays_text = "已设置的特殊休息日：无"
        self.holiday_list_label.config(text=holidays_text)

    def parse_date(self, date_str):
        """解析日期字符串"""
        if pd.isna(date_str):
            return None

        # 处理各种日期格式
        date_str = str(date_str).strip()

        # 尝试从 "25-09-08 星期一" 格式中提取日期
        if '星期' in date_str:
            date_part = date_str.split()[0]
            parts = date_part.split('-')
            if len(parts) == 3:
                year = '20' + parts[0] if len(parts[0]) == 2 else parts[0]
                return f"{year}-{parts[1].zfill(2)}-{parts[2].zfill(2)}"

        # 尝试其他格式
        for fmt in ['%Y-%m-%d', '%y-%m-%d', '%Y/%m/%d', '%y/%m/%d']:
            try:
                dt = datetime.strptime(date_str.split()[0], fmt)
                return dt.strftime('%Y-%m-%d')
            except:
                continue

        return None

    def get_week_key(self, date_str):
        """获取周标识（周一到周五）"""
        try:
            dt = datetime.strptime(date_str, '%Y-%m-%d')
            # 获取该日期所在周的周一
            monday = dt - timedelta(days=dt.weekday())
            friday = monday + timedelta(days=4)
            return f"{monday.strftime('%Y-%m-%d')} 至 {friday.strftime('%Y-%m-%d')}"
        except:
            return None

    def get_attendance_group(self, row):
        """获取员工的考勤组"""
        # 优先从考勤组列获取
        if '考勤组' in row and pd.notna(row['考勤组']):
            group = str(row['考勤组']).strip()
            # 匹配到已知的组
            for known_group in self.group_rules.keys():
                if known_group in group:
                    return known_group

        # 如果没有考勤组信息，返回当前选择的组
        return self.current_group

    def analyze_attendance(self, row, group_type):
        """分析单条考勤记录"""
        result = {
            'is_normal': False,
            'is_late': False,
            'is_absent': False,
            'normal_count': 0,
            'late_count': 0,
            'absent_count': 0,
            'total_punches': 0,
            'details': []
        }

        # 获取日期
        date_str = self.parse_date(row.get('日期', ''))

        # 检查是否是特殊休息日
        if date_str in self.special_holidays:
            rule = self.group_rules[group_type]
            result['is_normal'] = True
            result['normal_count'] = rule['daily_punches']
            result['total_punches'] = rule['daily_punches']
            result['details'].append("特殊休息日")
            return result

        # 检查是否是周末
        if date_str:
            try:
                dt = datetime.strptime(date_str, '%Y-%m-%d')
                if dt.weekday() >= 5:  # 周六或周日
                    rule = self.group_rules[group_type]
                    result['is_normal'] = True
                    result['normal_count'] = rule['daily_punches']
                    result['total_punches'] = rule['daily_punches']
                    result['details'].append("周末")
                    return result
            except:
                pass

        # 获取该组的打卡规则
        rule = self.group_rules[group_type]

        # 分析每次打卡
        punch_results = []

        if group_type == '教师组':
            # 教师组：2次打卡
            morning_result = self.check_punch_status(row.get('上班1打卡结果', ''))
            evening_result = self.check_punch_status(row.get('下班1打卡结果', ''))
            punch_results = [morning_result, evening_result]

        else:  # 行政组或后勤组
            # 行政/后勤组：4次打卡
            morning1_result = self.check_punch_status(row.get('上班1打卡结果', ''))
            noon1_result = self.check_punch_status(row.get('下班1打卡结果', ''))
            afternoon1_result = self.check_punch_status(row.get('上班2打卡结果', ''))
            evening1_result = self.check_punch_status(row.get('下班2打卡结果', ''))
            punch_results = [morning1_result, noon1_result, afternoon1_result, evening1_result]

        # 统计各类打卡情况
        for punch in punch_results:
            if punch in ['正常', '补卡', '请假']:
                result['normal_count'] += 1
            elif punch == '严重迟到':
                # 严重迟到特殊处理
                pass
            # elif punch == '缺卡':
            #     result['absent_count'] = 1
        if '缺卡' in punch_results[:2]:
            result['absent_count'] += 1
        if '缺卡' in punch_results[2:]:
            result['absent_count'] += 1

        result['total_punches'] = len(punch_results)

        # 判断当天状态
        if '缺卡' in punch_results:
            result['is_absent'] = True
            result['details'].append("旷工")
        elif '严重迟到' in punch_results:
            # 严重迟到+正常 = 迟到
            late_count = punch_results.count('严重迟到')
            normal_count = punch_results.count('正常')
            if late_count > 0 and normal_count > 0:
                result['is_late'] = True
                result['late_count'] = 1
                result['normal_count'] += 1  # 严重迟到算作一次正常
                result['details'].append("迟到")
        elif all(p in ['正常', '补卡', '请假'] for p in punch_results):
            result['is_normal'] = True
            result['details'].append("正常")

        return result

    def check_punch_status(self, status):
        """检查打卡状态"""
        if pd.isna(status) or status == '' or status == '未打卡':
            return '缺卡'

        status = str(status).strip().lower()

        # 包含"正常"字的都视为正常（适用于行政/后勤组规则）
        if '正常' in status or ('管理员' in status and '改为正常' in status):
            return '正常'
        elif '补卡' in status:
            return '补卡'
        elif '请假' in status:
            return '请假'
        elif '严重迟到' in status:
            return '严重迟到'
        elif '缺卡' in status:
            return '缺卡'
        else:
            # 对于行政/后勤组，包含"正常"字的都视为正常
            if self.current_group in ['行政组', '后勤组'] and '正常' in status:
                return '正常'
            return status

    def start_analysis(self):
        """开始统计分析"""
        if self.df is None:
            messagebox.showwarning("警告", "请先加载Excel文件！")
            return

        try:
            self.log(f"开始统计分析 - 当前选择: {self.current_group}")
            self.progress.start()

            # 显示当前规则
            rule = self.group_rules[self.current_group]
            rule_text = f"执行规则: {self.current_group} - 每天{rule['daily_punches']}次, 每周{rule['weekly_punches']}次"
            self.rule_label.config(text=rule_text)

            # 清空之前的结果
            for item in self.tree.get_children():
                self.tree.delete(item)

            # 按员工和周分组统计
            employee_weekly_stats = defaultdict(lambda: defaultdict(lambda: {
                'normal_count': 0,
                'late_count': 0,
                'absent_count': 0,
                'total_punches': 0,
                'expected_punches': 0,
                'days': [],
                'department': '',
                'name': '',
                'group': ''
            }))

            total_rows = len(self.df)
            processed = 0
            error_count = 0

            for idx, row in self.df.iterrows():
                processed += 1
                if processed % 100 == 0:
                    self.log(f"处理进度: {processed}/{total_rows}")

                try:
                    name = row.get('姓名', '')
                    department = row.get('部门', '')
                    date_str = self.parse_date(row.get('日期', ''))

                    if not date_str or pd.isna(name) or name == '姓名':
                        continue

                    # 获取员工的考勤组
                    group_type = self.get_attendance_group(row)

                    # 只统计当前选择的考勤组
                    if group_type != self.current_group:
                        continue

                    week_key = self.get_week_key(date_str)
                    if not week_key:
                        continue

                    # 分析考勤
                    analysis = self.analyze_attendance(row, group_type)

                    # 更新统计
                    stats = employee_weekly_stats[name][week_key]
                    stats['name'] = name
                    stats['department'] = department
                    stats['group'] = group_type
                    stats['days'].append(date_str)

                    # 累加统计
                    stats['total_punches'] += analysis['total_punches']
                    stats['normal_count'] += analysis['normal_count']
                    stats['late_count'] += analysis['late_count']
                    stats['absent_count'] += analysis['absent_count']

                except Exception as e:
                    error_count += 1
                    self.log(f"处理第 {idx + 1} 行时出错: {str(e)}", 'ERROR')
                    continue

            # 生成最终结果
            self.log("生成统计结果...")
            result_count = 0

            for name, weekly_data in employee_weekly_stats.items():
                for week_key, stats in weekly_data.items():
                    # 获取该组的规则
                    group_rule = self.group_rules[stats['group']]
                    expected_weekly = group_rule['weekly_punches']

                    # 计算周结果
                    week_result = self.calculate_week_result(stats, group_rule)

                    # 添加到树形视图
                    self.tree.insert('', 'end', values=(
                        name,
                        stats['department'],
                        stats['group'],
                        week_key,
                        expected_weekly,
                        stats['total_punches'],
                        stats['normal_count'],
                        stats['late_count'],
                        stats['absent_count'],
                        week_result
                    ))
                    result_count += 1

            self.progress.stop()

            if error_count > 0:
                self.log(f"统计完成！共生成 {result_count} 条周统计记录，出错 {error_count} 条", 'WARNING')
            else:
                self.log(f"统计完成！共生成 {result_count} 条周统计记录")

            # 更新统计信息
            self.update_final_stats()

        except Exception as e:
            self.progress.stop()
            self.log(f"统计过程出错: {str(e)}", 'ERROR')
            messagebox.showerror("错误", f"统计失败: {str(e)}")

    def calculate_week_result(self, stats, group_rule):
        """计算周结果"""
        expected_weekly = group_rule['weekly_punches']

        # 如果打卡次数少于预期，缺失的视为正常
        if stats['total_punches'] < expected_weekly:
            missing_punches = expected_weekly - stats['total_punches']
            stats['normal_count'] += missing_punches
            stats['total_punches'] = expected_weekly

        # 判断周结果
        if stats['normal_count'] >= expected_weekly:
            return "正常"
        elif stats['absent_count'] > 0:
            return f"旷工{stats['absent_count']}次"
        elif stats['late_count'] > 0:
            return f"迟到{stats['late_count']}次"
        else:
            abnormal = expected_weekly - stats['normal_count']
            return f"异常{abnormal}次"

    def update_stats(self):
        """更新基本统计信息"""
        if self.df is not None:
            try:
                unique_employees = self.df['姓名'].nunique()
                total_records = len(self.df)

                # 统计各考勤组人数
                group_info = ""
                if '考勤组' in self.df.columns:
                    group_counts = self.df.groupby('考勤组')['姓名'].nunique()
                    group_info = " | ".join([f"{g}:{c}人" for g, c in group_counts.items()])

                date_min = str(self.df['日期'].min())
                date_max = str(self.df['日期'].max())
                date_range = f"{date_min} 至 {date_max}"

                stats_text = f"员工数: {unique_employees} | 记录数: {total_records} | {group_info} | 日期: {date_range}"
                self.stats_label.config(text=stats_text)
            except:
                self.stats_label.config(text="数据已加载")

    def update_final_stats(self):
        """更新最终统计信息"""
        if self.tree.get_children():
            total = len(self.tree.get_children())
            normal = sum(1 for child in self.tree.get_children()
                         if self.tree.item(child)['values'][9] == '正常')
            abnormal = total - normal

            # 统计迟到和旷工
            late = sum(1 for child in self.tree.get_children()
                       if '迟到' in str(self.tree.item(child)['values'][9]))
            absent = sum(1 for child in self.tree.get_children()
                         if '旷工' in str(self.tree.item(child)['values'][9]))

            if total > 0:
                stats_text = (f"{self.current_group} 周统计: 总数{total} | "
                              f"正常{normal} | 迟到{late} | 旷工{absent} | "
                              f"正常率{normal / total * 100:.1f}%")
            else:
                stats_text = "暂无统计数据"
            self.stats_label.config(text=stats_text)

    def export_results(self):
        """导出统计结果"""
        if not self.tree.get_children():
            messagebox.showwarning("警告", "没有可导出的数据！")
            return

        try:
            # 选择保存位置
            default_name = f"考勤统计_{self.current_group}_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
            file_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                initialfile=default_name,
                filetypes=[("Excel files", "*.xlsx"), ("CSV files", "*.csv"), ("All files", "*.*")]
            )

            if file_path:
                # 收集数据
                data = []
                for child in self.tree.get_children():
                    values = self.tree.item(child)['values']
                    data.append(values)

                # 创建DataFrame
                columns = ['姓名', '部门', '考勤组', '周期', '应打卡', '实际打卡',
                           '正常次数', '迟到次数', '旷工次数', '周结果']
                result_df = pd.DataFrame(data, columns=columns)

                # 添加汇总统计
                summary_data = []

                # 统计每个人的总体情况
                for name in result_df['姓名'].unique():
                    person_data = result_df[result_df['姓名'] == name]
                    total_weeks = len(person_data)
                    normal_weeks = sum(person_data['周结果'] == '正常')
                    late_weeks = sum(person_data['周结果'].str.contains('迟到', na=False))
                    absent_weeks = sum(person_data['周结果'].str.contains('旷工', na=False))

                    summary_data.append({
                        '姓名': name,
                        '部门': person_data['部门'].iloc[0],
                        '考勤组': person_data['考勤组'].iloc[0],
                        '统计周数': total_weeks,
                        '正常周数': normal_weeks,
                        '迟到周数': late_weeks,
                        '旷工周数': absent_weeks,
                        '正常率': f"{normal_weeks / total_weeks * 100:.1f}%" if total_weeks > 0 else "0%"
                    })

                summary_df = pd.DataFrame(summary_data)

                # 根据文件扩展名保存
                if file_path.endswith('.csv'):
                    result_df.to_csv(file_path, index=False, encoding='utf-8-sig')
                    # CSV保存汇总到另一个文件
                    summary_path = file_path.replace('.csv', '_汇总.csv')
                    summary_df.to_csv(summary_path, index=False, encoding='utf-8-sig')
                    self.log(f"汇总已导出到: {os.path.basename(summary_path)}")
                else:
                    with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                        result_df.to_excel(writer, sheet_name='周统计明细', index=False)
                        summary_df.to_excel(writer, sheet_name='个人汇总', index=False)

                        # 添加规则说明
                        rules_df = pd.DataFrame([
                            ['考勤组', '每天打卡次数', '每周打卡次数', '规则说明'],
                            ['教师组', '2', '10', '每天2次打卡（8:30前上班，16:30后下班）'],
                            ['行政组', '4', '20', '每天4次打卡（8:00前、11:20后、13:40前、16:30后）'],
                            ['后勤组', '4', '20', '每天4次打卡（8:00前、11:20后、13:40前、16:30后）']
                        ])
                        rules_df.to_excel(writer, sheet_name='考勤规则', index=False, header=False)

                self.log(f"结果已导出到: {os.path.basename(file_path)}")
                messagebox.showinfo("成功", f"结果已成功导出到:\n{file_path}")

        except Exception as e:
            self.log(f"导出失败: {str(e)}", 'ERROR')
            messagebox.showerror("错误", f"导出失败: {str(e)}")


class AboutDialog:
    """关于对话框"""

    def __init__(self, parent):
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("关于")
        self.dialog.geometry("400x300")
        self.dialog.resizable(False, False)

        # 居中显示
        self.dialog.transient(parent)
        self.dialog.grab_set()

        # 添加内容
        frame = ttk.Frame(self.dialog, padding="20")
        frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(frame, text="考勤打卡统计系统", font=('Arial', 16, 'bold')).pack(pady=10)
        ttk.Label(frame, text="版本: v2.0 - 多组别版本").pack(pady=5)
        ttk.Label(frame, text="").pack(pady=5)

        info_text = """支持的考勤组:
• 教师组: 每天2次打卡，每周10次
• 行政组: 每天4次打卡，每周20次  
• 后勤组: 每天4次打卡，每周20次

功能特点:
• 自动识别考勤组
• 支持特殊休息日设置
• 周统计分析
• 导出Excel/CSV报表"""

        text_widget = tk.Text(frame, height=10, width=40, wrap=tk.WORD)
        text_widget.pack(pady=10)
        text_widget.insert('1.0', info_text)
        text_widget.config(state='disabled')

        ttk.Button(frame, text="确定", command=self.dialog.destroy).pack(pady=10)


def add_menu(root, app):
    """添加菜单栏"""
    menubar = tk.Menu(root)
    root.config(menu=menubar)

    # 文件菜单
    file_menu = tk.Menu(menubar, tearoff=0)
    menubar.add_cascade(label="文件", menu=file_menu)
    file_menu.add_command(label="打开文件", command=app.load_file)
    file_menu.add_command(label="导出结果", command=app.export_results)
    file_menu.add_separator()
    file_menu.add_command(label="退出", command=root.quit)

    # 工具菜单
    tools_menu = tk.Menu(menubar, tearoff=0)
    menubar.add_cascade(label="工具", menu=tools_menu)
    tools_menu.add_command(label="清空日志", command=app.clear_log)
    tools_menu.add_command(label="清空特殊休息日", command=app.clear_holidays)

    # 帮助菜单
    help_menu = tk.Menu(menubar, tearoff=0)
    menubar.add_cascade(label="帮助", menu=help_menu)
    help_menu.add_command(label="关于", command=lambda: AboutDialog(root))


def main():
    """主函数"""
    try:
        root = tk.Tk()

        # 设置应用图标（如果有的话）
        try:
            root.iconbitmap(default='icon.ico')
        except:
            pass

        app = AttendanceChecker(root)
        add_menu(root, app)

        # 居中显示窗口
        root.update_idletasks()
        width = root.winfo_width()
        height = root.winfo_height()
        x = (root.winfo_screenwidth() // 2) - (width // 2)
        y = (root.winfo_screenheight() // 2) - (height // 2)
        root.geometry(f'{width}x{height}+{x}+{y}')

        root.mainloop()

    except Exception as e:
        import traceback
        error_msg = f"程序启动失败:\n{str(e)}\n\n详细错误:\n{traceback.format_exc()}"
        print(error_msg)
        if 'root' in locals():
            messagebox.showerror("启动错误", error_msg)
        else:
            import tkinter.messagebox as msgbox
            msgbox.showerror("启动错误", error_msg)


if __name__ == "__main__":
    main()