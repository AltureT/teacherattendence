#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
考勤打卡统计程序 - 修复版
解决了DLL加载问题
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
        self.root.title("考勤打卡统计系统 v1.0")
        self.root.geometry("1200x800")

        # 设置样式
        style = ttk.Style()
        style.theme_use('clam')

        # 数据存储
        self.df = None
        self.file_path = None
        self.special_holidays = set()  # 存储特殊休息日
        self.weekly_results = {}  # 存储每周统计结果

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
        main_frame.rowconfigure(3, weight=1)

        # 文件选择区域
        file_frame = ttk.LabelFrame(main_frame, text="文件操作", padding="10")
        file_frame.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)

        ttk.Button(file_frame, text="选择Excel文件", command=self.load_file).grid(row=0, column=0, padx=5)
        self.file_label = ttk.Label(file_frame, text="未选择文件")
        self.file_label.grid(row=0, column=1, padx=5)

        # 特殊日期设置区域
        holiday_frame = ttk.LabelFrame(main_frame, text="特殊休息日设置", padding="10")
        holiday_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)

        ttk.Label(holiday_frame, text="日期 (YYYY-MM-DD):").grid(row=0, column=0, padx=5)
        self.holiday_entry = ttk.Entry(holiday_frame, width=15)
        self.holiday_entry.grid(row=0, column=1, padx=5)

        ttk.Button(holiday_frame, text="添加特殊休息日", command=self.add_holiday).grid(row=0, column=2, padx=5)
        ttk.Button(holiday_frame, text="清空特殊休息日", command=self.clear_holidays).grid(row=0, column=3, padx=5)

        self.holiday_list_label = ttk.Label(holiday_frame, text="已设置的特殊休息日：无")
        self.holiday_list_label.grid(row=1, column=0, columnspan=4, pady=5)

        # 统计控制区域
        control_frame = ttk.LabelFrame(main_frame, text="统计控制", padding="10")
        control_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)

        ttk.Button(control_frame, text="开始统计", command=self.start_analysis).grid(row=0, column=0, padx=5)
        ttk.Button(control_frame, text="导出结果", command=self.export_results).grid(row=0, column=1, padx=5)
        ttk.Button(control_frame, text="清空日志", command=self.clear_log).grid(row=0, column=2, padx=5)

        # 进度条
        self.progress = ttk.Progressbar(control_frame, mode='indeterminate')
        self.progress.grid(row=0, column=3, padx=5, sticky=(tk.W, tk.E))

        # 结果显示区域
        result_frame = ttk.LabelFrame(main_frame, text="统计结果", padding="10")
        result_frame.grid(row=3, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        result_frame.rowconfigure(0, weight=1)
        result_frame.columnconfigure(0, weight=1)

        # 创建树形视图显示结果
        columns = ('姓名', '部门', '周期', '打卡次数', '正常次数', '迟到次数', '旷工次数', '周结果')
        self.tree = ttk.Treeview(result_frame, columns=columns, show='headings', height=15)

        # 设置列标题
        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=100 if col != '姓名' else 120)

        # 添加滚动条
        tree_scroll_y = ttk.Scrollbar(result_frame, orient=tk.VERTICAL, command=self.tree.yview)
        tree_scroll_x = ttk.Scrollbar(result_frame, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.tree.configure(yscrollcommand=tree_scroll_y.set, xscrollcommand=tree_scroll_x.set)

        self.tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        tree_scroll_y.grid(row=0, column=1, sticky=(tk.N, tk.S))
        tree_scroll_x.grid(row=1, column=0, sticky=(tk.W, tk.E))

        # 日志显示区域
        log_frame = ttk.LabelFrame(main_frame, text="运行日志", padding="10")
        log_frame.grid(row=3, column=1, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        log_frame.rowconfigure(0, weight=1)
        log_frame.columnconfigure(0, weight=1)

        self.log_text = scrolledtext.ScrolledText(log_frame, height=20, width=50, wrap=tk.WORD)
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # 统计信息显示
        stats_frame = ttk.LabelFrame(main_frame, text="统计信息", padding="10")
        stats_frame.grid(row=4, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)

        self.stats_label = ttk.Label(stats_frame, text="等待数据加载...")
        self.stats_label.grid(row=0, column=0, padx=5)

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

                # 重命名列（如果需要）
                expected_columns = ["姓名", "考勤组", "部门", "工号", "职位", "UserId", "日期", "workDate",
                                    "班次", "上班1打卡时间", "上班1打卡结果", "下班1打卡时间", "下班1打卡结果"]

                if len(self.df.columns) >= len(expected_columns):
                    column_mapping = {self.df.columns[i]: expected_columns[i]
                                      for i in range(len(expected_columns))}
                    self.df.rename(columns=column_mapping, inplace=True)

                self.file_path = file_path
                self.file_label.config(text=os.path.basename(file_path))

                # 显示基本信息
                self.log(f"文件加载成功！共 {len(self.df)} 条记录")

                # 安全地获取日期范围
                try:
                    date_min = self.df['日期'].min()
                    date_max = self.df['日期'].max()
                    self.log(f"日期范围: {date_min} 至 {date_max}")
                except:
                    self.log("无法获取日期范围")

                # 安全地获取员工人数
                try:
                    unique_employees = self.df['姓名'].nunique()
                    self.log(f"员工人数: {unique_employees}")
                except:
                    self.log("无法统计员工人数")

                # 更新统计信息
                self.update_stats()

            except Exception as e:
                self.log(f"加载文件失败: {str(e)}", 'ERROR')
                messagebox.showerror("错误", f"无法加载文件: {str(e)}")

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

    def analyze_attendance(self, row):
        """分析单条考勤记录"""
        result = {
            'is_normal': False,
            'is_late': False,
            'is_absent': False,
            'details': []
        }

        # 获取打卡结果
        morning_result = str(row.get('上班1打卡结果', '')).strip() if pd.notna(row.get('上班1打卡结果')) else ''
        evening_result = str(row.get('下班1打卡结果', '')).strip() if pd.notna(row.get('下班1打卡结果')) else ''

        # 检查是否是特殊休息日
        date_str = self.parse_date(row.get('日期', ''))
        if date_str in self.special_holidays:
            result['is_normal'] = True
            result['details'].append("特殊休息日")
            return result

        # 检查是否是周末
        if date_str:
            try:
                dt = datetime.strptime(date_str, '%Y-%m-%d')
                if dt.weekday() >= 5:  # 周六或周日
                    result['is_normal'] = True
                    result['details'].append("周末")
                    return result
            except:
                pass

        # 分析上班打卡
        morning_status = self.check_punch_status(morning_result)
        evening_status = self.check_punch_status(evening_result)

        # 判断旷工（至少一次缺卡）
        if morning_status == '缺卡' or evening_status == '缺卡':
            result['is_absent'] = True
            result['details'].append("旷工")
        # 判断迟到（严重迟到 + 正常）
        elif morning_status == '严重迟到' and evening_status == '正常':
            result['is_late'] = True
            result['details'].append("迟到")
        elif morning_status == '正常' and evening_status == '严重迟到':
            result['is_late'] = True
            result['details'].append("迟到")
        # 判断正常
        elif morning_status in ['正常', '补卡', '请假'] and evening_status in ['正常', '补卡', '请假']:
            result['is_normal'] = True
            result['details'].append("正常")
        else:
            # 其他情况
            result['details'].append(f"上班:{morning_status}, 下班:{evening_status}")

        return result

    def check_punch_status(self, status):
        """检查打卡状态"""
        if pd.isna(status) or status == '' or status == '未打卡':
            return '缺卡'

        status = str(status).strip().lower()

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
            return status

    def start_analysis(self):
        """开始统计分析"""
        if self.df is None:
            messagebox.showwarning("警告", "请先加载Excel文件！")
            return

        try:
            self.log("开始统计分析...")
            self.progress.start()

            # 清空之前的结果
            for item in self.tree.get_children():
                self.tree.delete(item)

            # 按员工和周分组统计
            employee_weekly_stats = defaultdict(lambda: defaultdict(lambda: {
                'normal_count': 0,
                'late_count': 0,
                'absent_count': 0,
                'total_punches': 0,
                'days': [],
                'department': '',
                'name': ''
            }))

            total_rows = len(self.df)
            processed = 0

            for idx, row in self.df.iterrows():
                processed += 1
                if processed % 100 == 0:
                    self.log(f"处理进度: {processed}/{total_rows}")

                name = row.get('姓名', '')
                department = row.get('部门', '')
                date_str = self.parse_date(row.get('日期', ''))

                if not date_str or pd.isna(name) or name == '姓名':
                    continue

                week_key = self.get_week_key(date_str)
                if not week_key:
                    continue

                # 分析考勤
                analysis = self.analyze_attendance(row)

                # 更新统计
                stats = employee_weekly_stats[name][week_key]
                stats['name'] = name
                stats['department'] = department
                stats['days'].append(date_str)

                # 每天算2次打卡
                stats['total_punches'] += 2

                if analysis['is_normal']:
                    stats['normal_count'] += 2
                elif analysis['is_late']:
                    stats['late_count'] += 1
                    stats['normal_count'] += 1  # 迟到算一次正常
                elif analysis['is_absent']:
                    stats['absent_count'] += 1

            # 生成最终结果
            self.log("生成统计结果...")
            result_count = 0

            for name, weekly_data in employee_weekly_stats.items():
                for week_key, stats in weekly_data.items():
                    # 计算周结果
                    week_result = self.calculate_week_result(stats)

                    # 添加到树形视图
                    self.tree.insert('', 'end', values=(
                        name,
                        stats['department'],
                        week_key,
                        stats['total_punches'],
                        stats['normal_count'],
                        stats['late_count'],
                        stats['absent_count'],
                        week_result
                    ))
                    result_count += 1

            self.progress.stop()
            self.log(f"统计完成！共生成 {result_count} 条周统计记录")

            # 更新统计信息
            self.update_final_stats()

        except Exception as e:
            self.progress.stop()
            self.log(f"统计过程出错: {str(e)}", 'ERROR')
            messagebox.showerror("错误", f"统计失败: {str(e)}")

    def calculate_week_result(self, stats):
        """计算周结果"""
        # 如果打卡次数少于10次，缺失的视为正常
        if stats['total_punches'] < 10:
            missing_punches = 10 - stats['total_punches']
            stats['normal_count'] += missing_punches
            stats['total_punches'] = 10

        # 判断周结果
        if stats['normal_count'] >= 10:
            return "正常"
        elif stats['absent_count'] > 0:
            return f"旷工{stats['absent_count']}次"
        elif stats['late_count'] > 0:
            return f"迟到{stats['late_count']}次"
        else:
            return "异常"

    def update_stats(self):
        """更新基本统计信息"""
        if self.df is not None:
            try:
                unique_employees = self.df['姓名'].nunique()
                total_records = len(self.df)
                date_min = str(self.df['日期'].min())
                date_max = str(self.df['日期'].max())
                date_range = f"{date_min} 至 {date_max}"

                stats_text = f"员工数: {unique_employees} | 记录数: {total_records} | 日期范围: {date_range}"
                self.stats_label.config(text=stats_text)
            except:
                self.stats_label.config(text="数据已加载")

    def update_final_stats(self):
        """更新最终统计信息"""
        if self.tree.get_children():
            total = len(self.tree.get_children())
            normal = sum(1 for child in self.tree.get_children()
                         if self.tree.item(child)['values'][7] == '正常')
            abnormal = total - normal

            if total > 0:
                stats_text = f"周统计总数: {total} | 正常: {normal} | 异常: {abnormal} | 正常率: {normal / total * 100:.1f}%"
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
            file_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx"), ("CSV files", "*.csv"), ("All files", "*.*")]
            )

            if file_path:
                # 收集数据
                data = []
                for child in self.tree.get_children():
                    values = self.tree.item(child)['values']
                    data.append(values)

                # 创建DataFrame
                columns = ['姓名', '部门', '周期', '打卡次数', '正常次数', '迟到次数', '旷工次数', '周结果']
                result_df = pd.DataFrame(data, columns=columns)

                # 根据文件扩展名保存
                if file_path.endswith('.csv'):
                    result_df.to_csv(file_path, index=False, encoding='utf-8-sig')
                else:
                    with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                        result_df.to_excel(writer, index=False)

                self.log(f"结果已导出到: {os.path.basename(file_path)}")
                messagebox.showinfo("成功", f"结果已成功导出到:\n{file_path}")

        except Exception as e:
            self.log(f"导出失败: {str(e)}", 'ERROR')
            messagebox.showerror("错误", f"导出失败: {str(e)}")


def main():
    """主函数"""
    try:
        root = tk.Tk()
        app = AttendanceChecker(root)
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