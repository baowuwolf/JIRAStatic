import pandas as pd
import tkinter as tk
from tkinter import filedialog
import os
from datetime import datetime

# 1. 获取用户选择的 xlsx 文件路径
def select_file():
    root = tk.Tk()
    root.withdraw()  # 隐藏主窗口
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
    return file_path

# 2. 生成报告并保存
def generate_report(file_path):
    # 读取数据
    df = pd.read_excel(file_path, engine='openpyxl')

    # 计算问题总数和已解决问题个数
    total_issues = df.shape[0]
    resolved_issues = df[df['状态'].isin(['已解决', 'sloved'])].shape[0]

    # 分模块问题总数
    issues_per_module = df.groupby('项目名称').size()

    # 每个人名下的问题个数
    issues_per_person = df.groupby('项目主管').size()

    # 创建报告数据
    report_data = {
        '问题总数': [total_issues],
        '解决个数': [resolved_issues],
        '分模块问题总数': [issues_per_module.to_dict()],
        '每个人名下的问题个数': [issues_per_person.to_dict()],
    }

    report_df = pd.DataFrame(report_data)

    # 生成报告文件路径（与原文件同级）
    report_dir = os.path.dirname(file_path)  # 获取原文件所在目录
    report_name = os.path.basename(file_path).split('.')[0] + '_report.xlsx'  # 添加 "_report" 后缀
    report_path = os.path.join(report_dir, report_name)

    # 保存报告到同级路径
    report_df.to_excel(report_path, index=False, engine='openpyxl')
    print(f"报告已保存到: {report_path}")

# 主程序
if __name__ == '__main__':
    # 选择文件路径
    file_path = select_file()

    if file_path:
        # 生成报告
        generate_report(file_path)
    else:
        print("没有选择文件！")
