import pandas as pd
from datetime import datetime
import tkinter as tk
from tkinter import filedialog
import os

# 1. 获取用户选择的 xlsx 文件路径
def select_file():
    root = tk.Tk()
    root.withdraw()  # 隐藏主窗口
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.*")])
    return file_path

def generate_report(file_path):
    # 读取上传的 Excel 文件
    df = pd.read_excel(file_path, engine='openpyxl')

    # 查看数据的基本信息
    print(df.head())

    # 1. 问题总数和未解决问题个数
    total_issues = df.shape[0]
    noresolved_issues = df[df['状态'].isin(['Unresolved', '未解决', '解决中', '未开始', '再次复现'])].shape[0]
    print(f"问题总数: {total_issues}")
    print(f"未解决个数: {noresolved_issues}")

    # 1.1 按优先级显示问题数量
    issues_per_priority = df.groupby('优先级').size()
    print("按优先级问题数量:")
    print(issues_per_priority)

    # 2. 分模块问题总数
    issues_per_module = df.groupby('模块').size()
    print("分模块问题总数:")
    print(issues_per_module)

    # 3. 每个人名下的问题个数
    issues_per_person = df.groupby('经办人').size()
    print("每个人名下的问题个数:")
    print(issues_per_person)

    # 创建报告数据
    report_data = (
        f"问题总数: {total_issues}\n"
        f"未解决个数: {noresolved_issues}\n\n"
        f"按优先级问题数量:\n{issues_per_priority}\n\n"
        f"分模块问题总数:\n{issues_per_module}\n\n"
        f"每个人名下的问题个数:\n{issues_per_person}\n"
    )

    # 生成报告文件路径（与原文件同级）
    report_dir = os.path.dirname(file_path)  # 获取原文件所在目录
    report_name = os.path.basename(file_path).split('.')[0] + '_report.txt'  # 添加 "_report" 后缀
    report_path = os.path.join(report_dir, report_name)

    # 将报告数据保存到 txt 文件
    with open(report_path, 'w', encoding='utf-8') as f:
        f.write(report_data)

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
