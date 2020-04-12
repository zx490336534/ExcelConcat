# -*- encoding: utf-8 -*-
"""
@File    : ExcelConcat.py
@Time    : 2020/4/12 2:01 下午
@Author  : zhongxin
@Email   : 490336534@qq.com
"""
import os
import threading
from tkinter import *
import pandas as pd


def get_df(path, name):
    text.insert(1.0, f'开始读取{path}中sheet为{name}的信息\n')
    df = pd.DataFrame(pd.read_excel(path, sheet_name=name))
    return df


def concat_df(df_list):
    sum = 0
    for i in df_list:
        sum += len(i)
        print(len(i))
    result = pd.concat(df_list, sort=False)
    print(f'理论上合并后条数为{sum},实际为{len(result)}')
    text.insert(1.0, f'理论上合并后条数为{sum},实际为{len(result)}\n')
    return result


def write_into_xls(result, file_name='result.xls'):
    text.insert(1.0, f'开始写入{file_name},请稍等...\n')
    writer = pd.ExcelWriter(file_name)
    result.to_excel(writer, index=False)
    writer.save()
    text.insert(1.0, f'写入{file_name}完成。\n\n')


def work():
    file_path = path.get()
    sheet_list = sheet.get().split('|')
    df_list = []
    p, name = os.path.split(file_path)
    _, n = os.path.splitext(file_path)
    new_path = os.path.join(p, f'(合并后){name.replace(n, ".xlsx")}')
    for i in sheet_list:
        df_list.append(get_df(file_path, i))
    result = concat_df(df_list)
    write_into_xls(result, new_path)


def thread_it(func, *args):
    t = threading.Thread(target=func, args=args)
    t.setDaemon(True)
    t.start()


if __name__ == '__main__':
    top = Tk()
    top.title('Excel合并工具')
    top.geometry('600x400')
    frame = Frame(top)
    frame.pack()
    lab = Label(frame, text='待合并Excel路径:   ')
    lab.grid(row=0, column=0, sticky=W)
    path = Entry(frame)
    path.insert(0, '/Users/zhongxin/PycharmProjects/datawork/ZB-ECRC分销出货查询 4.10-4.11浙皖.xls')
    path.grid(row=0, column=1, sticky=W)

    lab = Label(frame, text='子sheet名称(使用|分割):   ')
    lab.grid(row=1, column=0, sticky=W)
    sheet = Entry(frame)
    sheet.insert(0, '分销明细|分销明细_1')
    sheet.grid(row=1, column=1, sticky=W)

    btn1 = Button(frame, text="开始合并", command=lambda: thread_it(work), width=20)
    btn1.grid(row=1, column=2, sticky=W)

    text = Text(top, width=20, height=100)
    text.pack(fill=X, side=BOTTOM)
    top.mainloop()
