# encoding: utf-8
"""
@author: Jeremiah
@time: 2018/7/28
"""

from re import match
from datetime import datetime
from tkinter import *
from tkinter.filedialog import askopenfilename
from pandas import DataFrame, isna, merge, read_excel, to_datetime

class ExcelCompare():

    def __init__(self, window_name):
        self.window_name = window_name

    def set_init_window(self):
        self.window_name.title("ExcelCompareTool   by：Jeremiah")
        # set window to the center of screen
        x = (self.window_name.winfo_screenwidth() / 2) - 240
        y = (self.window_name.winfo_screenheight() / 2) - 255
        self.window_name.geometry('480x510+%d+%d' % (x, y))

        # null Label for beauty
        Label(self.window_name).grid(row=0, column=0)
        Label(self.window_name).grid(row=1, column=2)
        Label(self.window_name).grid(row=2, column=0)
        # our excel
        self.button_i = Button(self.window_name, text="我方数据", bg="GhostWhite", relief="solid", command=self.load_file_i)
        self.button_i.grid(row=1, column=1)

        self.entry_i = Entry(self.window_name, width=55, bd=5)
        self.entry_i.grid(row=1, column=3, columnspan=3)
        # their excel
        self.button_w = Button(self.window_name, text="竞品数据", bg="GhostWhite", relief="solid", command=self.load_file_w)
        self.button_w.grid(row=3, column=1)

        self.entry_w = Entry(self.window_name, width=55, bd=5)
        self.entry_w.grid(row=3, column=3, columnspan=3)

        # params
        Label(self.window_name).grid(row=4, column=0)
        # Label(self.window_name).grid(row=5, column=0)

        Label(self.window_name, text="参数").grid(row=6, column=1, columnspan=3)
        Label(self.window_name).grid(row=7, column=0)

        Label(self.window_name, text="主键列").grid(row=8, column=1, sticky='e')
        self.entry_pk = Entry(self.window_name, width=10, bd=2)
        self.entry_pk.grid(row=8, column=3, sticky='w')
        Label(self.window_name, text="哪几列为主键列,逗号隔开,默认第一列").grid(row=9, column=1, columnspan=3, sticky='w')

        Label(self.window_name, text="日期列").grid(row=11, column=1, sticky='e')
        self.entry_date = Entry(self.window_name, width=10, bd=2)
        self.entry_date.grid(row=11, column=3, sticky='w')
        Label(self.window_name, text="哪几列是日期,默认以单元格日期格式").grid(row=12, column=1, columnspan=3, sticky='w')

        Label(self.window_name, text="小数位").grid(row=14, column=1, sticky='e')
        self.entry_round = Entry(self.window_name, width=3, bd=2)
        self.entry_round.grid(row=14, column=3, sticky='w')
        Label(self.window_name, text="保留几位小数,默认4位,不影响整数").grid(row=15, column=1, columnspan=3, sticky='w')

        Label(self.window_name, text="空值表").grid(row=16, column=1, sticky='e')
        self.entry_none = Entry(self.window_name, width=10, bd=2)
        self.entry_none.grid(row=16, column=3, sticky='w')
        Label(self.window_name, text="哪些字符可以判定为空,默认包括 \'--\'").grid(row=17, column=1, columnspan=3, sticky='w')
        Label(self.window_name).grid(row=18, column=0)

        # start button
        self.button_start = Button(self.window_name, text="开始对比", bg="GhostWhite", relief="solid", command=self.start)
        self.button_start.grid(row=19, column=1, columnspan=3)
        Label(self.window_name).grid(row=19, column=0)
        Label(self.window_name).grid(row=20, column=0)
        Label(self.window_name, text="程序如有Bug或有新功能需求").grid(row=21, column=1, columnspan=3, sticky='w')
        Label(self.window_name, text="请联系********@*****.com").grid(row=22, column=1, columnspan=3, sticky='w')

        # info text
        self.info_text = Text(self.window_name, width=32, height=27, bd=2, bg='lightblue')
        self.info_text.grid(row=6, column=5, rowspan=30)
        self.info_text.insert(1.0, 
            '********************************\n'\
            '*Excel数据对比工具  version 1.0*\n'\
            '*双方Excel列一一对应且表头一致 *\n'\
            '********************************\n')

    def load_file_i(self):
        path = askopenfilename()
        self.entry_i.insert(1, path)

    def load_file_w(self):
        path = askopenfilename()
        self.entry_w.insert(1, path)

    def writeLog(self, logMsg):
        now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        self.info_text.insert(END, f'{now}\n{logMsg}\n')

    def start(self):

        def getParam(str, type):
            # boolean type for judging the type of return value
            try:
                param = str.strip(',').split(',')
                if type:
                    # int type
                    return list(map(lambda i: (int)(i)-1, param))
                # str type
                return param
            except Exception as e:
                self.writeLog(f'参数输入错误！\n{repr(e)}')

        def isNone(value):
            # value = (str)(value_ori)
            if isna(value) or value == '' or str(value) in None_LIST:
                return True
            return False

        def compare(x, y):
            flag_x = isNone(x)
            flag_y = isNone(y)
            if flag_x and flag_y:
                return '空'
            elif flag_x and not flag_y:
                return '我方无'
            elif not flag_x and flag_y:
                return '竞品无'
            else:
                if re.match(r'^-?\d+(\.\d+)?$', str(x)) and re.match(r'^-?\d+(\.\d+)?$', str(y)):
                    return 'True' if round(float(x),ROUND) == round(float(y),ROUND) else 'False'
                return 'True' if x == y else 'False'
        
        # get path of excel files
        file_i = r'' + self.entry_i.get()
        file_w = r'' + self.entry_w.get()
        try:
            file_r = re.sub(re.compile(r'[^./]*\.'), 'result.', file_i)
        except:
            self.writeLog('Excel文件地址有误！')
            return None
        if file_i and file_w:
            self.writeLog('-------------初始化-------------\n')
            start_time = datetime.now()
            # primary key list
            PK_LIST = [0] if self.entry_pk.get() == '' else getParam(self.entry_pk.get(), True)
            # date list
            DATE_LIST = [] if self.entry_date.get() == '' else getParam(self.entry_date.get(), True)
            # round
            ROUND = 4 if self.entry_round.get() == '' else int(self.entry_round.get())
            # none list
            None_LIST = ['--']
            if self.entry_none.get() != '':
               None_LIST.extend(getParam(self.entry_none.get(), False))
            
            # get DataFrame
            try:
                data_i = DataFrame(read_excel(file_i))
                data_w = DataFrame(read_excel(file_w))
            except Exception as e:
                self.writeLog(f'Excel文件可能损坏!\n{repr(e)}')
                return None

            # all keys
            key = data_i.columns.values.tolist()
            # primary key
            pk = list(map(lambda i: key[i], PK_LIST))

            # process date cols
            if len(DATE_LIST) > 0:
                for date_index in DATE_LIST:
                    data_i[key[date_index]] = to_datetime(data_i[key[date_index]])
                    data_w[key[date_index]] = to_datetime(data_w[key[date_index]])

            # create result, join type:outer
            result = merge(data_i, data_w, on=pk, how='outer')
            # result columns
            cols = list(result)

            # move the primary key to the front
            self.writeLog('正在重构表格列......\n')
            i = 0
            for num in PK_LIST:
                key = cols[num]
                cols.insert(i, cols.pop(num))
                i += 1

            # reorder result by column
            result_key = result.columns.values.tolist()
            len_pk = len(PK_LIST)
            len_nk = len(result_key)
            len_x = (int)((len_nk - len_pk) / 2)

            l = len_x - 1
            for y in range(len_pk + len_x, len_nk - 1):
                # start by col which end with '_y'
                cols.insert(y - l, cols.pop(y))
                l -= 1

            # insert result col after every 2 cols
            a = 2
            for x in range(len_pk, len(cols), 2):
                pos = x + a
                cols.insert(pos, f'result_{a-1}')
                a += 1

            result = result.ix[:, cols]

            # compare
            self.writeLog('正在对比......\n')
            for i in range(len(result)):
                for j in range(len(cols)):
                    if cols[j].endswith('_x'):
                        value_i = result.iloc[i, j]
                        value_w = result.iloc[i, j+1]
                        value_r = compare(value_i, value_w)
                        result.iloc[i, j+2] = value_r

            try:
                DataFrame(result).to_excel(file_r, index=False)
                time = datetime.now() - start_time
                self.writeLog(f'对比完成！\n总用时：{time}')
            except Exception as e:
                self.writeLog(f'数据输出失败！\n{repr(e)}')

def main():

    gui = Tk()
    _gui = ExcelCompare(gui)
    _gui.set_init_window()
    gui.mainloop()

if __name__ == '__main__':
    main()
