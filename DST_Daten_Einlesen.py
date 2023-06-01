import contextlib
import time
import pandas as pd
import numpy as np
import os
import datetime
import re
import sys

''' 
创建数据输入类，该类主要处理和创建新的文件夹，获取所要导入数据文件的名称，整体的导入数据的框架。
'''


class DataLogin(object):

    def __init__(self, data_path):
        # 存储输入进来的文件夹地址
        self.data_path = data_path
        # 获取和存储输入进来的文件夹名称
        self.excelname = os.path.basename(data_path)

    ''' 
    check_filename 静态函数主要是检测输入进来的文件夹里的所有可执行文件的名称是否符合命名规则，返回真或假。（比如 01012020）
    '''
    @staticmethod
    def check_filename(unter_file):
        # 该函数输入进来的为含有所有文件夹中文件的名称 list（str）
        for file in unter_file:
            # 用 re.match 匹配通过正则表达式的方法来检测是否符合日月年的命名规则，如不符合则返回假。此处注意要去掉后缀。
            if not re.match(r"^[0-3]\d0[1-9]\d{4}$", file.removesuffix('.xlsx')):
                print("Dateiname: {} ist nicht im richtigen Form geschrieben, bitte ändern.\n"
                      "Die Dateiname muss in folgende Form geschrieben werden: TagMonatJahr\n"
                      "z.B. 04042021: Tag:04 Monat:04 Jahr:2021".format(file))
                return False
        return True

    ''' 
       datum_sort_path 函数用来对输入进来的文件名进行按日期排序，输入进来数据必须为有序的。对于DST方法来说，文件必须按日期有序
       的读取。此函数通过把文件名str类型转换成datetime类型，用sorted函数对文件名列表排序。再将datetime类型转回str类型，并把
       列表返回。此方法有些复杂，待以后优化处理。
    '''
    def datum_sort_path(self, unter_file):
        # 创建新列表，将列表中字符串类型转换为datetime类型并存储
        datum_unter_file = [datetime.datetime.strptime(datum.removesuffix('.xlsx'), '%d%m%Y')
                            for datum in unter_file]
        # 将列表中成员排序，并且转换回原始字符串格式。其中日月部分要补零。
        sorted_filename = [str(datum.day).zfill(2) + str(datum.month).zfill(2) + str(datum.year) +
                           '.xlsx' for datum in sorted(datum_unter_file)]
        return sorted_filename

    ''' 
        create_path 函数主要是创建导出excel文件的文件夹，用于设置存储数据excel输出文件的地址并返回该地址。
    '''
    def create_path(self):
        # Achten hier kann man auch ohne Korriegierung mit replace durchführen, das Programm läuft ok.
        # Grund dafür ist unbekannt.
        if not os.path.exists(os.getcwd().replace('\\', '/', os.getcwd().count('\\')) + '/DST-Daten'):
            os.mkdir(os.getcwd().replace('\\', '/', os.getcwd().count('\\')) + '/DST-Daten')
        return os.getcwd().replace('\\', '/', os.getcwd().count('\\')) + '/DST-Daten'

    ''' 
        get_subfile 函数为类中的主要函数，创建读取以及写入数据的框架。
    '''
    def get_subfile(self):
        # 获取文件夹中的文件名
        unter_file = os.listdir(self.data_path)
        # 调用check_filename函数，检查文件名是否符合命名规则
        if self.check_filename(unter_file):
            # 如文件名符合要求，则对其进行按时间排序
            sorted_filename = self.datum_sort_path(unter_file)
        else:
            sys.exit()
        # 创建工作目录，文件夹用来存储输出的excel文件
        excel_path = self.create_path()
        # 创建数据管理类，用于读取数据，将读取数据提取存入并导出到新的excel文件里
        data_manager = DataManager(sorted_filename, excel_path +
                                   '/' + self.excelname, self.data_path)
        data_manager.write_data()


''' 
创建数据管理类，该类主要处理数据的输入以及输出。
'''


class DataManager(object):
    def __init__(self, sorted_filename, excel_path, data_path):
        # 存储按时间排序好的文件名列表
        self.sorted_filename = sorted_filename
        # 定义DataFrame中数据列的名称
        self.col_names = ['Datum', 'Zeit', 'Eg', 'Vw', 'Tuk', 'Tus', 'Te_A', 'Ta_A',
                          'Te_B', 'Ta_B', 'V_A', 'Pel_A', 'V_B', 'Pel_B']
        # 存储对列数据进行筛选的函数在字典类型里，此处正好用到fromkeys这个函数。
        self.set_na_and_zero = dict.fromkeys(self.col_names[2:], self.set_na_and_zero)
        # 存储输出文件地址
        self.excel_path = excel_path
        # 存储获取要读取文件的目录地址
        self.data_path = data_path

    ''' 
        set_na_and_zero 静态函数用于筛选填补无效已经遗失数据。
    '''
    @staticmethod
    def set_na_and_zero(cell):
        if cell <= -90:
            return np.NAN
        elif -1 < cell < 0:
            return 0
        return cell

    ''' 
        remove_time 静态函数用于去掉datum列中无效的时间内容，只保留日期部分，返回该日期部分。
    '''
    @staticmethod
    def remove_time(datum_data):
        return datum_data.replace(r'\s\d\d:\d\d:\d\d', '', regex=True)

    ''' 
        get_data 函数用于将读取的数据存入DataFrame类型，返回DataFrame数据类型。
    '''
    def get_data(self, data_path):
        dst_day_data = pd.read_excel(data_path, header=None, skiprows=5,
                                     names=self.col_names, usecols='A:D,G:P',
                                     converters=self.set_na_and_zero)
        # 将Datum列中datetime类型改为str类型
        datum_from_data = dst_day_data['Datum'].astype(str)
        # 调用datum_data函数将列中数据去掉无效时间
        datum_data = self.remove_time(datum_from_data)
        # 有效日期部分与时间部分整合，例如 01-01-2021 00:00:00
        dst_day_data['Zeit'] = datum_data + ' ' + dst_day_data['Zeit'].astype(str)
        # 将Zeit列覆盖，用新整合好的datetime正确时间的列
        dst_day_data['Zeit'] = pd.DataFrame([datetime.datetime.strptime(day, '%Y-%m-%d %H:%M:%S')
                                             for day in dst_day_data['Zeit']])
        # 去除掉无用的Datum列，这里要用到inplace=True直接覆盖掉之前的数据
        dst_day_data.drop('Datum', axis=1, inplace=True)
        # 将Zeit列设置为索引列
        dst_day_data.set_index(['Zeit'], inplace=True)
        return dst_day_data

    ''' 
        write_data 函数用于将数据输出成为一个新的excel文件，此函数为此类的主函数。
    '''
    def write_data(self):
        # 设置起始处理文件数为1
        num_file = 1
        for filename in self.sorted_filename:
            print("{} Datei wird nun eingelesen.".format(filename))
            # 获取整合好的DataFrame数据
            dst_day_data = self.get_data(self.data_path + '/' + filename)
            # 如是读取了第一个文件，则创建excel文件在当前工作文件夹。此处是防止用mode='a'，如果该文件已经存在则会一直添加数据而不是覆盖
            if num_file == 1:
                dst_day_data.to_excel(self.excel_path + '.xlsx', sheet_name="Day{}".format(num_file))
            # 不是第一个文件，则直接通过mode='a'添加数据
            else:
                with pd.ExcelWriter(self.excel_path + '.xlsx', mode='a') as writer:
                    dst_day_data.to_excel(writer, sheet_name="Day{}".format(num_file))
            print("Einlesen von {} ist fertig. Prozess:[{}/{}]".format(filename, num_file, len(self.sorted_filename)))
            num_file += 1


'''
    如下是对输出excel文件程序的升级或者说是改动，对比excel文件读取和接入相对比csv文件以及txt文件的快与慢。通过对类的方法继承来改动，增加新的
    文件格式输出的选项。
'''
'''
    -------------------------------------------CSV-Datei-------------------------------------------------------------
'''


class DataManagerCsv(DataManager):
    def write_data(self):
        num_file = 1
        for filename in self.sorted_filename:
            print("{} Datei wird nun eingelesen.".format(filename))
            # 获取整合好的DataFrame数据
            dst_day_data = self.get_data(self.data_path + '/' + filename)
            # 如是读取了第一个文件，则创建excel文件在当前工作文件夹
            if num_file == 1:
                # 此处需要注意，在德国使用，当作数字的小数点，所以在输出csv文件的时候需要特别标明也就是通过decimal来告知在此使用欧标
                # 在excel文件中导入和导出不存在此问题，若excel与csv文件互相导入则会出现问题
                dst_day_data.to_csv(self.excel_path + '.csv', decimal=',')
            else:
                dst_day_data.to_csv(self.excel_path + '.csv', mode='a', decimal=',')
            print("Einlesen von {} ist fertig. Prozess:[{}/{}]".format(filename, num_file, len(self.sorted_filename)))
            num_file += 1


class DataLoginCsv(DataLogin):
    def get_subfile(self):
        # 获取文件夹中的文件名
        unter_file = os.listdir(self.data_path)
        # 调用check_filename函数，检查文件名是否符合命名规则
        if self.check_filename(unter_file):
            # 如文件名符合要求，则对其进行按时间排序
            sorted_filename = self.datum_sort_path(unter_file)
        else:
            sys.exit()
        # 创建工作目录，文件夹用来存储输出的excel文件
        excel_path = self.create_path()
        # 创建数据管理类，用于读取数据，将读取数据提取存入并导出到新的excel文件里
        data_manager = DataManagerCsv(sorted_filename, excel_path +
                                      '/' + self.excelname, self.data_path)
        data_manager.write_data()


'''
    -------------------------------------------TXT-Datei-------------------------------------------------------------
'''


class DataManagerTxt(DataManager):
    def write_data(self):
        num_file = 1
        for filename in self.sorted_filename:
            print("{} Datei wird nun eingelesen.".format(filename))
            # 获取整合好的DataFrame数据
            dst_day_data = self.get_data(self.data_path + '/' + filename)
            # 如是读取了第一个文件，则创建excel文件在当前工作文件夹
            if num_file == 1:
                # 在txt文件中使用分隔符|来分割数据
                dst_day_data.to_csv(self.excel_path + '.txt', decimal=',', sep='|')
            else:
                dst_day_data.to_csv(self.excel_path + '.txt', mode='a', decimal=',', sep='|')
            print("Einlesen von {} ist fertig. Prozess:[{}/{}]".format(filename, num_file, len(self.sorted_filename)))
            num_file += 1


class DataLoginTxt(DataLogin):
    def get_subfile(self):
        # 获取文件夹中的文件名
        unter_file = os.listdir(self.data_path)
        # 调用check_filename函数，检查文件名是否符合命名规则
        if self.check_filename(unter_file):
            # 如文件名符合要求，则对其进行按时间排序
            sorted_filename = self.datum_sort_path(unter_file)
        else:
            sys.exit()
        # 创建工作目录，文件夹用来存储输出的excel文件
        excel_path = self.create_path()
        # 创建数据管理类，用于读取数据，将读取数据提取存入并导出到新的excel文件里
        data_manager = DataManagerTxt(sorted_filename, excel_path +
                                      '/' + self.excelname, self.data_path)
        data_manager.write_data()


'''
    建立时间管理器来测读取以及写入数据的时间。从结果可以看出写入csv以及txt文件明显要快很多，时间是快了近三倍。但在内存的占用上，很有意思是excel
    占用的最少而csv文件占用近两倍。具体原因不明，猜测是字符串占用了更多的内存，可以以后探讨。
'''


@contextlib.contextmanager
def run_time(program):
    start = time.perf_counter()
    try:
        yield
    finally:
        print("Laufzeit von {} kostet {} sekunden.".format(program, time.perf_counter()-start))


'''
    -------------------------------------------调试主程序-------------------------------------------------------------
'''

with run_time('zu_Excel'):
    data_login = DataLogin("C:/Users/Tan/Desktop/Ausgewählt/Sol_A")
    data_login.get_subfile()

with run_time('zu_Csv'):
    data_login = DataLoginCsv("C:/Users/Tan/Desktop/Ausgewählt/Sol_A")
    data_login.get_subfile()

with run_time('zu_Txt'):
    data_login = DataLoginTxt("C:/Users/Tan/Desktop/Ausgewählt/Sol_A")
    data_login.get_subfile()
