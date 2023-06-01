from argparse import ArgumentParser
import DST_Daten_Einlesen as De

'''
    通过交互模式可以随意对想要处理的文件进行处理，但在此代码执行的时候出现了如下问题:
    1.可以看到打印，有可能读取了文件，但是没有新建文件夹
    2.没有输出excel文件
    初步猜测可能与编程环境有关系，需以后解决。
'''
if __name__ == '__main__':
    parser = ArgumentParser(description="DST-Daten-Manager")
    parser.add_argument("-p", "--Pfadname", help="Eingabe von Pfadname")
    args = parser.parse_args()
    filepath = args.Pfadname
    data_login = De.DataLogin(filepath)
    data_login.get_subfile()
