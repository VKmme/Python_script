#!/usr/bin/python3
# -*- coding: utf-8 -*-
import pymysql
import csv
import traceback
import xlrd
import json

# 将Excel或者csv文件中的数据新增/更新到数据库
# 可使用Pyinstaller打包成EXE使用

# 命令行输入配置参数
def input_type():
    # 接受用户输入参数
    print("本脚本仅支持csv/xlsx/xls文件对数据库的新增或修改（所有文件默认首行为列名；excel默认只操作第一个sheet），请按以下提示进行操作！")
    host = input("请输入数据库IP地址(默认为localhost):")
    if host == '':
        host = 'localhost'
    database = input("请输入数据库名称（默认为cms_zhouxingming）:")
    if database == '':
        database = 'cms_zhouxingming'
    user = input("请输入登陆用户名（默认为root）:")
    if user == '':
        user = 'root'
    password = input("请输入登陆密码（默认为root）:")
    if password == '':
        password = 'root'

    print("请选择操作文件类型：a、csv  b、xlsx或者xls")
    file_type = input("请选择文件类型(默认为a):")
    if file_type == '':
        file_type = 'a'

    print(r"例：C:\Users\user\Desktop\demo.csv")
    path = input("请输入将要执行操作的csv文件的完整路径以及完整文件名（格式如上所示）:")

    print("示例1：update user set cname = '%s'  where id = %s")
    print("示例2：insert into user(cname,active,date_created) values('%s',%s,now)")
    print("提示：%s为占位符，当要插入的数据为varchar或者其他的字符串类型时，需使用''包裹住")
    sql = input("请输入将要执行的sql语句（示例如上）：")

    print("示例1：2,1")
    print("示例2：2,3")
    col_list_str = input("请按照顺序输入sql参数对应的列，并使用英文,分割")

    print("输入参数完毕，程序正在运行，请等待。。。")

    try:
        # 连接数据库
        con = pymysql.Connect(host=host, port=3306, user=user, password=password, database=database,
                              charset='utf8')
        cur = con.cursor()

        col_num = col_list_str.split(',')   #参数对应的列数
        if file_type == 'a':
            # 读取csv文件
            csvFile = open(path, "r")
            reader = csv.reader(csvFile)
            for item in reader:
                if reader.line_num == 1:
                    continue
                params = list()
                for p in range(len(col_num)):
                    params.append(str(item[p]).replace("'", "''"))
                cur.execute(sql % params)

        elif file_type == 'b':
            # 读取Excel
            excel = xlrd.open_workbook(path)
            sheet = excel.sheets()[0]  # 获取第一个sheet
            rows = sheet.nrows  # 行数
            for r in range(1, rows):
                row = sheet.row_values(r)  # 一行数据
                params = list()
                for p in range(len(row)):
                    params.append(str(row[p]).replace("'", "''"))
                cur.execute(sql % params)

        else:
            print("请选择正确的文件类型！")


        con.commit()
        print("执行成功！")
    except:
        print("执行失败,错误日志如下：！")
        traceback.print_exc()
    finally:
        csvFile.close()
        cur.close()
        con.close()
        print("执行完毕！")
        input("按enter键退出")


def json_type():
    try:

        # 接受用户输入参数
        datasource = input("请输入数据源：")

        print("参数读取完毕，程序正在运行，请等待。。。")

        with open('cms.json', 'r') as jsonfile:
            json_string = json.load(jsonfile)
            setting = json_string.get(datasource)

        # 连接数据库
        con = pymysql.Connect(host=setting['host'], port=3306, user=setting['user'], password=setting['password'], database=setting['database'],
                              charset='utf8')
        cur = con.cursor()

        col_num = setting.get('params_col').split(',')  # 参数对应的列数
        print(col_num)
        if setting.get('file_type') == 'a':
            # 读取csv文件
            csv_file = open(setting.get('path'), "r", encoding="utf8")
            reader = csv.reader(csv_file)
            for item in reader:
                if reader.line_num == 1:
                    continue
                params = list()
                for p in col_num:
                    params.append(str(item[int(p)-1]).replace("'", "''"))
                print(setting.get('sql'), params)
                cur.execute(setting.get('sql') % tuple(params))

        elif setting.get('file_type') == 'b':
            # 读取Excel
            excel = xlrd.open_workbook(setting.get('path'))
            sheet = excel.sheets()[0]  # 获取第一个sheet
            rows = sheet.nrows  # 行数
            for r in range(1, rows):
                row = sheet.row_values(r)  # 一行数据
                params = list()
                for p in col_num:
                    params.append(str(row[int(p)-1]).replace("'", "''"))
                print(setting.get('sql'), params)
                cur.execute(setting.get('sql') % tuple(params))

        else:
            print("请选择正确的文件类型！")

        con.commit()
        print("执行成功！")
    except:
        print("执行失败,错误日志如下：！")
        traceback.print_exc()
    finally:
        cur.close()
        con.close()
        print("执行完毕！")
        input("按enter键退出")


if __name__ == '__main__':
    # input_type()
    json_type()
    pass
