# -*- coding: utf-8 -*-
from datetime import datetime

import xlwt as xlwt

from django.shortcuts import render, redirect

# Create your views here.
from rest_framework import status
from rest_framework.decorators import api_view
from rest_framework.response import Response
import os
import xlrd
from xlrd import xldate_as_tuple

from excel.settings import BASE_DIR


# @api_view(['POST', "GET"])
def uploadfile(request):
    myfile = request.FILES.get('fileField')
    if myfile.name.split('.')[-1] != 'xls':
        return Response('请传入.xls格式excel表格')
    data = xlrd.open_workbook(os.path.join('d:/excelFile.xls'))
    if myfile:
        excel_file = open(os.path.join('d:/', myfile.name), 'wb')
        for chuck in myfile.chunks():
            excel_file.write(chuck)
        excel_file.close()
    filename = os.path.join('d:/', myfile.name)

    sheet_name = '考勤表'
    sheet_first_rows = [[u'姓名', u'上班时间', u'下班时间', u'班次',u'备注']]
    colnameindex = 0  # 表头占的行数默认只有首行
    data, flag = read_table(filename=filename, colnameindex=0, by_name=sheet_name, sheet_first_rows=sheet_first_rows)
    if flag:
        return Response(flag, status=status.HTTP_200_OK)
    else:

        m = 0  # 统计异常次数
        n = 0  # 统计未打卡次数
        date_list=[]
        for info in data:
            flag = False
            Shifts_time = info[3].split('-')
            if info[1]:
                work = datetime.strptime(info[1][:11] + Shifts_time[0], '%Y/%m/%d %H:%M')
                if datetime.strptime(info[1], '%Y/%m/%d %H:%M:%S') > work:
                    info[4]='迟到'
                    m += 1
            else:
                info.append('未打卡')
                m += 1
                n += 1
            if info[2]:
                off_time = datetime.strptime(info[2][:11] + Shifts_time[1], '%Y/%m/%d %H:%M')
                if datetime.strptime(info[2], '%Y/%m/%d %H:%M:%S') < off_time:
                    if info[4]:
                        info[4] = info[4] + '; 早退'
                        m += 1
                    else:
                        info[4] = '早退'
                        m += 1
            else:
                if info[4] and '未打卡' not in info[4]:
                    info[4] = info[4] + '; 未打卡'
                    m += 1
                    n += 1
                elif info[4] and '未打卡' in info[4]:
                    m += 1
                    n += 1
                elif not info[4]:
                    info[4]='未打卡'
                    m += 1
                    n += 1

            date_list.append(info)
        if m > 0:
            err = '注意：考勤数据截至此时，异常{}次，其中未打卡{}次，迟到早退{}次'.format(m, n, m - n)
            date_list.append([err,'','','',''])
            filename=BASE_DIR+'/static/excel/newwork.xls'
            excel_export(date_list, filename, sheet_name, sheet_first_rows)
        return redirect('/static/excel/newwork.xls')


def excel_export(data, filename, sheet_name, sheet_first_rows):
    writebook = xlwt.Workbook()
    sheet = writebook.add_sheet(sheet_name)

    data.insert(0, sheet_first_rows[0])
    rownum=len(sheet_first_rows[0])
    for i in range(len(data)):
        for j in range(rownum):
            sheet.write(i, j, data[i][j])
    writebook.save(filename)
    print(filename)

def read_table(filename=None, colnameindex=None, by_name=None, sheet_first_rows=None):
    filedata = xlrd.open_workbook(filename)
    table = filedata.sheet_by_name(by_name)

    if not table:
        flag = '请检查excel表格的sheet名称是否为一致'
        return [], flag
    nrows = table.nrows  # 行数
    colnames = table.row_values(colnameindex)  # 第一行表头数据
    if sheet_first_rows[0] != colnames:
        flag = '请检查excel表格的表头是否为一致'
        return [], flag
    flag, all_content = None, []

    for rownum in range(1, nrows):
        row_content = []
        for col_value in range(len(colnames)):
            ctype = table.cell(rownum, col_value).ctype  # 表格的数据类型
            cell = table.cell_value(rownum, col_value)
            '''ctype： 
            0   empty
            1   string
            2   number
            3   date
            4   boolean
            5   Error'''
            if ctype == 2 and cell % 1 == 0:  # 如果是整形
                cell = int(cell)
            elif ctype == 3:
                # 转成datetime对象
                date = datetime(*xldate_as_tuple(cell, 0))
                cell = date.strftime('%Y/%m/%d %H:%M:%S')
            elif ctype == 4:
                cell = True if cell == 1 else False
            elif ctype == 1:

                cell = str(cell)
            row_content.append(cell)
        all_content.append(row_content)

    return all_content, flag



def index(request):
    return render(request, 'index.html')
