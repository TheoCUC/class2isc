# -*- coding: utf-8 -*-
import string
import csv
import datetime
import os
import xlrd



weekindex = {
    '星期一':1,
    '星期二':2,
    '星期三':3,
    '星期四':4,
    '星期五':5,
    '星期六':6,
    '星期七':7,
}
classbegin = {
    '1':'080000',
    '2':'090000',
    '3':'101000',
    '4':'111000',
    '5':'133000',
    '6':'142000',
    '7':'152000',
    '8':'161000',
    '9':'180000',
    '10':'190000',
    '11':'200000',
    '12':'210000',
}
classend = {
    '1':'085000',
    '2':'095000',
    '3':'110000',
    '4':'120000',
    '5':'142000',
    '6':'151000',
    '7':'161000',
    '8':'165000',
    '9':'185000',
    '10':'195000',
    '11':'205000',
    '12':'215000'
}
def get_week_num (star):
    if ((star) in weekindex):
        week_num = weekindex[star]
    else:
        week_num = 0
    return week_num

def jieci2time_begin (jieci):
    if (jieci in classbegin):
        timestr = classbegin[jieci]
    else:
        pass
    return timestr

def jieci2time_end (jieci):
    if (jieci in classend):
        timestr = classend[jieci]
    else:
        pass
    return timestr

def getweek_range (astr):
    str1 = astr.split(',')
    str2 = []
    str3 = []
    time = len(str1)
    for i in range(time) :
        str2.append(str1[i].split('周')[0])
    time = len(str2)
    for i in range (time) :
        str3.append(str2[i].split('-'))
    for i in range (time) :
        for j in range(len(str3[i])) :
            str3[i][j] = int (str3[i][j])
    return str3

def getweek (str_get):
    str3 = getweek_range(str_get)
    week = []
    time = len(str3)
    for i in range (time) :
        start = int(str3[i][0])
        if (len(str3[i]) == 2):
            end = int(str3[i][1])
        else:
            end = start
        for j in range(start,end + 1) :
            week.append(j)
    return week

def read_csv (path):
    class_num = 0;
    class_list = []
    with open(path,encoding='GB2312') as csvfile:
        reader = csv.DictReader(csvfile) # 注意函数是大写
        
        for row in reader: # 读取csv文件中的课程表，用class_list存取一个字典数组。
            class_info = {'课程号':'','课程名':'','上课周次':'','上课星期':'','开始节次':'','结束节次':'','上课教师':'','教室名称':''}
            class_info['课程号'] = row['课程号']
            class_info['课程名'] = row['课程名']
            class_info['上课周次'] = getweek(row['上课周次'])
            class_info['周次范围'] = getweek_range(row['上课周次'])
            class_info['上课星期'] = row['上课星期']
            class_info['开始节次'] = row['开始节次']
            class_info['结束节次'] = row['结束节次']
            class_info['上课教师'] = row['上课教师']
            class_info['教室名称'] = row['教室名称']
            class_num = class_num + 1
            class_list.append(class_info)
    class_info = {'class_num':class_num,'class_list':class_list}
    return class_info

def read_xls (path):
    class_num = 0
    class_list = []
    workbook = xlrd.open_workbook(path)

    worksheet = workbook.sheet_by_index(0)

    for i in range(1,worksheet.nrows):
        class_info = {'课程号':'','课程名':'','上课周次':'','上课星期':'','开始节次':'','结束节次':'','上课教师':'','教室名称':''}
        class_info['课程号'] = worksheet.cell_value(i,0)
        class_info['课程名'] = worksheet.cell_value(i,1)
        class_info['上课周次'] = getweek( worksheet.cell_value(i,5) )
        class_info['周次范围'] = getweek_range( worksheet.cell_value(i,5) )
        class_info['上课星期'] = worksheet.cell_value(i,6)
        class_info['开始节次'] = worksheet.cell_value(i,7)
        class_info['结束节次'] = worksheet.cell_value(i,8)
        class_info['上课教师'] = worksheet.cell_value(i,9)
        class_info['教室名称'] = worksheet.cell_value(i,10)
        class_num = class_num + 1
        class_list.append(class_info)
    class_info = {'class_num':class_num,'class_list':class_list}
    return class_info

def checkdate (date_begin,delta_week,number) :
    d1 = datetime.datetime.strptime(date_begin, '%Y-%m-%d')
    delta = datetime.timedelta(days=(delta_week - 1) * 7 + (number - 1))
    d2 = d1 + delta
    date_str = d2.strftime('%Y-%m-%d')
    # print(date_str)
    date_dir = {'year':int(date_str[0:4]),'month':int(date_str[5:7]),'day':int(date_str[8:10])}
    date_str = d2.strftime('%Y%m%d')
    return date_str




def writeisc(date_begin,class_info,wpath):
    if not os.path.exists(wpath):
        os.makedirs(wpath)
    os.chdir(wpath)
    for i in range(class_info['class_num']):
        
        for j in range(len(class_info['class_list'][i]['周次范围'])) :
            file_name = 'Cal' + str(i) + '_' + str(j) + '.ics'

            with open(file_name,'w') as file:
                if(len(class_info['class_list'][i]['周次范围'][j]) == 1) :
                    count = 1
                else :
                    count = class_info['class_list'][i]['周次范围'][j][1] - class_info['class_list'][i]['周次范围'][j][0] + 1
                
                count_str = str(count)
                begin_date_str = checkdate(date_begin, class_info['class_list'][i]['周次范围'][j][0] , get_week_num(class_info['class_list'][i]['上课星期']) )
                begin_time_str = jieci2time_begin(class_info['class_list'][i]['开始节次'])
                end_time_str = jieci2time_end(class_info['class_list'][i]['结束节次'])

                begin_str = begin_date_str + 'T' + begin_time_str
                end_str = begin_date_str + 'T' + end_time_str
                str_1 = ('BEGIN:VCALENDAR\n'
                        'CALSCALE:GREGORIAN\n'
                        'VERSION:2.0\n'
                        'X-WR-CALNAME:')
                str_2 = class_info['class_list'][i]['课程名'] + '\n'
                str_3 = ('METHOD:PUBLISH\n'
                            'PRODID:-//Apple Inc.//Mac OS X 10.15.2//EN\n'
                            'BEGIN:VTIMEZONE\n'
                            'TZID:Asia/Shanghai\n'
                            'BEGIN:STANDARD\n'
                            'TZOFFSETFROM:+0900\n'
                            'RRULE:FREQ=YEARLY;UNTIL=19910914T170000Z;BYMONTH=9;BYDAY=3SU\n'
                            'DTSTART:19890917T020000\n'
                            'TZNAME:GMT+8\n'
                            'TZOFFSETTO:+0800\n'
                            'END:STANDARD\n'
                            'BEGIN:DAYLIGHT\n'
                            'TZOFFSETFROM:+0800\n'
                            'DTSTART:19910414T020000\n'
                            'TZNAME:GMT+8\n'
                            'TZOFFSETTO:+0900\n'
                            'RDATE:19910414T020000\n'
                            'END:DAYLIGHT\n'
                            'END:VTIMEZONE\n'
                            'BEGIN:VEVENT\n'
                            'TRANSP:OPAQUE\n'
                            'DTEND;TZID=Asia/Shanghai:' + end_str + '\n')
                str_5 = 'RRULE:FREQ=WEEKLY;INTERVAL=1;COUNT=' + count_str + '\n'
                str_6 = ('UID:3972A8DD-681F-4DE1-BACF-2236BE083E' + str(i) + str(j) + '\n'
                            'DTSTAMP:20200131T020229Z\n')
                str_7 = 'LOCATION:' + class_info['class_list'][i]['教室名称'] + '\n'
                str_8 = 'DESCRIPTION:' + class_info['class_list'][i]['上课教师'] + '\n'
                str_9 = ('SEQUENCE:1\n'
                            ' X-APPLE-TRAVEL-ADVISORY-BEHAVIOR:AUTOMATIC\n'
                            'SUMMARY:') + class_info['class_list'][i]['课程名'] + '\n'
                str_10 = ('LAST-MODIFIED:20200131T020226Z\n'
                            'CREATED:20200131T020144Z\n'
                            'DTSTART;TZID=Asia/Shanghai:' + begin_str + '\n'
                            'END:VEVENT\n'
                            'END:VCALENDAR\n')
                str_final = str_1 + str_2 + str_3 + str_5 + str_6 + str_7 + str_8 + str_9 + str_10
                file.writelines(str_final)

