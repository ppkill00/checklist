#-*- coding: utf-8 -*-
#version : python 3.5~
#필요패키지 : xlsxwriter // 
#pip install xlsxwriter
#동일 디렉토리내 data폴더에 점검결과 파일을 점검

import os,re,xlsxwriter,time

workbook = xlsxwriter.Workbook('test.xlsx')
worksheet1 = workbook.add_worksheet('abstract')
worksheet2 = workbook.add_worksheet('detail')
glo_x = 0

def main():
    file_list = get_file_list()
    file_object = []

    for data in file_list :
        file_object.append( get_object(data))

    y = 0

    for i in range(1,74):
        worksheet1.write(i,0,"U-"+str(i))
    for content in file_object:
        parse_abstract(content,y)
        parse_detail(content)
        y = y+1
    worksheet2.set_column(4,4,50)
    workbook.close()

def get_object(data):
    # print(data) #데이터 파일명
    objects = open('data/'+data,'r',errors='ignore', encoding='utf8') #인코딩 문제있어 errors추가
    contents = data+ '\n\n'
    print (objects)
    for line in objects.readlines():
        contents += line
    return contents #파일 컨텐츠를 전달한다. (파일내용전달)


def get_file_list():
    path_dir = 'data'
    file_list = os.listdir(path_dir)
    file_list.sort()
    return file_list

def parse_detail(contents):
    filename = re.findall(r'(.*?)\.txt',contents,re.MULTILINE)
    os,hostname,ip = filename[0].split("@@")
    result_1 =re.findall(r'=\n\[(U-\d*)\](.*?)\n=.*?\[\d*-START\]\n(.*?)\n\[\d*-END\]\n\n\[U-\d*\]Result : (.*?)\n',contents,re.MULTILINE|re.DOTALL)
    global glo_x
    print(glo_x)

    for data in result_1:
        worksheet2.write(glo_x,0,os)
        worksheet2.write(glo_x,1,ip)
        worksheet2.write(glo_x,2,hostname)
        worksheet2.write(glo_x,3,data[0])
        worksheet2.write(glo_x,4,data[1])
        worksheet2.write(glo_x,5,data[3])
        worksheet2.write(glo_x,6,data[2])
        worksheet2.set_row(glo_x,15 if data.count('\n')==0 else (data.count('\n')+1)*15)
        glo_x=glo_x+1

def parse_abstract(contents,y):
    x=1
    y=y+1

    filename = re.findall(r'(.*?)\.txt',contents,re.MULTILINE)
    worksheet1.write(0,y,filename[0])
    result_1 =re.findall(r'\[U-\d*\]Result : (.*?)\n',contents,re.MULTILINE|re.DOTALL)
    for data in result_1:
        worksheet1.write(x, y, data.replace('\n',''))
        x=x+1



if __name__=="__main__":
    main()

