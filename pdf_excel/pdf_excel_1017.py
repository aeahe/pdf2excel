import sys
import os
from openpyxl import Workbook
from binascii import b2a_hex
from pdfminer.pdfparser import PDFParser
from pdfminer.pdfdocument import PDFDocument, PDFNoOutlines
from pdfminer.pdfpage import PDFPage
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import PDFPageAggregator
from pdfminer.layout import LAParams, LTTextBox, LTTextLine, LTFigure, LTImage, LTChar, LTText

path = r'C:\Users\15644\Desktop\pdf_file\test_pdf_list\test_1.pdf'

def parse():
    fp = open(path, 'rb') # 以二进制读模式打开
    #用文件对象来创建一个pdf文档分析器
    praser = PDFParser(fp)
    # 创建一个PDF文档
    doc = PDFDocument(praser)
    # 连接分析器 与文档对象
    praser.set_document(doc)
    # 创建PDf 资源管理器 来管理共享资源
    rsrcmgr = PDFResourceManager()
    # 创建一个PDF设备对象
    laparams = LAParams()
    device = PDFPageAggregator(rsrcmgr, laparams=laparams)
    # 创建一个PDF解释器对象
    interpreter = PDFPageInterpreter(rsrcmgr, device)
    # 循环遍历列表，每次处理一个page的内容

    wb = Workbook() #新建excel
    ws = wb.active

    # 记录page的行数
    text_number = 0

    for page in PDFPage.create_pages(doc): # doc.get_pages() 获取page列表
        interpreter.process_page(page)
        # 接受该页面的LTPage对象
        layout = device.get_result()
        # 这里layout是一个LTPage对象 里面存放着 这个page解析出的各种对象 一般包括LTTextBox, LTFigure, LTImage, LTTextBoxHorizontal 等等 想要获取文本就获得对象的text属性，
        # 得到box
        page_container = [] #存储所有该page的字符串字典
        page_rows = [] #存储行位置数据
        for text_box in layout:
            if (isinstance(text_box, LTTextBox)):
                # 得到line
                for text_line in text_box:
                    if(isinstance(text_line,LTTextLine)):
                        # 得到每个字符
                        temp = [] # 存储得到的字符
                        temp_loc = [] #存储字符串位置
                        isfirst = True #判断是否为字符串的第一个字符
                        for text_index in text_line:
                            # 判断是否为字符数据，并不断更新temp temp_loc
                            if(isinstance(text_index,LTChar)):
                                temp.append(text_index.get_text())
                                if isfirst == True:
                                    temp_loc.append(round(text_index.bbox[0],3))
                                    temp_loc.append(round(text_index.bbox[1],3))
                                    temp_loc.append(round(text_index.bbox[2],3))
                                    temp_loc.append(round(text_index.bbox[3],3))
                                    isfirst = False
                                temp_loc[2] = round(text_index.bbox[2],3)
                                temp_loc[3] = round(text_index.bbox[3],3)
                            # 判断是否为LTText，并将得到的字符串输入page_container的指定位置，最后更新temp 、temp_loc、 isfirst 
                            elif(isinstance(text_index,LTText)):
                                # 如果page_rows没有该行的位置数据，则将数据信息插入page_container，page_rows
                                # if temp_loc[1] not in page_rows:
                                if is_not_in(page_rows,temp_loc[1]):
                                    insert_loc = insert_into_page_rows(page_rows,temp_loc[1])
                                    page_container.insert(insert_loc,[{'value':''.join(temp),'location':temp_loc}])
                                    # page_rows.append(temp_loc[1])
                                    # page_container.append([{'value':''.join(temp),'location':temp_loc}])
                                # 如果有该行的信息
                                elif not is_not_in(page_rows,temp_loc[1]):
                                    # loc = page_rows.index(temp_loc[1])
                                    loc = get_page_rows_loc(page_rows,temp_loc[1])
                                    temp_list = insert_into_page_container(page_container[loc],{'value':''.join(temp),'location':temp_loc})
                                    page_container[loc] = temp_list[:] 
                                temp = []
                                temp_loc = []
                                isfirst = True
        rows_num = len(page_container)
        
        # 对最后一行进行重排 
        if len(page_container[rows_num - 1]) != len(page_container[rows_num - 2]):
            loc_for_no2 = []
            loc_for_no1 = []
            adjust_for_no1 = []
            temp_array = page_container[rows_num - 1][:]
            for i in page_container[rows_num - 2]:
                loc_for_no2.append([i['location'][0],i['location'][2]])
            for i in page_container[rows_num - 1]:
                loc_for_no1.append([i['location'][0],i['location'][2]])
            for i in range(len(loc_for_no1)):
                for j in range(len(loc_for_no2)):
                    if not(loc_for_no1[i][0] > loc_for_no2[j][1] or loc_for_no1[i][1] < loc_for_no2[j][0]):
                        adjust_for_no1.append(j)
                        break
            
            page_container[rows_num - 1] = []
            for i in range(len(page_container[rows_num - 2])):
                if i in adjust_for_no1:
                    page_container[rows_num - 1].append(temp_array[adjust_for_no1.index(i)])
                else:
                    page_container[rows_num - 1].append(None)

        # 对前五行进行重排
        if len(page_container[0]) != len(page_container[1]) or len(page_container[1]) != len(page_container[2]) or len(page_container[2]) != len(page_container[3]) or len(page_container[3]) != len(page_container[4]):
            rows_length = []
            the_max_row = []
            new_max_row = []
            for i in range(6):
                rows_length.append(len(page_container[i]))
            max_length = max(rows_length)
            the_max_row = page_container[rows_length.index(max_length)][:]
            for i in range(len(rows_length)):
                if rows_length[i] < max_length:
                    page_container[i] = align_row(the_max_row,page_container[i])
        # 检测表头

        

        # 输出验证
        for i in range(len(page_container)):
            for j in range(len(page_container[i])):
                print(page_container[i][j])
        # print(page_container)
        # print(page_rows)

        # 得到该页数据以后写入excel
        for i in range(len(page_container)):
            for j in range(len(page_container[i])):
                cell_index = ws.cell(row = i + 1 + text_number, column = j + 1 )
                if page_container[i][j] == None:
                    cell_index.value = ' '
                else:
                    cell_index.value = page_container[i][j]['value']
        
        # 更新text_number，保证page之间的数据连续
        text_number += rows_num
    
    wb.save(r'C:\Users\15644\Desktop\pdf_file\test_pdf_list\test_1.xlsx')

# 将两行对齐
def align_row(row_list_max,row_list_min):
    loc_for_max = []
    loc_for_min = []
    adjust_for_min = []
    temp_array = row_list_min[:]
    for i in row_list_max:
        loc_for_max.append([i['location'][0],i['location'][2]])
    for i in row_list_min:
        loc_for_min.append([i['location'][0],i['location'][2]])
    for i in range(len(loc_for_min)):
        for j in range(len(loc_for_max)):
            if not(loc_for_min[i][0] > loc_for_max[j][1] or loc_for_min[i][1] < loc_for_max[j][0]):
                adjust_for_min.append(j)
                break
    
    row_list_min = []
    for i in range(len(row_list_max)):
        if i in adjust_for_min:
            row_list_min.append(temp_array[adjust_for_min.index(i)])
        else:
            row_list_min.append(None)
    
    return row_list_min

# 将表头行进行合并？？？？？？
def insert_row(row_list_max,row_list_min):
    loc_for_max = []
    loc_for_min = []
    adjust_for_min = []
    temp_array = row_list_min[:]
    for i in row_list_max:
        loc_for_max.append([i['location'][0],i['location'][2]])
    for i in row_list_min:
        loc_for_min.append([i['location'][0],i['location'][2]])
    for i in range(len(loc_for_min)):
        for j in range(len(loc_for_max)):
            if not(loc_for_min[i][0] > loc_for_max[j][1] or loc_for_min[i][1] < loc_for_no2[j][0]):
                adjust_for_min.append(j)
                break
    
    row_list_min = []
    for i in range(len(row_list_max)):
        if i in adjust_for_min:
            row_list_min.append(temp_array[adjust_for_min.index(i)])
        else:
            row_list_min.append(None)
    
    return row_list_min


# 判断行位置坐标是否在page_rows中，差值为3
def is_not_in(row_list,num):
    flag = True
    for i in range(len(row_list)):
        if abs(num - row_list[i]) < 3:
            flag = False
            break
    return flag

# 得到page_rows中元素的坐标
def get_page_rows_loc(row_list,num):
    res = 0
    for i in range(len(row_list)):
        if abs(num - row_list[i]) < 3:
            res = i
            break
    return res

                                
# 将得到字符串插入page_container二维矩阵中                            
def insert_into_page_container(row_list,dic):
    if row_list == []:
        row_list.append([dic])
    else:        
        for i in range(len(row_list)):
            if dic['location'][0] < row_list[i]['location'][0]:
                row_list.insert(i,dic)
                break
        if dic['location'][0] > row_list[len(row_list) - 1]['location'][0]:
            row_list.append(dic)
    return row_list

# 将字符串的行数据插入page_rows中
def insert_into_page_rows(row_list,num):
    a = 0
    if row_list == []:
        row_list.append(num)
    else:
        for i in range(len(row_list)):
            if num > row_list[i]:
                row_list.insert(i,num)
                a = i
                break
        if num < row_list[len(row_list) - 1]:
            row_list.append(num)
            a = len(row_list) - 1
    return a


# 将坐标保留三位整数                             
def round_3(tuple_bbox):
    temp_bbox = []
    for i in range(len(tuple_bbox)):
        temp_bbox.append(round(tuple_bbox[i],3))
    return temp_bbox
    
                    

if __name__ == '__main__':

    parse()
